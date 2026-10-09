using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using RibbonXmlEditor.Config;
using RibbonXmlEditor.Macros;
using RibbonXmlEditor.Ribbon;

namespace RibbonXmlEditor.Packaging
{
    /// <summary>
    /// Đọc/ghi customUI14.xml + ảnh icon tùy chỉnh thẳng vào gói OOXML (.xlsm/.xlam) - đúng cơ
    /// chế Office RibbonX Editor dùng, không cần add-in/callback nào chạy nền trong Excel.
    /// Ghi macro wrapper (nếu cần) được xử lý qua Excel COM ở bước riêng (Save), TRƯỚC khi
    /// đụng vào package bằng System.IO.Packaging - 2 bước không được chạy đồng thời vì Excel
    /// giữ file khóa trong lúc mở.
    /// </summary>
    public static class CustomUiPackageService
    {
        private const string ExtensibilityRelType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility";
        private const string ImageRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        private static readonly Uri CustomUiPartUri = new Uri("/customUI/customUI14.xml", UriKind.Relative);

        public class LoadResult
        {
            public RibbonConfig Config { get; set; } = new RibbonConfig();
            public string ExtractedImagesDir { get; set; }
            public string Error { get; set; }
        }

        public static LoadResult Load(string filePath)
        {
            var result = new LoadResult();
            try
            {
                using (var package = Package.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    var rel = package.GetRelationshipsByType(ExtensibilityRelType).FirstOrDefault();
                    if (rel == null)
                        return result; // chưa có customUI trong file - danh sách rỗng là bình thường

                    var partUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), rel.TargetUri);
                    if (!package.PartExists(partUri))
                        return result;

                    var part = package.GetPart(partUri);
                    string xml;
                    using (var reader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
                        xml = reader.ReadToEnd();

                    result.Config = CustomUiXmlParser.Parse(xml, out var imageRelIds);

                    if (imageRelIds.Count > 0)
                    {
                        string extractDir = Path.Combine(Path.GetTempPath(), "RibbonXmlEditor", Guid.NewGuid().ToString("N"));
                        Directory.CreateDirectory(extractDir);
                        result.ExtractedImagesDir = extractDir;

                        foreach (var kv in imageRelIds)
                        {
                            try
                            {
                                var imgRel = part.GetRelationship(kv.Value);
                                var imgUri = PackUriHelper.ResolvePartUri(part.Uri, imgRel.TargetUri);
                                if (!package.PartExists(imgUri)) continue;
                                var imgPart = package.GetPart(imgUri);

                                string ext = Path.GetExtension(imgUri.OriginalString);
                                string outPath = Path.Combine(extractDir, kv.Key + ext);
                                using (var outStream = File.Create(outPath))
                                using (var inStream = imgPart.GetStream(FileMode.Open, FileAccess.Read))
                                    inStream.CopyTo(outStream);

                                var button = result.Config.Items.SelectMany(i => i.Buttons)
                                    .FirstOrDefault(b => b.Id == kv.Key);
                                if (button?.Icon != null)
                                    button.Icon.ExtractedPreviewPath = outPath;
                            }
                            catch
                            {
                                // Bỏ qua ảnh lỗi - nút vẫn giữ được, chỉ thiếu preview icon.
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result.Error = "Không đọc được ribbon hiện có trong file (có thể file đang mở ở Excel khác " +
                                "hoặc không phải file Excel hợp lệ).\n\nChi tiết: " + ex.Message;
            }

            return result;
        }

        public static bool Save(string filePath, RibbonConfig config, IReadOnlyList<MacroEntry> knownMacros, out string error)
        {
            error = null;

            var ribbonShapedNames = new HashSet<string>(
                knownMacros.Where(m => m.HasRibbonSignature).Select(m => m.Name),
                StringComparer.OrdinalIgnoreCase);

            var buttons = config.Items.SelectMany(i => i.Buttons).ToList();

            var wrapperNeeded = buttons
                .Where(b => !string.IsNullOrWhiteSpace(b.MacroName) && !ribbonShapedNames.Contains(b.MacroName))
                .Select(b => b.MacroName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var onActionNames = new Dictionary<string, string>();
            foreach (var b in buttons)
            {
                if (string.IsNullOrWhiteSpace(b.MacroName)) continue;
                onActionNames[b.Id] = ribbonShapedNames.Contains(b.MacroName)
                    ? b.MacroName
                    : RibbonWrapperWriter.WrapperName(b.MacroName);
            }

            using (var session = new ExcelComSession())
            {
                if (!session.Open(filePath))
                {
                    error = session.Error;
                    return false;
                }

                if (!RibbonWrapperWriter.EnsureWrappers(session.Workbook, wrapperNeeded, out string wrapErr))
                {
                    error = wrapErr;
                    return false;
                }

                if (!session.SaveAndClose(out string saveErr))
                {
                    error = saveErr;
                    return false;
                }
            }

            try
            {
                using (var package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    RemoveExistingCustomUi(package);

                    var customUiPart = package.CreatePart(CustomUiPartUri, "application/xml");
                    var imageRelIds = new Dictionary<string, string>();

                    int imageIndex = 1;
                    foreach (var b in buttons)
                    {
                        if (b.Icon?.Kind != IconKind.Custom) continue;

                        string sourcePath = !string.IsNullOrWhiteSpace(b.Icon.SourceImagePath)
                            ? b.Icon.SourceImagePath
                            : b.Icon.ExtractedPreviewPath;
                        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
                            continue;

                        string ext = Path.GetExtension(sourcePath);
                        if (string.IsNullOrEmpty(ext)) ext = ".png";
                        var imgUri = new Uri($"/customUI/images/image{imageIndex}{ext}", UriKind.Relative);
                        imageIndex++;

                        var imgPart = package.CreatePart(imgUri, ImageContentType(ext));
                        using (var outStream = imgPart.GetStream(FileMode.Create, FileAccess.Write))
                        using (var inStream = File.OpenRead(sourcePath))
                            inStream.CopyTo(outStream);

                        var imgRel = customUiPart.CreateRelationship(imgUri, TargetMode.Internal, ImageRelType);
                        imageRelIds[b.Id] = imgRel.Id;
                    }

                    string xml = RibbonXmlBuilder.Build(config, onActionNames, imageRelIds);
                    using (var writer = new StreamWriter(customUiPart.GetStream(FileMode.Create, FileAccess.Write)))
                        writer.Write(xml);

                    package.CreateRelationship(CustomUiPartUri, TargetMode.Internal, ExtensibilityRelType);
                    package.Flush();
                }
            }
            catch (Exception ex)
            {
                error = "Đã lưu macro nhưng không ghi được ribbon XML vào file: " + ex.Message;
                return false;
            }

            return true;
        }

        private static void RemoveExistingCustomUi(Package package)
        {
            foreach (var rel in package.GetRelationshipsByType(ExtensibilityRelType).ToList())
            {
                try
                {
                    var partUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), rel.TargetUri);
                    if (package.PartExists(partUri))
                    {
                        var part = package.GetPart(partUri);
                        foreach (var imgRel in part.GetRelationshipsByType(ImageRelType).ToList())
                        {
                            try
                            {
                                var imgUri = PackUriHelper.ResolvePartUri(part.Uri, imgRel.TargetUri);
                                if (package.PartExists(imgUri))
                                    package.DeletePart(imgUri);
                            }
                            catch { /* best effort */ }
                            part.DeleteRelationship(imgRel.Id);
                        }

                        package.DeletePart(partUri);
                    }
                }
                catch { /* best effort */ }

                package.DeleteRelationship(rel.Id);
            }
        }

        private static string ImageContentType(string ext)
        {
            switch (ext.ToLowerInvariant())
            {
                case ".png": return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".bmp": return "image/bmp";
                case ".gif": return "image/gif";
                case ".ico": return "image/x-icon";
                default: return "image/png";
            }
        }
    }
}
