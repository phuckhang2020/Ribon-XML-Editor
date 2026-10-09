using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace RibbonXmlEditor.Macros
{
    public static class WorkbookVbaMacroReader
    {
        public const string CallbacksModuleName = "RibbonCustomizerCallbacks";

        // Sub không tham số - macro "bình thường" mà người dùng hay tự viết/ghi macro.
        private static readonly Regex PlainSubRegex = new Regex(
            @"^[ \t]*(?<modifier>Public\s+|Private\s+|Friend\s+)?Sub\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\s*\(\s*\)",
            RegexOptions.Multiline);

        // Đúng chữ ký callback ribbon: Sub Name(control As IRibbonControl)
        private static readonly Regex RibbonSubRegex = new Regex(
            @"^[ \t]*(?<modifier>Public\s+|Private\s+|Friend\s+)?Sub\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\s*\(\s*(ByVal\s+|ByRef\s+)?\w+\s+As\s+IRibbonControl\s*\)",
            RegexOptions.Multiline | RegexOptions.IgnoreCase);

        // Dò lại "Sub Wrapper(...) ... Call OriginalName ... End Sub" trong module wrapper do
        // chính tool này sinh ra, để khi mở lại file, wizard biết chọn sẵn đúng macro gốc.
        private static readonly Regex WrapperPairRegex = new Regex(
            @"Sub\s+(?<wrapper>[A-Za-z_][A-Za-z0-9_]*)\s*\([^)]*\)\s*.*?Call\s+(?<original>[A-Za-z_][A-Za-z0-9_]*)",
            RegexOptions.Singleline);

        public static List<MacroEntry> List(dynamic workbook, out string error)
        {
            error = null;
            var result = new List<MacroEntry>();

            dynamic components;
            try { components = workbook.VBProject.VBComponents; }
            catch (Exception ex)
            {
                error = "Không đọc được VBA project: " + ex.Message;
                return result;
            }

            foreach (dynamic component in components)
            {
                if (!TryReadModule(component, out string componentName, out string code, requireStdModule: true))
                    continue;

                if (string.Equals(componentName, CallbacksModuleName, StringComparison.OrdinalIgnoreCase))
                    continue; // module wrapper do tool sinh ra - không hiện lại cho người dùng chọn

                var ribbonNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Match m in RibbonSubRegex.Matches(code))
                {
                    if (IsPrivate(m)) continue;
                    string name = m.Groups["name"].Value;
                    ribbonNames.Add(name);
                    result.Add(new MacroEntry(name, componentName, true));
                }

                foreach (Match m in PlainSubRegex.Matches(code))
                {
                    if (IsPrivate(m)) continue;
                    string name = m.Groups["name"].Value;
                    if (ribbonNames.Contains(name)) continue;
                    result.Add(new MacroEntry(name, componentName, false));
                }
            }

            return result;
        }

        /// <summary>Map "tên wrapper" -> "tên macro gốc" đọc từ module wrapper (nếu có) - dùng để
        /// khi mở lại 1 file đã lưu trước đó bằng tool này, wizard chọn sẵn đúng macro gốc.</summary>
        public static Dictionary<string, string> BuildWrapperReverseMap(dynamic workbook)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                dynamic components = workbook.VBProject.VBComponents;
                foreach (dynamic component in components)
                {
                    if (!TryReadModule(component, out string componentName, out string code, requireStdModule: false))
                        continue;
                    if (!string.Equals(componentName, CallbacksModuleName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    foreach (Match m in WrapperPairRegex.Matches(code))
                        map[m.Groups["wrapper"].Value] = m.Groups["original"].Value;
                }
            }
            catch { /* best effort */ }

            return map;
        }

        private static bool TryReadModule(dynamic component, out string componentName, out string code, bool requireStdModule)
        {
            componentName = null;
            code = null;

            try
            {
                if (requireStdModule && (int)component.Type != 1) return false;
                componentName = (string)component.Name;

                dynamic codeModule = component.CodeModule;
                int countLines = (int)codeModule.CountOfLines;
                if (countLines <= 0) return false;

                code = (string)codeModule.Lines[1, countLines];
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsPrivate(Match m) =>
            m.Groups["modifier"].Value.Trim().Equals("Private", StringComparison.OrdinalIgnoreCase);
    }
}
