using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using RibbonXmlEditor.Config;

namespace RibbonXmlEditor.Ribbon
{
    /// <summary>
    /// Đọc ngược 1 customUI XML có sẵn (2007 hay 2010 schema đều được, so khớp theo LocalName
    /// nên không phụ thuộc namespace) thành RibbonConfig để hiện lên danh sách cho người dùng
    /// sửa tiếp - kể cả file đã được tạo bằng Office RibbonX Editor trước đó.
    /// </summary>
    public static class CustomUiXmlParser
    {
        /// <param name="imageRelIdsByButtonId">
        /// rId thô của icon tùy chỉnh cho từng button (key = ButtonInfo.Id) - CustomUiPackageService
        /// dùng để resolve ra file ảnh trích xuất tạm ngay sau khi Parse xong.
        /// </param>
        public static RibbonConfig Parse(string xml, out Dictionary<string, string> imageRelIdsByButtonId)
        {
            imageRelIdsByButtonId = new Dictionary<string, string>();
            var config = new RibbonConfig();
            if (string.IsNullOrWhiteSpace(xml))
                return config;

            XDocument doc;
            try { doc = XDocument.Parse(xml); }
            catch { return config; }

            var ribbon = doc.Root?.Elements().FirstOrDefault(e => e.Name.LocalName == "ribbon");
            var tabs = ribbon?.Elements().FirstOrDefault(e => e.Name.LocalName == "tabs");
            if (tabs == null) return config;

            config.StartFromScratch = string.Equals(Attr(ribbon, "startFromScratch"), "true", StringComparison.OrdinalIgnoreCase);

            foreach (var tabEl in tabs.Elements().Where(e => e.Name.LocalName == "tab"))
            {
                string idMso = Attr(tabEl, "idMso");
                bool isExisting = !string.IsNullOrEmpty(idMso);

                if (isExisting)
                {
                    config.BuiltInTabs.Add(new BuiltInTabVisibility
                    {
                        IdMso = idMso,
                        Visible = !string.Equals(Attr(tabEl, "visible"), "false", StringComparison.OrdinalIgnoreCase),
                    });
                }

                foreach (var groupEl in tabEl.Elements().Where(e => e.Name.LocalName == "group"))
                {
                    var target = new TargetInfo();
                    if (isExisting)
                    {
                        target.Kind = TargetKind.ExistingTab;
                        target.ExistingTabIdMso = idMso;
                        target.ExistingTabLabel = LookupTabLabel(idMso);

                        string before = Attr(groupEl, "insertBeforeMso");
                        string after = Attr(groupEl, "insertAfterMso");
                        if (!string.IsNullOrEmpty(before))
                        {
                            target.Position = AnchorPosition.Before;
                            target.AnchorGroupIdMso = before;
                            target.AnchorGroupLabel = LookupGroupLabel(idMso, before);
                        }
                        else if (!string.IsNullOrEmpty(after))
                        {
                            target.Position = AnchorPosition.After;
                            target.AnchorGroupIdMso = after;
                            target.AnchorGroupLabel = LookupGroupLabel(idMso, after);
                        }
                        else
                        {
                            target.Position = AnchorPosition.End;
                        }
                    }
                    else
                    {
                        target.Kind = TargetKind.NewTab;
                        target.NewTabId = Attr(tabEl, "id") ?? target.NewTabId;
                        target.NewTabLabel = Attr(tabEl, "label");
                        target.NewTabInsertBeforeMso = Attr(tabEl, "insertBeforeMso");
                        target.NewTabInsertAfterMso = Attr(tabEl, "insertAfterMso");
                    }

                    var item = new CustomizationItem
                    {
                        Id = Attr(groupEl, "id") ?? ("rc_grp_" + Guid.NewGuid().ToString("N")),
                        GroupLabel = Attr(groupEl, "label"),
                        Target = target,
                    };

                    foreach (var buttonEl in groupEl.Elements().Where(e => e.Name.LocalName == "button"))
                    {
                        var button = ParseButton(buttonEl, out string imageRelId);
                        item.Buttons.Add(button);
                        if (!string.IsNullOrEmpty(imageRelId))
                            imageRelIdsByButtonId[button.Id] = imageRelId;
                    }

                    foreach (var labelEl in groupEl.Elements().Where(e => e.Name.LocalName == "labelControl"))
                    {
                        item.LabelControls.Add(new LabelControlInfo
                        {
                            Id = Attr(labelEl, "id") ?? ("rc_label_" + Guid.NewGuid().ToString("N")),
                            Label = Attr(labelEl, "label"),
                        });
                    }

                    if (item.Buttons.Count > 0 || item.LabelControls.Count > 0)
                        config.Items.Add(item);
                }
            }

            return config;
        }

        private static ButtonInfo ParseButton(XElement buttonEl, out string imageRelId)
        {
            imageRelId = null;

            var button = new ButtonInfo
            {
                Id = Attr(buttonEl, "id") ?? ("rc_btn_" + Guid.NewGuid().ToString("N")),
                Label = Attr(buttonEl, "label"),
                ScreenTip = Attr(buttonEl, "screentip"),
                SuperTip = Attr(buttonEl, "supertip"),
                LargeSize = Attr(buttonEl, "size") != "normal",
                MacroName = Attr(buttonEl, "onAction"),
            };

            string imageMso = Attr(buttonEl, "imageMso");
            string image = Attr(buttonEl, "image");
            if (!string.IsNullOrEmpty(imageMso))
            {
                button.Icon = new IconInfo { Kind = IconKind.ImageMso, ImageMso = imageMso };
            }
            else if (!string.IsNullOrEmpty(image))
            {
                button.Icon = new IconInfo { Kind = IconKind.Custom };
                imageRelId = image;
            }

            return button;
        }

        private static string LookupTabLabel(string idMso)
        {
            var tab = BuiltInRibbonCatalog.Tabs.FirstOrDefault(t => t.IdMso == idMso);
            return tab?.DisplayLabel ?? idMso;
        }

        private static string LookupGroupLabel(string tabIdMso, string groupIdMso)
        {
            var tab = BuiltInRibbonCatalog.Tabs.FirstOrDefault(t => t.IdMso == tabIdMso);
            var group = tab?.Groups.FirstOrDefault(g => g.IdMso == groupIdMso);
            return group?.DisplayLabel ?? groupIdMso;
        }

        private static string Attr(XElement el, string name) => el.Attribute(name)?.Value;
    }
}
