using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using RibbonXmlEditor.Config;

namespace RibbonXmlEditor.Ribbon
{
    /// <summary>
    /// Sinh Custom UI XML (customUI14) thuần túy từ RibbonConfig - nhúng thẳng vào file,
    /// không cần add-in/callback nào chạy nền. onAction gọi thẳng tên macro (đã resolve sẵn
    /// qua wrapper nếu cần); icon tùy chỉnh dùng attribute "image" trỏ tới rId ảnh đã nhúng
    /// trong package (resolve sẵn bởi CustomUiPackageService trước khi gọi Build).
    /// </summary>
    public static class RibbonXmlBuilder
    {
        public static string Build(
            RibbonConfig config,
            IReadOnlyDictionary<string, string> buttonOnActionNames,
            IReadOnlyDictionary<string, string> buttonImageRelIds)
        {
            var sb = new StringBuilder();
            sb.Append("<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");
            sb.Append("<ribbon");
            if (config.StartFromScratch)
                sb.Append(" startFromScratch=\"true\"");
            sb.Append("><tabs>");

            AppendExistingTabs(sb, config, buttonOnActionNames, buttonImageRelIds);
            AppendNewTabs(sb, config, buttonOnActionNames, buttonImageRelIds);

            sb.Append("</tabs></ribbon></customUI>");
            return sb.ToString();
        }

        private static void AppendExistingTabs(
            StringBuilder sb, RibbonConfig config,
            IReadOnlyDictionary<string, string> onActionNames, IReadOnlyDictionary<string, string> imageRelIds)
        {
            var configuredGroups = config.Items
                .Where(item => item.Target.Kind == TargetKind.ExistingTab
                    && !string.IsNullOrWhiteSpace(item.Target.ExistingTabIdMso))
                .GroupBy(item => item.Target.ExistingTabIdMso)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);

            var visibility = (config.BuiltInTabs ?? new List<BuiltInTabVisibility>())
                .Where(tab => !string.IsNullOrWhiteSpace(tab.IdMso))
                .GroupBy(tab => tab.IdMso, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Last().Visible, StringComparer.OrdinalIgnoreCase);

            var tabIds = new List<string>();
            if (config.StartFromScratch)
                tabIds.AddRange(BuiltInRibbonCatalog.Tabs.Select(tab => tab.IdMso));
            tabIds.AddRange(visibility.Keys);
            tabIds.AddRange(configuredGroups.Keys);

            foreach (string tabId in tabIds.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                bool hasVisibility = visibility.TryGetValue(tabId, out bool visibleSetting);
                bool visible = !hasVisibility || visibleSetting;
                bool hasGroups = configuredGroups.TryGetValue(tabId, out var groups);
                if (!config.StartFromScratch && !hasGroups && (!hasVisibility || visible))
                    continue;

                sb.Append("<tab idMso=\"").Append(Escape(tabId)).Append("\"");
                if (config.StartFromScratch || (hasVisibility && !visible))
                    sb.Append(" visible=\"").Append(visible ? "true" : "false").Append("\"");
                sb.Append(">");
                if (groups != null)
                {
                    foreach (var item in groups)
                        AppendGroup(sb, item, onActionNames, imageRelIds);
                }
                sb.Append("</tab>");
            }
        }

        private static void AppendNewTabs(
            StringBuilder sb, RibbonConfig config,
            IReadOnlyDictionary<string, string> onActionNames, IReadOnlyDictionary<string, string> imageRelIds)
        {
            // Gom theo NewTabId (giống AppendExistingTabInsertions gom theo ExistingTabIdMso):
            // nhiều group cùng trỏ 1 tab tự tạo phải nằm chung 1 phần tử <tab>, nếu không sẽ
            // sinh ra 2 <tab id="..."> trùng id - XML không hợp lệ. Thứ tự group trong tab tự
            // tạo là thứ tự xuất hiện trong config.Items; vị trí tương đối so với tab Office
            // được quyết định bởi insertBeforeMso/insertAfterMso.
            var byTab = config.Items
                .Where(i => i.Target.Kind == TargetKind.NewTab && !string.IsNullOrWhiteSpace(i.Target.NewTabId))
                .GroupBy(i => i.Target.NewTabId);

            foreach (var tabGroup in byTab)
            {
                string label = tabGroup.First().Target.NewTabLabel;
                sb.Append("<tab id=\"").Append(Escape(tabGroup.Key)).Append("\" ")
                  .Append("label=\"").Append(Escape(label)).Append("\"");
                var tabTarget = tabGroup.First().Target;
                if (!string.IsNullOrWhiteSpace(tabTarget.NewTabInsertBeforeMso))
                    sb.Append(" insertBeforeMso=\"").Append(Escape(tabTarget.NewTabInsertBeforeMso)).Append("\"");
                else if (!string.IsNullOrWhiteSpace(tabTarget.NewTabInsertAfterMso))
                    sb.Append(" insertAfterMso=\"").Append(Escape(tabTarget.NewTabInsertAfterMso)).Append("\"");
                sb.Append(">");
                foreach (var item in tabGroup)
                    AppendGroup(sb, item, onActionNames, imageRelIds);
                sb.Append("</tab>");
            }
        }

        private static void AppendGroup(
            StringBuilder sb, CustomizationItem item,
            IReadOnlyDictionary<string, string> onActionNames, IReadOnlyDictionary<string, string> imageRelIds)
        {
            sb.Append("<group id=\"").Append(Escape(item.Id)).Append("\" ")
              .Append("label=\"").Append(Escape(item.GroupLabel)).Append("\"");

            if (item.Target.Kind == TargetKind.ExistingTab)
            {
                switch (item.Target.Position)
                {
                    case AnchorPosition.Before when !string.IsNullOrWhiteSpace(item.Target.AnchorGroupIdMso):
                        sb.Append(" insertBeforeMso=\"").Append(Escape(item.Target.AnchorGroupIdMso)).Append("\"");
                        break;
                    case AnchorPosition.After when !string.IsNullOrWhiteSpace(item.Target.AnchorGroupIdMso):
                        sb.Append(" insertAfterMso=\"").Append(Escape(item.Target.AnchorGroupIdMso)).Append("\"");
                        break;
                        // AnchorPosition.End: không thêm attribute => Office tự thêm vào cuối tab.
                }
            }

            sb.Append(">");

            foreach (var button in item.Buttons)
                AppendButton(sb, button, onActionNames, imageRelIds);

            foreach (var labelControl in item.LabelControls ?? new List<LabelControlInfo>())
            {
                sb.Append("<labelControl id=\"").Append(Escape(labelControl.Id)).Append("\" ")
                    .Append("label=\"").Append(Escape(labelControl.Label)).Append("\"/>");
            }

            sb.Append("</group>");
        }

        private static void AppendButton(
            StringBuilder sb, ButtonInfo button,
            IReadOnlyDictionary<string, string> onActionNames, IReadOnlyDictionary<string, string> imageRelIds)
        {
            sb.Append("<button id=\"").Append(Escape(button.Id)).Append("\" ")
              .Append("label=\"").Append(Escape(button.Label)).Append("\" ")
              .Append("size=\"").Append(button.LargeSize ? "large" : "normal").Append("\" ");

            if (onActionNames.TryGetValue(button.Id, out string actionName) && !string.IsNullOrWhiteSpace(actionName))
                sb.Append("onAction=\"").Append(Escape(actionName)).Append("\" ");

            if (button.Icon != null && button.Icon.Kind == IconKind.ImageMso && !string.IsNullOrWhiteSpace(button.Icon.ImageMso))
            {
                sb.Append("imageMso=\"").Append(Escape(button.Icon.ImageMso)).Append("\" ");
            }
            else if (button.Icon != null && button.Icon.Kind == IconKind.Custom &&
                     imageRelIds.TryGetValue(button.Id, out string relId) && !string.IsNullOrWhiteSpace(relId))
            {
                sb.Append("image=\"").Append(Escape(relId)).Append("\" ");
            }

            if (!string.IsNullOrWhiteSpace(button.ScreenTip))
                sb.Append("screentip=\"").Append(Escape(button.ScreenTip)).Append("\" ");

            if (!string.IsNullOrWhiteSpace(button.SuperTip))
                sb.Append("supertip=\"").Append(Escape(button.SuperTip)).Append("\" ");

            sb.Append("/>");
        }

        private static string Escape(string value) => SecurityElement.Escape(value ?? string.Empty);
    }
}
