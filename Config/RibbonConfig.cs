using System;
using System.Collections.Generic;

namespace RibbonXmlEditor.Config
{
    public enum TargetKind
    {
        ExistingTab,
        NewTab
    }

    public enum AnchorPosition
    {
        Before,
        After,
        End
    }

    public enum IconKind
    {
        ImageMso,
        Custom
    }

    public class IconInfo
    {
        public IconKind Kind { get; set; }
        public string ImageMso { get; set; }

        /// <summary>Đường dẫn file ảnh nguồn trên đĩa - dùng khi Save để nhúng ảnh vào package.</summary>
        public string SourceImagePath { get; set; }

        /// <summary>
        /// Khi nạp lại từ 1 file .xlsm đã có customUI: đường dẫn tạm đã trích xuất ảnh nhúng
        /// ra để xem trước / giữ nguyên nếu người dùng không đổi ảnh khác.
        /// </summary>
        public string ExtractedPreviewPath { get; set; }
    }

    public class ButtonInfo
    {
        public string Id { get; set; } = "rc_btn_" + Guid.NewGuid().ToString("N");
        public string Label { get; set; }
        public string ScreenTip { get; set; }
        public string SuperTip { get; set; }
        public bool LargeSize { get; set; } = true;
        public IconInfo Icon { get; set; } = new IconInfo();

        /// <summary>Tên Sub VBA trong chính workbook này, gọi thẳng qua onAction.</summary>
        public string MacroName { get; set; }
    }

    public class TargetInfo
    {
        public TargetKind Kind { get; set; }

        // Kind == ExistingTab
        public string ExistingTabIdMso { get; set; }
        public string ExistingTabLabel { get; set; }
        public string AnchorGroupIdMso { get; set; }
        public string AnchorGroupLabel { get; set; }
        public AnchorPosition Position { get; set; } = AnchorPosition.End;

        // Kind == NewTab
        public string NewTabLabel { get; set; }
        public string NewTabId { get; set; } = "rc_tab_" + Guid.NewGuid().ToString("N");
        public string NewTabInsertBeforeMso { get; set; }
        public string NewTabInsertAfterMso { get; set; }
    }

    public class LabelControlInfo
    {
        public string Id { get; set; } = "rc_label_" + Guid.NewGuid().ToString("N");
        public string Label { get; set; }
    }

    public class CustomizationItem
    {
        public string Id { get; set; } = "rc_grp_" + Guid.NewGuid().ToString("N");
        public string GroupLabel { get; set; }
        public TargetInfo Target { get; set; } = new TargetInfo();
        public List<ButtonInfo> Buttons { get; set; } = new List<ButtonInfo>();
        public List<LabelControlInfo> LabelControls { get; set; } = new List<LabelControlInfo>();
    }

    public class BuiltInTabVisibility
    {
        public string IdMso { get; set; }
        public bool Visible { get; set; } = true;
    }

    public class RibbonConfig
    {
        public bool StartFromScratch { get; set; }
        public List<CustomizationItem> Items { get; set; } = new List<CustomizationItem>();
        public List<BuiltInTabVisibility> BuiltInTabs { get; set; } = new List<BuiltInTabVisibility>();
    }
}
