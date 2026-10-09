using RibbonXmlEditor.Ribbon;

namespace RibbonXmlEditor.UI
{
    /// <summary>
    /// 1 dòng trong combo "chọn tab": hoặc là tab có sẵn của Excel (từ BuiltInRibbonCatalog,
    /// cần chọn group neo + trước/sau/cuối), hoặc là 1 tab tự tạo đã có trong config hiện tại
    /// (group mới chỉ đơn giản được thêm vào cuối tab đó, không cần anchor).
    /// </summary>
    public class TabPickerEntry
    {
        public bool IsBuiltIn { get; }
        public RibbonTabRef BuiltIn { get; }
        public string Id { get; }
        public string Label { get; }

        public TabPickerEntry(RibbonTabRef builtIn)
        {
            IsBuiltIn = true;
            BuiltIn = builtIn;
            Id = builtIn.IdMso;
            Label = builtIn.DisplayLabel;
        }

        public TabPickerEntry(string customTabId, string customTabLabel)
        {
            IsBuiltIn = false;
            Id = customTabId;
            Label = customTabLabel;
        }

        public override string ToString() => IsBuiltIn ? Label : $"{Label} (tab tự tạo)";
    }
}
