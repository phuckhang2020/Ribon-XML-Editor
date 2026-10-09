using RibbonXmlEditor.Config;

namespace RibbonXmlEditor.UI
{
    /// <summary>1 dòng trong combo "chọn group đã có" - bọc 1 CustomizationItem đã tồn tại
    /// trong đúng tab đang chọn, để thêm nút mới vào chung group đó thay vì tạo group mới.</summary>
    public class GroupPickerEntry
    {
        public CustomizationItem Item { get; }

        public GroupPickerEntry(CustomizationItem item)
        {
            Item = item;
        }

        public override string ToString() => Item.GroupLabel;
    }
}
