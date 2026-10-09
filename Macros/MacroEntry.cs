namespace RibbonXmlEditor.Macros
{
    /// <summary>Một macro (Sub) VBA tìm thấy trong workbook đang được chỉnh sửa.</summary>
    public class MacroEntry
    {
        public string Name { get; }
        public string ModuleName { get; }

        /// <summary>True nếu đã đúng chữ ký callback ribbon: Sub X(control As IRibbonControl).</summary>
        public bool HasRibbonSignature { get; }

        public MacroEntry(string name, string moduleName, bool hasRibbonSignature)
        {
            Name = name;
            ModuleName = moduleName;
            HasRibbonSignature = hasRibbonSignature;
        }

        public override string ToString() => $"{ModuleName}.{Name}";
    }
}
