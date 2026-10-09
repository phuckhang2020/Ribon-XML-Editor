using System.Collections.Generic;

namespace RibbonXmlEditor.Ribbon
{
    public class RibbonGroupRef
    {
        public string IdMso { get; }
        public string DisplayLabel { get; }

        public RibbonGroupRef(string idMso, string displayLabel)
        {
            IdMso = idMso;
            DisplayLabel = displayLabel;
        }

        public override string ToString() => DisplayLabel;
    }

    public class RibbonTabRef
    {
        public string IdMso { get; }
        public string DisplayLabel { get; }
        public IReadOnlyList<RibbonGroupRef> Groups { get; }

        public RibbonTabRef(string idMso, string displayLabel, IReadOnlyList<RibbonGroupRef> groups)
        {
            IdMso = idMso;
            DisplayLabel = displayLabel;
            Groups = groups;
        }

        public override string ToString() => DisplayLabel;
    }

    /// <summary>
    /// Danh sách idMso của các tab/group có sẵn trong Excel, dùng để chọn vị trí chèn group mới.
    /// Nguồn: tài liệu "Office Fluent UI Command Identifiers" (Excel 2010+) công khai của Microsoft.
    /// Nếu một vị trí chèn không đúng, Excel chỉ âm thầm bỏ qua group đó (không báo lỗi) - vì vậy
    /// wizard luôn cho phép nhập idMso tùy chỉnh (Advanced) để tự sửa nếu danh sách này thiếu/sai.
    /// </summary>
    public static class BuiltInRibbonCatalog
    {
        public static readonly IReadOnlyList<RibbonTabRef> Tabs = new List<RibbonTabRef>
        {
            new RibbonTabRef("TabHome", "Home", new List<RibbonGroupRef>
            {
                new RibbonGroupRef("GroupClipboard", "Clipboard"),
                new RibbonGroupRef("GroupFont", "Font"),
                new RibbonGroupRef("GroupAlignment", "Alignment"),
                new RibbonGroupRef("GroupNumber", "Number"),
                new RibbonGroupRef("GroupStylesExcel", "Styles"),
                new RibbonGroupRef("GroupCells", "Cells"),
                new RibbonGroupRef("GroupEditingExcel", "Editing"),
            }),
            new RibbonTabRef("TabInsert", "Insert", new List<RibbonGroupRef>
            {
                new RibbonGroupRef("GroupTables", "Tables"),
                new RibbonGroupRef("GroupChartsInsert", "Charts"),
                new RibbonGroupRef("GroupSparklines", "Sparklines"),
                new RibbonGroupRef("GroupFilterExcel", "Filters"),
                new RibbonGroupRef("GroupLinks", "Links"),
                new RibbonGroupRef("GroupText", "Text"),
                new RibbonGroupRef("GroupSymbols", "Symbols"),
            }),
            new RibbonTabRef("TabPageLayoutExcel", "Page Layout", new List<RibbonGroupRef>
            {
                new RibbonGroupRef("GroupThemes", "Themes"),
                new RibbonGroupRef("GroupPageSetup", "Page Setup"),
                new RibbonGroupRef("GroupScaleToFit", "Scale to Fit"),
                new RibbonGroupRef("GroupSheetOptions", "Sheet Options"),
                new RibbonGroupRef("GroupArrangeExcel", "Arrange"),
            }),
            new RibbonTabRef("TabFormulas", "Formulas", new List<RibbonGroupRef>
            {
                new RibbonGroupRef("GroupFunctionLibrary", "Function Library"),
                new RibbonGroupRef("GroupDefinedNames", "Defined Names"),
                new RibbonGroupRef("GroupFormulaAuditing", "Formula Auditing"),
                new RibbonGroupRef("GroupCalculation", "Calculation"),
            }),
            new RibbonTabRef("TabData", "Data", new List<RibbonGroupRef>
            {
                new RibbonGroupRef("GroupGetExternalData", "Get & Transform Data"),
                new RibbonGroupRef("GroupConnections", "Queries & Connections"),
                new RibbonGroupRef("GroupSortAndFilterExcel", "Sort & Filter"),
                new RibbonGroupRef("GroupDataTools", "Data Tools"),
                new RibbonGroupRef("GroupOutline", "Outline"),
            }),
            new RibbonTabRef("TabReview", "Review", new List<RibbonGroupRef>
            {
                new RibbonGroupRef("GroupProofing", "Proofing"),
                new RibbonGroupRef("GroupCommentsExcel", "Comments"),
                new RibbonGroupRef("GroupChanges", "Changes"),
            }),
            new RibbonTabRef("TabView", "View", new List<RibbonGroupRef>
            {
                new RibbonGroupRef("GroupSheetView", "Sheet View"),
                new RibbonGroupRef("GroupWorkbookViews", "Workbook Views"),
                new RibbonGroupRef("GroupShowHide", "Show"),
                new RibbonGroupRef("GroupZoom", "Zoom"),
                new RibbonGroupRef("GroupWindow", "Window"),
            }),
            new RibbonTabRef("TabAutomate", "Automate", new List<RibbonGroupRef>()),
            new RibbonTabRef("TabDeveloper", "Developer", new List<RibbonGroupRef>()),
        };
    }
}
