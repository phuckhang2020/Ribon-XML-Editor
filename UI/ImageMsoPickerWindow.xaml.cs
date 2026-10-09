using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using RibbonXmlEditor.Localization;

namespace RibbonXmlEditor.UI
{
    public partial class ImageMsoPickerWindow : Window
    {
        private static readonly string[] CommonImageMso =
        {
            "HappyFace", "FileSave", "FileSaveAs", "FileOpen", "FilePrintPreview", "FilePrint",
            "Copy", "Paste", "Cut", "Refresh", "RefreshAll",
            "FilterApply", "FilterClearAllFilters", "SortAscending", "SortDescending",
            "TableInsert", "ChartInsert", "PivotTableInsert", "PivotChartInsert",
            "ChartPie", "ChartBarStacked", "ChartLine", "ChartColumnClustered",
            "ImageInsertFromFile", "HyperlinkInsert", "TextBoxInsert",
            "CommentInsert", "CommentInsertNew", "Bold", "Italic", "Underline",
            "FontColorPicker", "FillColorPicker", "BordersAll",
            "AlignLeft", "AlignCenter", "AlignRight", "MergeCenter",
            "NumberFormatIncreaseDecimal", "NumberFormatGeneral", "CellStylesGallery",
            "FreezePanes", "ViewNormalWorkbook", "ViewPageLayout", "ZoomToSelection",
            "MacroPlay", "MacroRecord", "MacroSecurity", "VisualBasic", "AddInsManage",
            "Calculator", "DatabaseRefreshAll", "DataFromWeb", "DataFromText", "DataValidation",
            "GroupRows", "UngroupRows", "Subtotal", "GoalSeek", "SolverDialog",
            "TrackChangesToggle", "ProtectSheet", "ProtectWorkbook",
            "SpellingAndGrammar", "Translate", "Find", "FindAndReplace", "GoTo",
            "NameManager", "FunctionWizard", "AutoSum", "InsertFunction",
            "Export", "ExportExcel", "Import", "ImportExcel",
            "Info", "Help", "Warning", "BlockedError", "Delete", "Add",
            "New", "MailEmailAsAttachment", "Printer", "Play", "Stop", "Pause",
            "Timeline", "Calendar", "LightBulb", "Tools", "Database", "Table",
        };

        private readonly string _currentValue;

        public string SelectedImageMso { get; private set; }

        public ImageMsoPickerWindow(string currentValue)
        {
            _currentValue = currentValue;
            InitializeComponent();
            LanguageService.Apply(this);
            ApplyFilter();
            if (!string.IsNullOrWhiteSpace(currentValue))
                IconList.SelectedItem = currentValue;
        }

        private void ApplyFilter()
        {
            if (IconList == null)
                return;

            string query = SearchBox?.Text?.Trim() ?? string.Empty;
            string selected = IconList?.SelectedItem as string ?? _currentValue;
            IconList.ItemsSource = CommonImageMso
                .Where(name => name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (!string.IsNullOrEmpty(selected) && IconList.Items.Contains(selected))
                IconList.SelectedItem = selected;
        }

        private void SearchBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void IconList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            SelectedIconText.Text = IconList.SelectedItem == null
                ? LanguageService.Get("icon.notSelected")
                : "imageMso: " + IconList.SelectedItem;
        }

        private void IconList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            AcceptSelection();
        }

        private void AcceptButton_Click(object sender, RoutedEventArgs e)
        {
            AcceptSelection();
        }

        private void AcceptSelection()
        {
            if (!(IconList.SelectedItem is string selected))
            {
                MessageBox.Show(this, LanguageService.Get("icon.chooseFirst"), LanguageService.Get("icon.chooseTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SelectedImageMso = selected;
            DialogResult = true;
        }
    }
}
