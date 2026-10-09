using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32;
using RibbonXmlEditor.Config;
using RibbonXmlEditor.Localization;
using RibbonXmlEditor.Macros;
using RibbonXmlEditor.Packaging;
using RibbonXmlEditor.Ribbon;

namespace RibbonXmlEditor.UI
{
    public partial class MainWindow : Window
    {
        private string _filePath;
        private RibbonConfig _config = new RibbonConfig();
        private List<MacroEntry> _macros = new List<MacroEntry>();
        private List<RibbonButtonRow> _rows = new List<RibbonButtonRow>();
        private ICollectionView _buttonView;
        private bool _isDirty;

        public MainWindow()
        {
            InitializeComponent();
            LanguageService.Apply(this);
            LanguageService.LanguageChanged += LanguageChanged;
            _buttonView = CollectionViewSource.GetDefaultView(_rows);
            _buttonView.Filter = FilterButton;
            ButtonGrid.ItemsSource = _buttonView;
            ButtonGrid.Columns[0].Header = LanguageService.Get("main.nameColumn");
            ButtonGrid.Columns[1].Header = LanguageService.Get("main.groupColumn");
            ButtonGrid.Columns[2].Header = LanguageService.Get("main.location");
            ButtonGrid.Columns[3].Header = LanguageService.Get("main.macroColumn");
            SettingsButton.ToolTip = LanguageService.Get("settings.iconHint");
            RefreshUi();
        }

        private void LanguageChanged(object sender, EventArgs e)
        {
            ButtonGrid.Columns[0].Header = LanguageService.Get("main.nameColumn");
            ButtonGrid.Columns[1].Header = LanguageService.Get("main.groupColumn");
            ButtonGrid.Columns[2].Header = LanguageService.Get("main.location");
            ButtonGrid.Columns[3].Header = LanguageService.Get("main.macroColumn");
            SettingsButton.ToolTip = LanguageService.Get("settings.iconHint");
            RefreshUi();
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            new SettingsWindow { Owner = this }.ShowDialog();
        }

        private bool FilterButton(object value)
        {
            if (!(value is RibbonButtonRow row))
                return false;

            string query = SearchBox.Text?.Trim();
            if (string.IsNullOrEmpty(query))
                return true;

            return Contains(row.Button?.Label, query)
                || Contains(row.Item.GroupLabel, query)
                || Contains(row.Location, query)
                || Contains(row.Button?.MacroName, query)
                || (row.Item.LabelControls ?? new List<LabelControlInfo>()).Any(label => Contains(label.Label, query));
        }

        private static bool Contains(string value, string query)
        {
            return !string.IsNullOrEmpty(value)
                && value.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel macro-enabled (*.xlsm;*.xlam)|*.xlsm;*.xlam|" + LanguageService.Get("main.fileFilter"),
                Title = LanguageService.Get("main.openDialogTitle"),
            };
            if (dialog.ShowDialog(this) != true || !ConfirmDiscardChanges())
                return;

            LoadFile(dialog.FileName);
        }

        private void LoadFile(string path)
        {
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            try
            {
                var loadResult = CustomUiPackageService.Load(path);
                if (!string.IsNullOrEmpty(loadResult.Error))
                {
                    MessageBox.Show(this, loadResult.Error, LanguageService.Get("main.openErrorTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string macroError;
                var macros = ScanMacrosAndResolveNames(path, loadResult.Config, out macroError);

                _filePath = path;
                _config = loadResult.Config;
                _macros = macros;
                _isDirty = false;
                RebuildRows();
                RefreshUi();

                if (!string.IsNullOrEmpty(macroError))
                    MessageBox.Show(this, macroError, LanguageService.Get("main.scanMacros"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, string.Format(LanguageService.Get("main.openErrorDetails"), ex.Message),
                    LanguageService.Get("main.openErrorTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private List<MacroEntry> ScanMacrosAndResolveNames(string path, RibbonConfig config, out string error)
        {
            error = null;
            using (var session = new ExcelComSession())
            {
                if (!session.Open(path))
                {
                    error = session.Error;
                    return new List<MacroEntry>();
                }

                var macros = WorkbookVbaMacroReader.List(session.Workbook, out string listError);
                var reverseMap = WorkbookVbaMacroReader.BuildWrapperReverseMap(session.Workbook);

                foreach (var button in config.Items.SelectMany(item => item.Buttons))
                {
                    if (string.IsNullOrEmpty(button.MacroName))
                        continue;

                    if (reverseMap.TryGetValue(button.MacroName, out string original))
                        button.MacroName = original;
                }

                if (!string.IsNullOrEmpty(listError))
                    error = listError;
                return macros;
            }
        }

        private List<MacroEntry> RescanMacros()
        {
            if (string.IsNullOrEmpty(_filePath))
                return _macros;

            using (var session = new ExcelComSession())
            {
                if (!session.Open(_filePath))
                {
                    MessageBox.Show(this, session.Error, LanguageService.Get("main.scanMacros"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return _macros;
                }

                _macros = WorkbookVbaMacroReader.List(session.Workbook, out string error);
                if (!string.IsNullOrEmpty(error))
                    MessageBox.Show(this, error, LanguageService.Get("main.scanMacros"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            RefreshUi();
            return _macros;
        }

        private void RebuildRows(string selectedButtonId = null)
        {
            _rows = _config.Items
                .SelectMany(item => item.Buttons.Count > 0
                    ? item.Buttons.Select(button => new RibbonButtonRow(item, button, DescribeTarget(item.Target)))
                    : new[] { new RibbonButtonRow(item, null, DescribeTarget(item.Target)) })
                .ToList();
            _buttonView = CollectionViewSource.GetDefaultView(_rows);
            _buttonView.Filter = FilterButton;
            ButtonGrid.ItemsSource = _buttonView;
            _buttonView.Refresh();

            if (selectedButtonId != null)
            {
                var selected = _rows.FirstOrDefault(row => row.Button.Id == selectedButtonId);
                if (selected != null)
                    ButtonGrid.SelectedItem = selected;
            }
        }

        private string DescribeTarget(TargetInfo target)
        {
            if (target.Kind == TargetKind.NewTab)
            {
                string insertion = !string.IsNullOrWhiteSpace(target.NewTabInsertBeforeMso)
                    ? string.Format(LanguageService.Get("main.insertBefore"),
                        BuiltInRibbonCatalog.Tabs.FirstOrDefault(tab => tab.IdMso == target.NewTabInsertBeforeMso)?.DisplayLabel ?? target.NewTabInsertBeforeMso)
                    : !string.IsNullOrWhiteSpace(target.NewTabInsertAfterMso)
                        ? string.Format(LanguageService.Get("main.insertAfter"),
                            BuiltInRibbonCatalog.Tabs.FirstOrDefault(tab => tab.IdMso == target.NewTabInsertAfterMso)?.DisplayLabel ?? target.NewTabInsertAfterMso)
                        : string.Empty;
                return string.Format(LanguageService.Get("main.newTabLocation"), target.NewTabLabel, insertion);
            }

            string position = target.Position == AnchorPosition.End
                ? LanguageService.Get("main.endOfTab")
                : target.Position == AnchorPosition.Before
                    ? string.Format(LanguageService.Get("main.beforeGroup"), target.AnchorGroupLabel)
                    : string.Format(LanguageService.Get("main.afterGroup"), target.AnchorGroupLabel);
            return string.Format(LanguageService.Get("main.existingTabLocation"), target.ExistingTabLabel, position);
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            var window = new AddButtonWindow(null, null, null, _config.Items, _macros, RescanMacros) { Owner = this };
            if (window.ShowDialog() != true)
                return;

            if (window.ResultExistingGroupItem != null)
            {
                window.ResultExistingGroupItem.Buttons.Add(window.ResultButton);
            }
            else
            {
                _config.Items.Add(new CustomizationItem
                {
                    GroupLabel = window.ResultGroupLabel,
                    Target = window.ResultTarget,
                    Buttons = window.ResultButton == null ? new List<ButtonInfo>() : new List<ButtonInfo> { window.ResultButton },
                    LabelControls = window.ResultLabelControls ?? new List<LabelControlInfo>(),
                });
            }

            MarkDirty();
            RebuildRows(window.ResultButton?.Id);
        }

        private void AddTextGroupButton_Click(object sender, RoutedEventArgs e)
        {
            var window = new AddButtonWindow(null, null, null, _config.Items, _macros, RescanMacros,
                allowTextOnlyGroup: true) { Owner = this };
            if (window.ShowDialog() != true)
                return;

            _config.Items.Add(new CustomizationItem
            {
                GroupLabel = window.ResultGroupLabel,
                Target = window.ResultTarget,
                LabelControls = window.ResultLabelControls ?? new List<LabelControlInfo>(),
            });
            MarkDirty();
            RebuildRows();
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(ButtonGrid.SelectedItem is RibbonButtonRow row))
                return;

            var window = new AddButtonWindow(row.Item.Target, row.Button, row.Item.GroupLabel,
                _config.Items, _macros, RescanMacros, row.Item, row.Button == null)
            {
                Owner = this,
            };
            if (window.ShowDialog() != true)
                return;

            if (row.Button == null)
            {
                row.Item.GroupLabel = window.ResultGroupLabel;
                row.Item.Target = window.ResultTarget;
                row.Item.LabelControls = window.ResultLabelControls ?? new List<LabelControlInfo>();
                MarkDirty();
                RebuildRows();
                return;
            }

            bool movingToDifferentGroup = window.ResultExistingGroupItem != null
                && window.ResultExistingGroupItem.Id != row.Item.Id;

            if (movingToDifferentGroup)
            {
                row.Item.Buttons.Remove(row.Button);
                if (row.Item.Buttons.Count == 0 && row.Item.LabelControls.Count == 0)
                    _config.Items.Remove(row.Item);
                window.ResultExistingGroupItem.Buttons.Add(window.ResultButton);
            }
            else
            {
                row.Item.GroupLabel = window.ResultGroupLabel;
                row.Item.Target = window.ResultTarget;
                CopyButton(window.ResultButton, row.Button);
                row.Item.LabelControls = window.ResultLabelControls ?? new List<LabelControlInfo>();
            }

            MarkDirty();
            RebuildRows(window.ResultButton.Id);
        }

        private static void CopyButton(ButtonInfo source, ButtonInfo destination)
        {
            destination.Label = source.Label;
            destination.ScreenTip = source.ScreenTip;
            destination.SuperTip = source.SuperTip;
            destination.LargeSize = source.LargeSize;
            destination.Icon = source.Icon;
            destination.MacroName = source.MacroName;
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(ButtonGrid.SelectedItem is RibbonButtonRow row))
                return;

            string confirmation = row.Button == null
                ? string.Format(LanguageService.Get("main.deleteGroupConfirm"), row.Item.GroupLabel)
                : string.Format(LanguageService.Get("main.deleteButtonConfirm"), row.Button.Label);
            if (MessageBox.Show(this, confirmation, LanguageService.Get("main.confirmDelete"),
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            if (row.Button != null)
                row.Item.Buttons.Remove(row.Button);
            if (row.Button == null || row.Item.Buttons.Count == 0 && row.Item.LabelControls.Count == 0)
                _config.Items.Remove(row.Item);

            MarkDirty();
            RebuildRows();
        }

        private void MoveButton(int direction)
        {
            if (!(ButtonGrid.SelectedItem is RibbonButtonRow row))
                return;

            int index = _config.Items.IndexOf(row.Item);
            int newIndex = index + direction;
            if (index < 0 || newIndex < 0 || newIndex >= _config.Items.Count)
                return;

            _config.Items.RemoveAt(index);
            _config.Items.Insert(newIndex, row.Item);
            MarkDirty();
            RebuildRows(row.Button.Id);
        }

        private void MoveUpButton_Click(object sender, RoutedEventArgs e) => MoveButton(-1);

        private void MoveDownButton_Click(object sender, RoutedEventArgs e) => MoveButton(1);

        private void RibbonOptionsButton_Click(object sender, RoutedEventArgs e)
        {
            var window = new RibbonOptionsWindow(_config) { Owner = this };
            if (window.ShowDialog() != true)
                return;

            _config.StartFromScratch = window.StartFromScratch;
            _config.BuiltInTabs = window.BuiltInTabs;
            MarkDirty();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_filePath))
                return;

            if (MessageBox.Show(this,
                    LanguageService.Get("main.saveWarning"),
                    LanguageService.Get("main.saveTitle"), MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            SaveButton.IsEnabled = false;
            try
            {
                if (CustomUiPackageService.Save(_filePath, _config, _macros, out string error))
                {
                    _isDirty = false;
                    RefreshUi();
                    MessageBox.Show(this, LanguageService.Get("main.saveSuccess"),
                        LanguageService.Get("main.saveSuccessTitle"), MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show(this, error, LanguageService.Get("main.saveErrorTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            finally
            {
                Mouse.OverrideCursor = null;
                RefreshUi();
            }
        }

        private void SearchBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            _buttonView?.Refresh();
            RefreshUi();
        }

        private void ButtonGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            RefreshUi();
        }

        private void MarkDirty()
        {
            _isDirty = true;
            RefreshUi();
        }

        private void RefreshUi()
        {
            bool hasFile = !string.IsNullOrEmpty(_filePath);
            bool hasSelection = ButtonGrid?.SelectedItem is RibbonButtonRow;
            int buttonCount = _config.Items.Sum(item => item.Buttons.Count);

            WorkbookNameText.Text = hasFile ? System.IO.Path.GetFileName(_filePath) : LanguageService.Get("main.noWorkbook");
            WorkbookPathText.Text = hasFile ? _filePath : LanguageService.Get("main.chooseWorkbook");
            SaveStateText.Text = !hasFile ? LanguageService.Get("main.noFile")
                : _isDirty ? LanguageService.Get("main.unsaved") : LanguageService.Get("main.saved");
            SaveStateText.Foreground = _isDirty
                ? (System.Windows.Media.Brush)FindResource("AccentBrush")
                : (System.Windows.Media.Brush)FindResource("MutedBrush");
            ButtonCountText.Text = buttonCount.ToString();
            GroupCountText.Text = _config.Items.Count.ToString();
            MacroCountText.Text = _macros.Count.ToString();
            AddButton.IsEnabled = hasFile;
            AddTextGroupButton.IsEnabled = hasFile;
            RibbonOptionsButton.IsEnabled = hasFile;
            EditButton.IsEnabled = hasFile && hasSelection;
            DeleteButton.IsEnabled = hasFile && hasSelection;
            SaveButton.IsEnabled = hasFile;
            MoveUpButton.IsEnabled = hasFile && hasSelection && SelectedGroupIndex() > 0;
            MoveDownButton.IsEnabled = hasFile && hasSelection && SelectedGroupIndex() < _config.Items.Count - 1;
            EmptyStatePanel.Visibility = _rows.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            ResultCountText.Text = string.Format(LanguageService.Get("main.resultCount"),
                _buttonView == null ? 0 : _buttonView.Cast<object>().Count());
        }

        private int SelectedGroupIndex()
        {
            return ButtonGrid.SelectedItem is RibbonButtonRow row ? _config.Items.IndexOf(row.Item) : -1;
        }

        private bool ConfirmDiscardChanges()
        {
            if (!_isDirty)
                return true;

            return MessageBox.Show(this,
                LanguageService.Get("main.dirtyConfirm"),
                LanguageService.Get("main.dirtyTitle"), MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes;
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!ConfirmDiscardChanges())
                e.Cancel = true;
        }

        private sealed class RibbonButtonRow
        {
            public RibbonButtonRow(CustomizationItem item, ButtonInfo button, string location)
            {
                Item = item;
                Button = button;
                Location = location;
            }

            public CustomizationItem Item { get; }
            public ButtonInfo Button { get; }
            public string Location { get; }
            public string ButtonDisplay => Button?.Label
                ?? string.Join(" · ", (Item.LabelControls ?? new List<LabelControlInfo>()).Select(label => label.Label));
        }
    }
}
