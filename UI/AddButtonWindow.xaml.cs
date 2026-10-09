using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using RibbonXmlEditor.Config;
using RibbonXmlEditor.Localization;
using RibbonXmlEditor.Macros;
using RibbonXmlEditor.Ribbon;

namespace RibbonXmlEditor.UI
{
    public partial class AddButtonWindow : Window
    {
        private readonly TargetInfo _editingTarget;
        private readonly ButtonInfo _editingButton;
        private readonly CustomizationItem _editingItem;
        private readonly bool _allowTextOnlyGroup;
        private readonly List<CustomizationItem> _existingItems;
        private readonly Func<List<MacroEntry>> _rescanMacros;
        private List<MacroEntry> _macros;
        private string _selectedImageMso;
        private string _selectedImagePath;
        private bool _updatingPositionControls;
        private bool _isInitialized;

        public TargetInfo ResultTarget { get; private set; }
        public ButtonInfo ResultButton { get; private set; }
        public string ResultGroupLabel { get; private set; }
        public CustomizationItem ResultExistingGroupItem { get; private set; }
        public List<LabelControlInfo> ResultLabelControls { get; private set; }

        public AddButtonWindow(TargetInfo editingTarget, ButtonInfo editingButton, string editingGroupLabel,
            List<CustomizationItem> existingItems, List<MacroEntry> macros, Func<List<MacroEntry>> rescanMacros,
            CustomizationItem editingItem = null, bool allowTextOnlyGroup = false)
        {
            _editingItem = editingItem;
            _editingTarget = editingItem?.Target ?? editingTarget;
            _editingButton = editingButton;
            _allowTextOnlyGroup = allowTextOnlyGroup;
            _existingItems = existingItems ?? new List<CustomizationItem>();
            _macros = macros ?? new List<MacroEntry>();
            _rescanMacros = rescanMacros;

            InitializeComponent();
            _isInitialized = true;
            string titleKey = editingItem != null && editingButton == null
                ? "target.editGroupTitle"
                : allowTextOnlyGroup ? "target.createGroupTitle"
                : editingButton == null ? "target.addTitle" : "target.editTitle";
            bool labelOnlyMode = allowTextOnlyGroup || editingItem != null && editingButton == null;
            TextOnlyGroupCheck.Visibility = allowTextOnlyGroup ? Visibility.Visible : Visibility.Collapsed;
            PopulateTabs();
            PopulateNewTabAnchors();
            PopulateMacroCombo();
            PrefillFromExisting(editingGroupLabel);
            LanguageService.Apply(this);
            Title = LanguageService.Get(titleKey);
            WindowHeading.Text = Title;
            if (labelOnlyMode)
            {
                MacroSection.Visibility = Visibility.Collapsed;
                IconSection.Visibility = Visibility.Collapsed;
                ButtonDetailsSection.Visibility = Visibility.Collapsed;
                TextOnlyGroupCheck.Visibility = Visibility.Collapsed;
                LabelSection.Header = LanguageService.Get("section.addLabels");
                LabelHintText.Text = LanguageService.Get("labels.addHint");
                Height = 720;
                MinHeight = 620;
                SaveButton.Content = LanguageService.Get("target.addLabelSave");
            }
            if (labelOnlyMode)
                SubtitleText.Text = LanguageService.Get(
                    allowTextOnlyGroup && editingItem == null ? "target.addLabelHint" : "target.editLabelHint");
            ApplyDynamicLanguageText();
        }

        private void ApplyDynamicLanguageText()
        {
            if (MacroErrorText.Visibility == Visibility.Visible)
            {
                var selectedMacro = MacroCombo.SelectedItem as MacroEntry;
                MacroErrorText.Text = selectedMacro != null && selectedMacro.ModuleName == "?"
                    ? string.Format(LanguageService.Get("target.macroMissing"), selectedMacro.Name)
                    : LanguageService.Get("target.noMacro");
            }

            if (!string.IsNullOrEmpty(_selectedImageMso))
                IconValueText.Text = "imageMso: " + _selectedImageMso;
            else if (string.IsNullOrEmpty(_selectedImagePath))
            {
                IconValueText.Text = LanguageService.Get("icon.none");
                ImagePathText.Text = LanguageService.Get("icon.imageFormats");
            }
        }

        private void PopulateTabs()
        {
            TabCombo.Items.Clear();
            foreach (var tab in BuiltInRibbonCatalog.Tabs)
                TabCombo.Items.Add(new TabPickerEntry(tab));

            var customTabs = _existingItems
                .Where(item => item.Target.Kind == TargetKind.NewTab && !string.IsNullOrWhiteSpace(item.Target.NewTabId))
                .GroupBy(item => item.Target.NewTabId)
                .Select(group => new TabPickerEntry(group.Key, group.First().Target.NewTabLabel));
            foreach (var tab in customTabs)
                TabCombo.Items.Add(tab);

            if (TabCombo.Items.Count > 0)
                TabCombo.SelectedIndex = 0;
        }

        private void PopulateNewTabAnchors()
        {
            foreach (var tab in BuiltInRibbonCatalog.Tabs)
                NewTabAnchorCombo.Items.Add(tab);
            if (NewTabAnchorCombo.Items.Count > 0)
                NewTabAnchorCombo.SelectedIndex = 0;
        }

        private void SelectNewTabAnchor(string idMso)
        {
            if (string.IsNullOrEmpty(idMso))
                return;

            for (int i = 0; i < NewTabAnchorCombo.Items.Count; i++)
            {
                if (NewTabAnchorCombo.Items[i] is RibbonTabRef tab && tab.IdMso == idMso)
                {
                    NewTabAnchorCombo.SelectedIndex = i;
                    return;
                }
            }
        }

        private void PopulateMacroCombo()
        {
            string previousName = (MacroCombo.SelectedItem as MacroEntry)?.Name;
            MacroCombo.Items.Clear();
            foreach (var macro in _macros)
                MacroCombo.Items.Add(macro);

            if (_macros.Count == 0)
            {
                MacroErrorText.Text = LanguageService.Get("target.noMacro");
                MacroErrorText.Visibility = Visibility.Visible;
            }
            else
            {
                MacroErrorText.Visibility = Visibility.Collapsed;
            }

            string nameToSelect = previousName ?? _editingButton?.MacroName;
            if (!string.IsNullOrEmpty(nameToSelect))
                SelectMacroByName(nameToSelect);
            else if (MacroCombo.Items.Count > 0)
                MacroCombo.SelectedIndex = 0;
        }

        private void SelectMacroByName(string name)
        {
            for (int i = 0; i < MacroCombo.Items.Count; i++)
            {
                if (MacroCombo.Items[i] is MacroEntry entry
                    && string.Equals(entry.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    MacroCombo.SelectedIndex = i;
                    return;
                }
            }

            var stale = new MacroEntry(name, "?", false);
            MacroCombo.Items.Insert(0, stale);
            MacroCombo.SelectedIndex = 0;
            MacroErrorText.Text = string.Format(LanguageService.Get("target.macroMissing"), name);
            MacroErrorText.Visibility = Visibility.Visible;
        }

        private void TargetMode_Changed(object sender, RoutedEventArgs e)
        {
            UpdateTargetMode();
        }

        private void UpdateTargetMode()
        {
            if (!_isInitialized)
                return;

            bool existingMode = ExistingTabRadio.IsChecked == true;
            ExistingTabPanel.Visibility = existingMode ? Visibility.Visible : Visibility.Collapsed;
            NewTabPanel.Visibility = existingMode ? Visibility.Collapsed : Visibility.Visible;
            TabCombo.IsEnabled = existingMode;
            NewTabLabelBox.IsEnabled = !existingMode;
            NewTabAnchorCombo.IsEnabled = !existingMode;
            RefreshPositionControlsState();
        }

        private void TabCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            RefreshPositionControlsState();
        }

        private void GroupMode_Changed(object sender, RoutedEventArgs e)
        {
            RefreshPositionControlsState();
        }

        private void ExistingGroupCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ExistingGroupCombo.SelectedItem is GroupPickerEntry group && ExistingGroupRadio.IsChecked == true)
                GroupLabelBox.Text = group.Item.GroupLabel;
        }

        private void RefreshPositionControlsState()
        {
            if (!_isInitialized || _updatingPositionControls || TabCombo == null)
                return;

            _updatingPositionControls = true;
            try
            {
                bool existingMode = ExistingTabRadio.IsChecked == true;
                var tab = TabCombo.SelectedItem as TabPickerEntry;
                var candidates = tab == null
                    ? new List<GroupPickerEntry>()
                    : _existingItems
                        .Where(item => tab.IsBuiltIn
                            ? item.Target.Kind == TargetKind.ExistingTab && item.Target.ExistingTabIdMso == tab.Id
                            : item.Target.Kind == TargetKind.NewTab && item.Target.NewTabId == tab.Id)
                        .Select(item => new GroupPickerEntry(item))
                        .ToList();

                string previousId = (ExistingGroupCombo.SelectedItem as GroupPickerEntry)?.Item.Id;
                ExistingGroupCombo.Items.Clear();
                foreach (var candidate in candidates)
                    ExistingGroupCombo.Items.Add(candidate);
                int selectedIndex = candidates.FindIndex(candidate => candidate.Item.Id == previousId);
                if (selectedIndex < 0 && candidates.Count > 0)
                    selectedIndex = 0;
                ExistingGroupCombo.SelectedIndex = selectedIndex;

                ExistingGroupRadio.IsEnabled = existingMode && candidates.Count > 0;
                if (_editingItem != null && _editingButton == null)
                    ExistingGroupRadio.IsEnabled = false;
                if (candidates.Count == 0)
                    NewGroupRadio.IsChecked = true;

                bool joiningExisting = existingMode && ExistingGroupRadio.IsChecked == true && candidates.Count > 0;
                NewGroupRadio.Visibility = existingMode ? Visibility.Visible : Visibility.Collapsed;
                ExistingGroupRadio.Visibility = existingMode ? Visibility.Visible : Visibility.Collapsed;
                ExistingGroupCombo.IsEnabled = joiningExisting;
                ExistingGroupCombo.Visibility = joiningExisting ? Visibility.Visible : Visibility.Collapsed;
                GroupChooserLabel.Visibility = joiningExisting ? Visibility.Visible : Visibility.Collapsed;
                GroupLabelPanel.Visibility = joiningExisting ? Visibility.Collapsed : Visibility.Visible;

                bool builtInTab = existingMode && tab?.IsBuiltIn == true;
                bool showPosition = builtInTab && !joiningExisting;
                ExistingTabPositionPanel.Visibility = showPosition ? Visibility.Visible : Visibility.Collapsed;
                PositionBeforeRadio.IsEnabled = showPosition;
                PositionAfterRadio.IsEnabled = showPosition;
                PositionEndRadio.IsEnabled = showPosition;
                AnchorGroupCombo.IsEnabled = showPosition && AdvancedAnchorCheck.IsChecked != true && PositionEndRadio.IsChecked != true;
                AnchorGroupCombo.Visibility = showPosition && PositionEndRadio.IsChecked != true
                    ? Visibility.Visible
                    : Visibility.Collapsed;
                AdvancedAnchorCheck.IsEnabled = showPosition && PositionEndRadio.IsChecked != true;
                AdvancedAnchorBox.IsEnabled = showPosition && AdvancedAnchorCheck.IsChecked == true
                    && PositionEndRadio.IsChecked != true;
                AdvancedAnchorBox.Visibility = AdvancedAnchorCheck.IsChecked == true && showPosition
                    ? Visibility.Visible
                    : Visibility.Collapsed;
                CustomTabHint.Visibility = existingMode && tab != null && !tab.IsBuiltIn && !joiningExisting
                    ? Visibility.Visible
                    : Visibility.Collapsed;

                if (showPosition)
                    PopulateAnchorGroups(tab);
            }
            finally
            {
                _updatingPositionControls = false;
            }
        }

        private void PopulateAnchorGroups(TabPickerEntry tab)
        {
            string selectedId = (AnchorGroupCombo.SelectedItem as RibbonGroupRef)?.IdMso;
            AnchorGroupCombo.Items.Clear();
            foreach (var group in tab.BuiltIn.Groups)
                AnchorGroupCombo.Items.Add(group);

            int index = -1;
            for (int i = 0; i < AnchorGroupCombo.Items.Count; i++)
            {
                if (AnchorGroupCombo.Items[i] is RibbonGroupRef group && group.IdMso == selectedId)
                {
                    index = i;
                    break;
                }
            }

            if (index < 0 && AnchorGroupCombo.Items.Count > 0)
                index = 0;
            AnchorGroupCombo.SelectedIndex = index;
        }

        private void AdvancedAnchorCheck_Changed(object sender, RoutedEventArgs e)
        {
            RefreshPositionControlsState();
        }

        private void Position_Changed(object sender, RoutedEventArgs e)
        {
            RefreshPositionControlsState();
        }

        private void IconMode_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized)
                return;

            PickOfficeIconButton.IsEnabled = OfficeIconRadio.IsChecked == true;
            PickImageButton.IsEnabled = CustomIconRadio.IsChecked == true;
        }

        private void TextOnlyGroupCheck_Changed(object sender, RoutedEventArgs e)
        {
            bool textOnly = TextOnlyGroupCheck.IsChecked == true;
            ButtonLabelBox.IsEnabled = !textOnly;
            ScreenTipBox.IsEnabled = !textOnly;
            SuperTipBox.IsEnabled = !textOnly;
            LargeSizeCheck.IsEnabled = !textOnly;
            MacroCombo.IsEnabled = !textOnly;
            RefreshMacrosButton.IsEnabled = !textOnly;
            OfficeIconRadio.IsEnabled = !textOnly;
            CustomIconRadio.IsEnabled = !textOnly;
            PickOfficeIconButton.IsEnabled = !textOnly && OfficeIconRadio.IsChecked == true;
            PickImageButton.IsEnabled = !textOnly && CustomIconRadio.IsChecked == true;
            IconValueText.IsEnabled = !textOnly;
            ImagePathText.IsEnabled = !textOnly;
            IconPreview.Opacity = textOnly ? 0.45 : 1;
        }

        private void PickOfficeIconButton_Click(object sender, RoutedEventArgs e)
        {
            var picker = new ImageMsoPickerWindow(_selectedImageMso) { Owner = this };
            if (picker.ShowDialog() != true)
                return;

            _selectedImageMso = picker.SelectedImageMso;
            _selectedImagePath = null;
            IconValueText.Text = "imageMso: " + _selectedImageMso;
            ImagePathText.Text = LanguageService.Get("icon.imageFormats");
            IconPreview.Source = null;
        }

        private void PickImageButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = LanguageService.Get("target.imageFilter"),
                Title = LanguageService.Get("target.pickImageTitle"),
            };
            if (dialog.ShowDialog(this) != true)
                return;

            try
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.UriSource = new Uri(dialog.FileName);
                image.EndInit();

                _selectedImagePath = dialog.FileName;
                _selectedImageMso = null;
                IconPreview.Source = image;
                IconValueText.Text = LanguageService.Get("icon.customSelected");
                ImagePathText.Text = Path.GetFileName(dialog.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, string.Format(LanguageService.Get("target.imageReadError"), ex.Message),
                    LanguageService.Get("target.invalidImageTitle"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void RefreshMacrosButton_Click(object sender, RoutedEventArgs e)
        {
            if (_rescanMacros == null)
                return;

            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            try
            {
                _macros = _rescanMacros();
                PopulateMacroCombo();
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void PrefillFromExisting(string editingGroupLabel)
        {
            if (_editingButton != null)
            {
                ButtonLabelBox.Text = _editingButton.Label;
                ScreenTipBox.Text = _editingButton.ScreenTip;
                SuperTipBox.Text = _editingButton.SuperTip;
                LargeSizeCheck.IsChecked = _editingButton.LargeSize;
                GroupLabelBox.Text = editingGroupLabel;

                if (_editingButton.Icon?.Kind == IconKind.Custom)
                {
                    CustomIconRadio.IsChecked = true;
                    _selectedImagePath = !string.IsNullOrEmpty(_editingButton.Icon.SourceImagePath)
                        ? _editingButton.Icon.SourceImagePath
                        : _editingButton.Icon.ExtractedPreviewPath;
                    ImagePathText.Text = !string.IsNullOrEmpty(_editingButton.Icon.SourceImagePath)
                        ? Path.GetFileName(_selectedImagePath)
                        : LanguageService.Get("target.embeddedImage");
                    IconValueText.Text = LanguageService.Get("icon.customSelected");
                    LoadPreview(_selectedImagePath);
                }
                else if (_editingButton.Icon?.Kind == IconKind.ImageMso)
                {
                    OfficeIconRadio.IsChecked = true;
                    _selectedImageMso = _editingButton.Icon.ImageMso;
                    IconValueText.Text = "imageMso: " + _selectedImageMso;
                }
            }
            else if (_editingItem != null)
            {
                GroupLabelBox.Text = _editingItem.GroupLabel;
                LabelControlsBox.Text = string.Join(Environment.NewLine,
                    (_editingItem.LabelControls ?? new List<LabelControlInfo>()).Select(label => label.Label));
                TextOnlyGroupCheck.IsChecked = true;
            }

            if (_editingItem != null && _editingButton != null)
            {
                LabelControlsBox.Text = string.Join(Environment.NewLine,
                    (_editingItem.LabelControls ?? new List<LabelControlInfo>()).Select(label => label.Label));
            }

            if (_editingTarget != null)
            {
                ExistingTabRadio.IsChecked = true;
                for (int i = 0; i < TabCombo.Items.Count; i++)
                {
                    if (!(TabCombo.Items[i] is TabPickerEntry entry))
                        continue;
                    bool match = _editingTarget.Kind == TargetKind.NewTab
                        ? !entry.IsBuiltIn && entry.Id == _editingTarget.NewTabId
                        : entry.IsBuiltIn && entry.Id == _editingTarget.ExistingTabIdMso;
                    if (match)
                    {
                        TabCombo.SelectedIndex = i;
                        break;
                    }
                }

                if (_editingTarget.Kind == TargetKind.ExistingTab)
                {
                    bool found = false;
                    for (int i = 0; i < AnchorGroupCombo.Items.Count; i++)
                    {
                        if (AnchorGroupCombo.Items[i] is RibbonGroupRef group
                            && group.IdMso == _editingTarget.AnchorGroupIdMso)
                        {
                            AnchorGroupCombo.SelectedIndex = i;
                            found = true;
                            break;
                        }
                    }

                    if (!found && !string.IsNullOrEmpty(_editingTarget.AnchorGroupIdMso))
                    {
                        AdvancedAnchorCheck.IsChecked = true;
                        AdvancedAnchorBox.Text = _editingTarget.AnchorGroupIdMso;
                    }

                    PositionBeforeRadio.IsChecked = _editingTarget.Position == AnchorPosition.Before;
                    PositionAfterRadio.IsChecked = _editingTarget.Position == AnchorPosition.After;
                    PositionEndRadio.IsChecked = _editingTarget.Position == AnchorPosition.End;
                }
                else
                {
                    NewTabLabelBox.Text = _editingTarget.NewTabLabel;
                    SelectNewTabAnchor(_editingTarget.NewTabInsertBeforeMso);
                }
            }

            if (_allowTextOnlyGroup && _editingItem == null)
                TextOnlyGroupCheck.IsChecked = true;
            UpdateTargetMode();
            IconMode_Changed(this, null);
            RefreshPositionControlsState();
            TextOnlyGroupCheck_Changed(this, null);
        }

        private void LoadPreview(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            try
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.UriSource = new Uri(path);
                image.EndInit();
                IconPreview.Source = image;
            }
            catch (Exception)
            {
                IconPreview.Source = null;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            string label = ButtonLabelBox.Text?.Trim();
            bool textOnly = TextOnlyGroupCheck.IsChecked == true;
            var labelControls = BuildLabelControls();
            if (textOnly && labelControls.Count == 0)
            {
                ShowValidation(LanguageService.Get("target.validationTextOnly"));
                LabelControlsBox.Focus();
                return;
            }
            if (!textOnly && string.IsNullOrEmpty(label))
            {
                ShowValidation(LanguageService.Get("target.validationLabel"));
                ButtonLabelBox.Focus();
                return;
            }

            bool joiningExisting = ExistingTabRadio.IsChecked == true && ExistingGroupRadio.IsChecked == true;
            var selectedGroup = ExistingGroupCombo.SelectedItem as GroupPickerEntry;
            if (joiningExisting && selectedGroup == null)
            {
                ShowValidation(LanguageService.Get("target.validationChooseGroup"));
                return;
            }

            string groupLabel = joiningExisting ? selectedGroup.Item.GroupLabel : GroupLabelBox.Text?.Trim();
            if (string.IsNullOrEmpty(groupLabel))
            {
                ShowValidation(LanguageService.Get("target.validationGroup"));
                GroupLabelBox.Focus();
                return;
            }

            if (!textOnly && OfficeIconRadio.IsChecked == true && string.IsNullOrEmpty(_selectedImageMso))
            {
                ShowValidation(LanguageService.Get("target.validationIcon"));
                return;
            }
            if (!textOnly && CustomIconRadio.IsChecked == true && string.IsNullOrEmpty(_selectedImagePath))
            {
                ShowValidation(LanguageService.Get("target.validationImage"));
                return;
            }

            var target = BuildTarget(joiningExisting, selectedGroup);
            if (target == null)
                return;

            ButtonInfo button = null;
            if (!textOnly)
            {
                button = new ButtonInfo
                {
                    Id = _editingButton?.Id ?? "rc_btn_" + Guid.NewGuid().ToString("N"),
                    Label = label,
                    ScreenTip = ScreenTipBox.Text?.Trim(),
                    SuperTip = SuperTipBox.Text?.Trim(),
                    LargeSize = LargeSizeCheck.IsChecked == true,
                    Icon = OfficeIconRadio.IsChecked == true
                        ? new IconInfo { Kind = IconKind.ImageMso, ImageMso = _selectedImageMso }
                        : new IconInfo
                        {
                            Kind = IconKind.Custom,
                            SourceImagePath = _selectedImagePath,
                            ExtractedPreviewPath = _editingButton?.Icon?.ExtractedPreviewPath,
                        },
                    MacroName = (MacroCombo.SelectedItem as MacroEntry)?.Name,
                };
            }

            ResultTarget = target;
            ResultButton = button;
            ResultGroupLabel = groupLabel;
            ResultLabelControls = labelControls;
            ResultExistingGroupItem = joiningExisting ? selectedGroup.Item : null;
            DialogResult = true;
        }

        private List<LabelControlInfo> BuildLabelControls()
        {
            var existing = _editingItem?.LabelControls ?? new List<LabelControlInfo>();
            return (LabelControlsBox.Text ?? string.Empty)
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
                .Select((text, index) => new { Text = text.Trim(), Index = index })
                .Where(item => !string.IsNullOrWhiteSpace(item.Text))
                .Select(item => new LabelControlInfo
                {
                    Id = item.Index < existing.Count
                        ? existing[item.Index].Id
                        : "rc_label_" + Guid.NewGuid().ToString("N"),
                    Label = item.Text,
                })
                .ToList();
        }

        private TargetInfo BuildTarget(bool joiningExisting, GroupPickerEntry selectedGroup)
        {
            if (joiningExisting)
                return selectedGroup.Item.Target;

            if (NewTabRadio.IsChecked == true)
            {
                string label = NewTabLabelBox.Text?.Trim();
                if (string.IsNullOrEmpty(label))
                {
                    ShowValidation(LanguageService.Get("target.validationNewTab"));
                    NewTabLabelBox.Focus();
                    return null;
                }

                var anchorTab = NewTabAnchorCombo.SelectedItem as RibbonTabRef;
                return new TargetInfo
                {
                    Kind = TargetKind.NewTab,
                    NewTabLabel = label,
                    NewTabInsertBeforeMso = anchorTab?.IdMso,
                };
            }

            if (!(TabCombo.SelectedItem is TabPickerEntry tab))
            {
                ShowValidation(LanguageService.Get("target.validationChooseTab"));
                return null;
            }

            if (!tab.IsBuiltIn)
            {
                return new TargetInfo
                {
                    Kind = TargetKind.NewTab,
                    NewTabId = tab.Id,
                    NewTabLabel = tab.Label,
                    NewTabInsertBeforeMso = (NewTabAnchorCombo.SelectedItem as RibbonTabRef)?.IdMso,
                };
            }

            var target = new TargetInfo
            {
                Kind = TargetKind.ExistingTab,
                ExistingTabIdMso = tab.Id,
                ExistingTabLabel = tab.Label,
            };

            if (PositionEndRadio.IsChecked == true)
            {
                target.Position = AnchorPosition.End;
                return target;
            }

            string anchorId;
            string anchorLabel;
            if (AdvancedAnchorCheck.IsChecked == true)
            {
                anchorId = AdvancedAnchorBox.Text?.Trim();
                anchorLabel = anchorId;
                if (string.IsNullOrEmpty(anchorId))
                {
                    ShowValidation(LanguageService.Get("target.validationAdvanced"));
                    AdvancedAnchorBox.Focus();
                    return null;
                }
            }
            else if (AnchorGroupCombo.SelectedItem is RibbonGroupRef group)
            {
                anchorId = group.IdMso;
                anchorLabel = group.DisplayLabel;
            }
            else
            {
                ShowValidation(LanguageService.Get("target.validationAnchor"));
                return null;
            }

            target.AnchorGroupIdMso = anchorId;
            target.AnchorGroupLabel = anchorLabel;
            target.Position = PositionBeforeRadio.IsChecked == true ? AnchorPosition.Before : AnchorPosition.After;
            return target;
        }

        private void ShowValidation(string message)
        {
            MessageBox.Show(this, message, LanguageService.Get("common.validationTitle"), MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}
