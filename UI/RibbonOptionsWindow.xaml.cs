using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using RibbonXmlEditor.Config;
using RibbonXmlEditor.Localization;
using RibbonXmlEditor.Ribbon;

namespace RibbonXmlEditor.UI
{
    public partial class RibbonOptionsWindow : Window
    {
        private readonly List<TabVisibilityRow> _tabRows;
        private readonly List<BuiltInTabVisibility> _unrecognizedSettings;

        public bool StartFromScratch => StartFromScratchCheck.IsChecked == true;

        public List<BuiltInTabVisibility> BuiltInTabs { get; private set; }

        public RibbonOptionsWindow(RibbonConfig config)
        {
            InitializeComponent();
            LanguageService.Apply(this);
            StartFromScratchCheck.IsChecked = config.StartFromScratch;

            var configured = config.BuiltInTabs ?? new List<BuiltInTabVisibility>();
            _unrecognizedSettings = configured
                .Where(setting => !BuiltInRibbonCatalog.Tabs.Any(tab => tab.IdMso == setting.IdMso))
                .Select(setting => new BuiltInTabVisibility { IdMso = setting.IdMso, Visible = setting.Visible })
                .ToList();

            _tabRows = BuiltInRibbonCatalog.Tabs.Select(tab =>
            {
                var setting = configured.FirstOrDefault(item => item.IdMso == tab.IdMso);
                return new TabVisibilityRow(tab, setting?.Visible ?? true);
            }).ToList();

            foreach (var row in _tabRows)
            {
                var checkBox = new CheckBox
                {
                    Content = row.Tab.DisplayLabel,
                    IsChecked = row.Visible,
                    Tag = row,
                    Margin = new Thickness(2, 8, 2, 8),
                    FontSize = 14,
                };
                TabsPanel.Children.Add(checkBox);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            BuiltInTabs = _unrecognizedSettings
                .Concat(_tabRows.Select(row => new BuiltInTabVisibility
                {
                    IdMso = row.Tab.IdMso,
                    Visible = TabsPanel.Children.OfType<CheckBox>()
                        .First(checkBox => checkBox.Tag == row).IsChecked == true,
                }))
                .ToList();
            DialogResult = true;
        }

        private sealed class TabVisibilityRow
        {
            public TabVisibilityRow(RibbonTabRef tab, bool visible)
            {
                Tab = tab;
                Visible = visible;
            }

            public RibbonTabRef Tab { get; }
            public bool Visible { get; set; }
        }
    }
}
