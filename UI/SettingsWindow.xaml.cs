using System;
using System.Windows;
using RibbonXmlEditor.Localization;

namespace RibbonXmlEditor.UI
{
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
            LanguageService.Apply(this);
            VietnameseRadio.IsChecked = LanguageService.CurrentLanguage == "vi";
            EnglishRadio.IsChecked = LanguageService.CurrentLanguage == "en";
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            string language = EnglishRadio.IsChecked == true ? "en" : "vi";
            try
            {
                LanguageService.SetLanguage(language);
                DialogResult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, LanguageService.Get("settings.failed") + "\n\n" + ex.Message,
                    LanguageService.Get("settings.title"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
