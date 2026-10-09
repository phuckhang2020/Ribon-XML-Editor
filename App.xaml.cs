using System;
using System.Windows;
using System.Windows.Threading;
using RibbonXmlEditor.Localization;

namespace RibbonXmlEditor
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            DispatcherUnhandledException += OnDispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            try
            {
                LanguageService.Initialize();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể đọc cài đặt ngôn ngữ đã lưu.\n\nChi tiết: " + ex.Message,
                    "Ribon Studio", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            base.OnStartup(e);
        }

        private static void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            ShowUnhandledError(e.Exception);
            e.Handled = true;
        }

        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            ShowUnhandledError(e.ExceptionObject as Exception);
        }

        private static void ShowUnhandledError(Exception exception)
        {
            MessageBox.Show(
                exception == null ? "Đã xảy ra lỗi ngoài dự kiến." : "Đã xảy ra lỗi ngoài dự kiến:\n\n" + exception,
                "Ribon Studio - Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
