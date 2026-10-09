using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace RibbonXmlEditor.Localization
{
    public static class LanguageService
    {
        private const string Vietnamese = "vi";
        private const string English = "en";
        private static readonly string SettingsDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RibonStudio");
        private static readonly string SettingsPath = Path.Combine(SettingsDirectory, "language.txt");
        private static readonly Dictionary<string, string> VietnameseText = CreateVietnameseText();
        private static readonly Dictionary<string, string> EnglishText = CreateEnglishText();

        public static string CurrentLanguage { get; private set; } = Vietnamese;
        public static event EventHandler LanguageChanged;

        public static void Initialize()
        {
            if (!File.Exists(SettingsPath))
                return;

            string language = File.ReadAllText(SettingsPath).Trim();
            if (language == Vietnamese || language == English)
            {
                CurrentLanguage = language;
                return;
            }

            throw new InvalidDataException("The saved language preference is not supported.");
        }

        public static string Get(string key)
        {
            var dictionary = CurrentLanguage == English ? EnglishText : VietnameseText;
            return dictionary.TryGetValue(key, out string value) ? value : key;
        }

        public static void SetLanguage(string language)
        {
            if (language != Vietnamese && language != English)
                throw new ArgumentException("Unsupported language.", nameof(language));

            Directory.CreateDirectory(SettingsDirectory);
            File.WriteAllText(SettingsPath, language);
            CurrentLanguage = language;

            if (Application.Current != null)
            {
                foreach (Window window in Application.Current.Windows)
                    Apply(window);
            }
            LanguageChanged?.Invoke(null, EventArgs.Empty);
        }

        public static void Apply(DependencyObject root)
        {
            if (root == null)
                return;

            if (root is FrameworkElement element && element.Tag is string tag && tag.StartsWith("L:", StringComparison.Ordinal))
                ApplyToElement(element, Get(tag.Substring(2)));

            int childCount = VisualTreeHelper.GetChildrenCount(root);
            for (int i = 0; i < childCount; i++)
                Apply(VisualTreeHelper.GetChild(root, i));

            if (childCount == 0)
            {
                foreach (object child in LogicalTreeHelper.GetChildren(root))
                {
                    if (child is DependencyObject dependencyObject)
                        Apply(dependencyObject);
                }
            }
        }

        private static void ApplyToElement(FrameworkElement element, string text)
        {
            if (element is Window window)
                window.Title = text;
            else if (element is GroupBox groupBox)
                groupBox.Header = text;
            else if (element is TextBlock textBlock)
                textBlock.Text = text;
            else if (element is ContentControl contentControl)
                contentControl.Content = text;
            else if (element is TextBox textBox)
                textBox.ToolTip = text;
        }

        private static Dictionary<string, string> CreateVietnameseText()
        {
            return new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["app.title"] = "Ribon Studio",
                ["main.eyebrow"] = "RIBON STUDIO",
                ["main.title"] = "Tùy chỉnh Ribbon cho Excel",
                ["main.subtitle"] = "Gán macro và sắp xếp các nút Ribbon trực tiếp trong workbook.",
                ["main.open"] = "Mở workbook...",
                ["main.noWorkbook"] = "Chưa mở workbook",
                ["main.chooseWorkbook"] = "Chọn một file .xlsm hoặc .xlam để bắt đầu",
                ["main.noFile"] = "Chưa có file",
                ["main.saved"] = "✓  Đã đồng bộ",
                ["main.unsaved"] = "●  Có thay đổi chưa lưu",
                ["main.embedded"] = "Ribbon được lưu bên trong workbook đã chọn",
                ["main.buttons"] = "NÚT RIBBON",
                ["main.groups"] = "GROUP TÙY CHỈNH",
                ["main.macros"] = "MACRO ĐÃ QUÉT",
                ["main.configured"] = "Các nút đã cấu hình",
                ["main.selectionHint"] = "Chọn một nút để sửa, đổi thứ tự hoặc xóa.",
                ["main.ribbonOptions"] = "Tùy chọn Ribbon",
                ["main.moveUp"] = "↑  Lên",
                ["main.moveDown"] = "↓  Xuống",
                ["main.edit"] = "Sửa",
                ["main.delete"] = "Xóa",
                ["main.addTextGroup"] = "+  Thêm nhãn",
                ["main.addButton"] = "+  Thêm nút",
                ["main.search"] = "Tìm nhanh",
                ["main.searchHint"] = "Tìm theo tên nút, group, vị trí hoặc macro",
                ["main.resultCount"] = "{0} mục",
                ["main.emptyTitle"] = "Chưa có nút Ribbon nào",
                ["main.emptyDescription"] = "Mở workbook, sau đó thêm nút và gán macro để tạo Ribbon riêng.",
                ["main.openExcel"] = "Mở workbook Excel",
                ["main.saveHint"] = "Lưu ý: workbook cần được đóng trong các cửa sổ Excel khác khi chỉnh sửa. Mở lại file trong Excel sau khi lưu để xem thay đổi.",
                ["main.save"] = "Lưu thay đổi vào file",
                    ["main.footer"] = "Ribon Studio  ·  Công cụ tùy chỉnh Ribbon cho Excel",
                ["main.location"] = "VỊ TRÍ TRÊN RIBBON",
                ["main.macroColumn"] = "MACRO ĐƯỢC GÁN",
                ["main.nameColumn"] = "NÚT / NỘI DUNG",
                ["main.groupColumn"] = "GROUP",
                ["main.fileFilter"] = "Tất cả file Excel (*.xls*)|*.xls*",
                ["main.openDialogTitle"] = "Chọn workbook cần chỉnh Ribbon",
                ["main.openErrorTitle"] = "Không thể mở workbook",
                ["main.openErrorDetails"] = "Không thể mở workbook.\n\nChi tiết: {0}",
                ["main.scanMacros"] = "Quét macro",
                ["main.newTabLocation"] = "Tab mới: {0}{1}",
                ["main.insertBefore"] = " · trước {0}",
                ["main.insertAfter"] = " · sau {0}",
                ["main.existingTabLocation"] = "{0} · {1}",
                ["main.endOfTab"] = "cuối tab",
                ["main.beforeGroup"] = "trước \"{0}\"",
                ["main.afterGroup"] = "sau \"{0}\"",
                ["main.deleteGroupConfirm"] = "Xóa group \"{0}\" và toàn bộ nội dung trong group?",
                ["main.deleteButtonConfirm"] = "Xóa nút \"{0}\"?",
                ["main.confirmDelete"] = "Xác nhận xóa",
                ["main.saveWarning"] = "Đảm bảo workbook đang đóng trong mọi cửa sổ Excel khác trước khi lưu.\n\nTiếp tục lưu?",
                ["main.saveTitle"] = "Lưu Ribbon vào workbook",
                ["main.saveSuccess"] = "Đã lưu Ribbon vào workbook. Mở lại file trong Excel để xem thay đổi.",
                ["main.saveSuccessTitle"] = "Lưu thành công",
                ["main.saveErrorTitle"] = "Không thể lưu workbook",
                ["main.dirtyConfirm"] = "Các thay đổi chưa được lưu. Bạn có muốn tiếp tục mà không lưu không?",
                ["main.dirtyTitle"] = "Thay đổi chưa lưu",
                ["target.noMacro"] = "Không tìm thấy macro phù hợp trong workbook này.",
                ["target.macroMissing"] = "Không tìm thấy macro \"{0}\". Có thể macro đã bị đổi tên hoặc xóa.",
                ["target.imageFilter"] = "Ảnh (*.png;*.jpg;*.jpeg;*.bmp;*.ico)|*.png;*.jpg;*.jpeg;*.bmp;*.ico",
                ["target.pickImageTitle"] = "Chọn ảnh làm biểu tượng",
                ["target.imageReadError"] = "Không thể đọc ảnh đã chọn.\n\nChi tiết: {0}",
                ["target.invalidImageTitle"] = "Ảnh không hợp lệ",
                ["target.embeddedImage"] = "Ảnh đã nhúng trong workbook",
                ["icon.customSelected"] = "Ảnh tùy chỉnh",
                ["icon.imageFormats"] = "PNG, JPG, BMP hoặc ICO",
                ["icon.notSelected"] = "Chưa chọn biểu tượng",
                ["icon.chooseFirst"] = "Chọn một biểu tượng trong danh sách.",
                ["icon.chooseTitle"] = "Chưa chọn biểu tượng",
                ["target.window"] = "Cấu hình nút Ribbon",
                ["target.addTitle"] = "Thêm nút Ribbon",
                ["target.editTitle"] = "Sửa nút Ribbon",
                ["target.editGroupTitle"] = "Sửa group hướng dẫn",
                ["target.createGroupTitle"] = "Thêm nhãn",
                ["target.addLabelHint"] = "Chọn vị trí trên Ribbon và nhập nội dung cho các nhãn.",
                ["target.editLabelHint"] = "Chọn vị trí trên Ribbon và chỉnh sửa nội dung các nhãn.",
                ["target.addLabelSave"] = "Thêm nhãn",
                ["section.addLabels"] = "Nhãn",
                ["labels.addHint"] = "Mỗi dòng sẽ được thêm thành một nhãn riêng trong group hướng dẫn.",
                ["target.addHint"] = "Chọn nơi hiển thị, macro cần chạy và nội dung người dùng sẽ thấy.",
                ["section.target"] = "1. Vị trí hiển thị",
                ["target.existing"] = "Tab có sẵn",
                ["target.new"] = "Tạo tab mới",
                ["target.tab"] = "Tab:",
                ["target.newName"] = "Tên tab:",
                ["target.newPosition"] = "Chọn vị trí trên Ribbon",
                ["target.insertBefore"] = "Chèn trước tab:",
                ["target.groupPosition"] = "Vị trí trong tab:",
                ["target.customTabHint"] = "Tab tự tạo: group mới sẽ được thêm vào cuối tab này.",
                ["group.create"] = "Tạo group mới",
                ["group.existing"] = "Thêm vào group hiện có",
                ["group.choose"] = "Group hiện có:",
                ["group.name"] = "Tên group:",
                ["position.before"] = "Trước group",
                ["position.after"] = "Sau group",
                ["position.end"] = "Cuối tab",
                ["position.advanced"] = "Nhập idMso tùy chỉnh",
                ["section.macro"] = "2. Macro",
                ["common.refresh"] = "Làm mới",
                ["macro.hint"] = "Macro không tham số cũng dùng được; ứng dụng tự tạo callback Ribbon cần thiết.",
                ["section.icon"] = "3. Biểu tượng",
                ["icon.office"] = "Biểu tượng Office (imageMso)",
                ["icon.pickOffice"] = "Chọn biểu tượng...",
                ["icon.custom"] = "Ảnh tùy chỉnh",
                ["icon.pickImage"] = "Chọn ảnh...",
                ["icon.none"] = "Chưa chọn",
                ["icon.imageTypes"] = "PNG, JPG, BMP hoặc ICO",
                ["icon.selected"] = "Ảnh tùy chỉnh",
                ["icon.embedded"] = "Ảnh đã nhúng trong workbook",
                ["section.content"] = "4. Nội dung và trợ giúp",
                ["button.name"] = "Tên nút:",
                ["button.screenTip"] = "Screen tip:",
                ["button.superTip"] = "Mô tả chi tiết:",
                ["button.large"] = "Hiển thị nút ở kích thước lớn",
                ["section.labels"] = "Nội dung chữ trong group",
                ["labels.hint"] = "Mỗi dòng tạo một labelControl riêng, giống các dòng hướng dẫn và ghi chú trong Ribbon.",
                ["labels.textOnly"] = "Chỉ tạo group hướng dẫn (không thêm nút)",
                ["target.validationLabel"] = "Nhập tên cho nút.",
                ["target.validationTextOnly"] = "Nhập ít nhất một dòng nội dung cho group hướng dẫn.",
                ["target.validationGroup"] = "Nhập tên group chứa nút này.",
                ["target.validationNewTab"] = "Nhập tên cho tab mới.",
                ["target.validationChooseTab"] = "Chọn một tab.",
                ["target.validationChooseGroup"] = "Chọn một group hiện có.",
                ["target.validationIcon"] = "Chọn một biểu tượng Office.",
                ["target.validationImage"] = "Chọn một file ảnh cho biểu tượng.",
                ["target.validationAnchor"] = "Chọn group neo, hoặc chọn \"Cuối tab\".",
                ["target.validationAdvanced"] = "Nhập idMso của group cần neo, hoặc chọn \"Cuối tab\".",
                ["common.validationTitle"] = "Kiểm tra thông tin",
                ["common.cancel"] = "Hủy",
                ["common.saveButton"] = "Lưu nút",
                ["settings.title"] = "Cài đặt",
                ["settings.heading"] = "Cài đặt",
                ["settings.language"] = "Ngôn ngữ",
                ["settings.languageHint"] = "Thay đổi ngôn ngữ hiển thị của ứng dụng.",
                ["settings.vietnamese"] = "Tiếng Việt",
                ["settings.english"] = "English",
                ["settings.save"] = "Lưu cài đặt",
                ["settings.failed"] = "Không thể lưu cài đặt ngôn ngữ.",
                ["settings.iconHint"] = "Mở cài đặt ngôn ngữ",
                ["options.title"] = "Tùy chọn Ribbon",
                ["options.hint"] = "Các thiết lập này được nhúng trong workbook đang chỉnh sửa.",
                ["options.scratch"] = "Ẩn Ribbon mặc định và chỉ hiển thị các tab được chọn",
                ["options.scratchHint"] = "Tạo customUI với startFromScratch=\"true\". Chỉ áp dụng khi mở chính workbook này.",
                ["options.builtin"] = "Tab có sẵn của Excel",
                ["options.tabsHint"] = "Bỏ chọn để ẩn tab đó. Khi bật startFromScratch, chỉ các tab được chọn sẽ xuất hiện.",
                ["common.saveOptions"] = "Lưu tùy chọn",
                ["iconPicker.title"] = "Chọn biểu tượng Office",
                ["iconPicker.heading"] = "Biểu tượng Office",
                ["iconPicker.hint"] = "Tìm theo tên imageMso. Xem trước hình thật trong Excel sau khi lưu.",
                ["iconPicker.selectedNone"] = "Chưa chọn biểu tượng",
                ["common.choose"] = "Chọn",
            };
        }

        private static Dictionary<string, string> CreateEnglishText()
        {
            var text = CreateVietnameseText();
            var english = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["app.title"] = "Ribon Studio",
                ["main.eyebrow"] = "RIBON STUDIO",
                ["main.title"] = "Customize Excel Ribbon",
                ["main.subtitle"] = "Assign macros and organize Ribbon buttons directly in your workbook.",
                ["main.open"] = "Open workbook...",
                ["main.noWorkbook"] = "No workbook open",
                ["main.chooseWorkbook"] = "Choose an .xlsm or .xlam file to get started",
                ["main.noFile"] = "No file",
                ["main.saved"] = "✓  Up to date",
                ["main.unsaved"] = "●  Unsaved changes",
                ["main.embedded"] = "Ribbon settings are embedded in the selected workbook",
                ["main.buttons"] = "RIBBON BUTTONS",
                ["main.groups"] = "CUSTOM GROUPS",
                ["main.macros"] = "MACROS FOUND",
                ["main.configured"] = "Configured Ribbon items",
                ["main.selectionHint"] = "Select an item to edit, reorder, or delete.",
                ["main.ribbonOptions"] = "Ribbon options",
                ["main.moveUp"] = "↑  Move up",
                ["main.moveDown"] = "↓  Move down",
                ["main.edit"] = "Edit",
                ["main.delete"] = "Delete",
                ["main.addTextGroup"] = "+  Add label",
                ["main.addButton"] = "+  Add button",
                ["main.search"] = "Search",
                ["main.searchHint"] = "Search by button, group, location, or macro",
                ["main.resultCount"] = "{0} items",
                ["main.emptyTitle"] = "No Ribbon items yet",
                ["main.emptyDescription"] = "Open a workbook, then add a button and assign a macro to build your Ribbon.",
                ["main.openExcel"] = "Open Excel workbook",
                ["main.saveHint"] = "Note: close the workbook in other Excel windows before editing. Reopen it in Excel after saving to see your changes.",
                ["main.save"] = "Save changes to workbook",
                ["main.footer"] = "Ribon Studio  ·  Excel Ribbon customization tool",
                ["main.location"] = "RIBBON LOCATION",
                ["main.macroColumn"] = "ASSIGNED MACRO",
                ["main.nameColumn"] = "BUTTON / CONTENT",
                ["main.groupColumn"] = "GROUP",
                ["main.fileFilter"] = "All Excel workbooks (*.xls*)|*.xls*",
                ["main.openDialogTitle"] = "Choose workbook to customize",
                ["main.openErrorTitle"] = "Could not open workbook",
                ["main.openErrorDetails"] = "Could not open workbook.\n\nDetails: {0}",
                ["main.scanMacros"] = "Scan macros",
                ["main.newTabLocation"] = "New tab: {0}{1}",
                ["main.insertBefore"] = " · before {0}",
                ["main.insertAfter"] = " · after {0}",
                ["main.existingTabLocation"] = "{0} · {1}",
                ["main.endOfTab"] = "end of tab",
                ["main.beforeGroup"] = "before \"{0}\"",
                ["main.afterGroup"] = "after \"{0}\"",
                ["main.deleteGroupConfirm"] = "Delete group \"{0}\" and all its contents?",
                ["main.deleteButtonConfirm"] = "Delete button \"{0}\"?",
                ["main.confirmDelete"] = "Confirm deletion",
                ["main.saveWarning"] = "Make sure the workbook is closed in all other Excel windows before saving.\n\nContinue saving?",
                ["main.saveTitle"] = "Save Ribbon to workbook",
                ["main.saveSuccess"] = "Ribbon was saved to the workbook. Reopen it in Excel to see the changes.",
                ["main.saveSuccessTitle"] = "Saved successfully",
                ["main.saveErrorTitle"] = "Could not save workbook",
                ["main.dirtyConfirm"] = "There are unsaved changes. Continue without saving?",
                ["main.dirtyTitle"] = "Unsaved changes",
                ["target.noMacro"] = "No compatible macros were found in this workbook.",
                ["target.macroMissing"] = "Macro \"{0}\" was not found. It may have been renamed or deleted.",
                ["target.imageFilter"] = "Images (*.png;*.jpg;*.jpeg;*.bmp;*.ico)|*.png;*.jpg;*.jpeg;*.bmp;*.ico",
                ["target.pickImageTitle"] = "Choose button image",
                ["target.imageReadError"] = "Could not read the selected image.\n\nDetails: {0}",
                ["target.invalidImageTitle"] = "Invalid image",
                ["target.embeddedImage"] = "Image embedded in workbook",
                ["icon.customSelected"] = "Custom image",
                ["icon.imageFormats"] = "PNG, JPG, BMP, or ICO",
                ["icon.notSelected"] = "No icon selected",
                ["icon.chooseFirst"] = "Select an icon from the list.",
                ["icon.chooseTitle"] = "No icon selected",
                ["target.window"] = "Ribbon button settings",
                ["target.addTitle"] = "Add Ribbon button",
                ["target.editTitle"] = "Edit Ribbon button",
                ["target.editGroupTitle"] = "Edit guide group",
                ["target.createGroupTitle"] = "Add label",
                ["target.addLabelHint"] = "Choose a Ribbon location and enter the label text.",
                ["target.editLabelHint"] = "Choose a Ribbon location and edit the label text.",
                ["target.addLabelSave"] = "Add label",
                ["section.addLabels"] = "Labels",
                ["labels.addHint"] = "Each line will be added as a separate label in the guide group.",
                ["target.addHint"] = "Choose where it appears, which macro to run, and what users will see.",
                ["section.target"] = "1. Placement",
                ["target.existing"] = "Existing tab",
                ["target.new"] = "Create new tab",
                ["target.tab"] = "Tab:",
                ["target.newName"] = "Tab name:",
                ["target.newPosition"] = "Choose position on the Ribbon",
                ["target.insertBefore"] = "Insert before tab:",
                ["target.groupPosition"] = "Group position:",
                ["target.customTabHint"] = "Custom tab: new groups are added to the end of this tab.",
                ["group.create"] = "Create a new group",
                ["group.existing"] = "Add to an existing group",
                ["group.choose"] = "Existing group:",
                ["group.name"] = "Group name:",
                ["position.before"] = "Before group",
                ["position.after"] = "After group",
                ["position.end"] = "End of tab",
                ["position.advanced"] = "Enter a custom idMso",
                ["section.macro"] = "2. Macro",
                ["common.refresh"] = "Refresh",
                ["macro.hint"] = "Macros without parameters are supported; the app creates the required Ribbon callback.",
                ["section.icon"] = "3. Icon",
                ["icon.office"] = "Office icon (imageMso)",
                ["icon.pickOffice"] = "Choose icon...",
                ["icon.custom"] = "Custom image",
                ["icon.pickImage"] = "Choose image...",
                ["icon.none"] = "No icon selected",
                ["icon.imageTypes"] = "PNG, JPG, BMP, or ICO",
                ["icon.selected"] = "Custom image",
                ["icon.embedded"] = "Image embedded in workbook",
                ["section.content"] = "4. Button content and help",
                ["button.name"] = "Button label:",
                ["button.screenTip"] = "Screen tip:",
                ["button.superTip"] = "Description:",
                ["button.large"] = "Show button using large size",
                ["section.labels"] = "Text content in group",
                ["labels.hint"] = "Each line creates a separate labelControl, suitable for Ribbon guides and notes.",
                ["labels.textOnly"] = "Create a guide group only (no button)",
                ["target.validationLabel"] = "Enter a button label.",
                ["target.validationTextOnly"] = "Enter at least one line of text for the guide group.",
                ["target.validationGroup"] = "Enter a name for this group.",
                ["target.validationNewTab"] = "Enter a name for the new tab.",
                ["target.validationChooseTab"] = "Select a tab.",
                ["target.validationChooseGroup"] = "Select an existing group.",
                ["target.validationIcon"] = "Choose an Office icon.",
                ["target.validationImage"] = "Choose an image file for the icon.",
                ["target.validationAnchor"] = "Choose an anchor group or select \"End of tab\".",
                ["target.validationAdvanced"] = "Enter the anchor group's idMso or select \"End of tab\".",
                ["common.validationTitle"] = "Check required information",
                ["common.cancel"] = "Cancel",
                ["common.saveButton"] = "Save button",
                ["settings.title"] = "Settings",
                ["settings.heading"] = "Settings",
                ["settings.language"] = "Language",
                ["settings.languageHint"] = "Change the display language of the application.",
                ["settings.vietnamese"] = "Tiếng Việt",
                ["settings.english"] = "English",
                ["settings.save"] = "Save settings",
                ["settings.failed"] = "Could not save the language preference.",
                ["settings.iconHint"] = "Open language settings",
                ["options.title"] = "Ribbon options",
                ["options.hint"] = "These settings are embedded in the workbook being edited.",
                ["options.scratch"] = "Hide the default Ribbon and show only selected tabs",
                ["options.scratchHint"] = "Creates customUI with startFromScratch=\"true\". Applies only when this workbook is opened.",
                ["options.builtin"] = "Built-in Excel tabs",
                ["options.tabsHint"] = "Clear a checkbox to hide that tab. When startFromScratch is enabled, only selected tabs are shown.",
                ["common.saveOptions"] = "Save options",
                ["iconPicker.title"] = "Choose Office icon",
                ["iconPicker.heading"] = "Office icons",
                ["iconPicker.hint"] = "Search by imageMso name. Preview the actual icon in Excel after saving.",
                ["iconPicker.selectedNone"] = "No icon selected",
                ["common.choose"] = "Choose",
            };
            foreach (var entry in english)
                text[entry.Key] = entry.Value;
            return text;
        }
    }
}
