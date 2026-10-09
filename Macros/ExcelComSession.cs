using System;
using System.Runtime.InteropServices;

namespace RibbonXmlEditor.Macros
{
    /// <summary>
    /// Mở 1 Excel instance ẩn qua COM automation để đọc/ghi VBA project của đúng 1 file.
    /// File phải đang ĐÓNG trong Excel của người dùng khi RibbonXmlEditor thao tác trên nó
    /// (giống hệt yêu cầu của Office RibbonX Editor gốc).
    /// </summary>
    public sealed class ExcelComSession : IDisposable
    {
        private dynamic _app;
        private dynamic _workbook;
        private IDisposable _messageFilterRegistration;

        public string Error { get; private set; }

        public dynamic Workbook => _workbook;

        public bool Open(string filePath)
        {
            try
            {
                _messageFilterRegistration = ExcelComMessageFilter.Register();
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                    throw new InvalidOperationException("Microsoft Excel chưa được cài đặt hoặc chưa đăng ký COM automation.");

                _app = Activator.CreateInstance(excelType);
                _app.Visible = false;
                _app.DisplayAlerts = false;
                _app.ScreenUpdating = false;
                _workbook = _app.Workbooks.Open(filePath, false, false);
            }
            catch (Exception ex)
            {
                Error = IsCallRejected(ex)
                    ? "Excel đang bận và từ chối lời gọi tự động hóa. Đóng các hộp thoại Excel đang mở, " +
                      "đợi Excel hoàn tất thao tác hiện tại rồi thử lại. Nếu workbook đang mở trong Excel, hãy đóng workbook đó trước.\n\n" +
                      "Chi tiết: " + ex.Message
                    : "Không mở được file bằng Excel (COM automation). Đảm bảo file đang đóng trong mọi cửa sổ Excel khác " +
                      "trước khi chỉnh sửa bằng Ribbon Studio.\n\nChi tiết: " + ex.Message;
                Close();
                return false;
            }

            try
            {
                var probe = _workbook.VBProject.VBComponents.Count;
            }
            catch (Exception ex)
            {
                if (IsCallRejected(ex))
                {
                    Error = "Excel đang bận và từ chối truy cập VBA Project. Đóng các hộp thoại Excel đang mở, " +
                            "đợi Excel hoàn tất thao tác hiện tại, xác nhận workbook đã đóng trong các cửa sổ Excel khác rồi thử lại.\n\n" +
                            "Chi tiết: " + ex.Message;
                }
                else
                {
                    Error = "Không đọc được VBA project. Kiểm tra trong Excel: File → Options → Trust Center → " +
                            "Trust Center Settings → Macro Settings và bật \"Trust access to the VBA project object model\". " +
                            "Nếu tùy chọn này đã bật, kiểm tra chính sách bảo mật của tổ chức hoặc trạng thái khóa/bảo vệ của VBA project.\n\n" +
                            "Chi tiết: " + ex.Message;
                }
                Close();
                return false;
            }

            return true;
        }

        public bool SaveAndClose(out string error)
        {
            error = null;
            try
            {
                _workbook?.Save();
                return true;
            }
            catch (Exception ex)
            {
                error = "Không lưu được thay đổi macro vào file: " + ex.Message;
                return false;
            }
            finally
            {
                Close();
            }
        }

        public void Close()
        {
            try { if (_workbook != null) _workbook.Close(false); } catch { /* best effort */ }
            _workbook = null;

            try { _app?.Quit(); } catch { /* best effort */ }
            _app = null;

            _messageFilterRegistration?.Dispose();
            _messageFilterRegistration = null;

        }

        public void Dispose() => Close();

        private static bool IsCallRejected(Exception exception)
        {
            for (var current = exception; current != null; current = current.InnerException)
            {
                if (current.HResult == unchecked((int)0x80010001)
                    || current.HResult == unchecked((int)0x8001010A))
                    return true;
            }

            return false;
        }
    }
}
