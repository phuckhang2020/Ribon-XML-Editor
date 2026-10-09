using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RibbonXmlEditor.Macros
{
    /// <summary>
    /// Ribbon customUI bắt buộc callback đúng chữ ký "Sub X(control As IRibbonControl)".
    /// Với macro người dùng chọn mà không đúng chữ ký này (Sub thường không tham số), tool tự
    /// sinh 1 Sub wrapper đúng chữ ký trong module riêng "RibbonCustomizerCallbacks" rồi gọi
    /// macro gốc bên trong - người dùng không cần biết/sửa gì cả. Module này được ghi đè lại
    /// hoàn toàn mỗi lần Save cho khớp đúng danh sách nút hiện tại (idempotent).
    /// </summary>
    public static class RibbonWrapperWriter
    {
        public const string WrapperSuffix = "_RibbonWrapper";

        /// <summary>originalMacroNames: các tên macro (không tham số) cần có wrapper.</summary>
        public static bool EnsureWrappers(dynamic workbook, IEnumerable<string> originalMacroNames, out string error)
        {
            error = null;
            var names = new List<string>();
            foreach (var n in originalMacroNames)
                if (!string.IsNullOrWhiteSpace(n) && !names.Any(x => string.Equals(x, n, StringComparison.OrdinalIgnoreCase)))
                    names.Add(n);

            try
            {
                dynamic components = workbook.VBProject.VBComponents;
                RemoveModuleIfExists(components);

                if (names.Count == 0)
                    return true;

                dynamic module = components.Add(1); // vbext_ct_StdModule
                module.Name = WorkbookVbaMacroReader.CallbacksModuleName;

                var sb = new StringBuilder();
                sb.AppendLine("' Sinh tự động bởi RibbonXmlEditor - đừng sửa tay, sẽ bị ghi đè mỗi lần Save.");
                foreach (var name in names)
                {
                    sb.AppendLine($"Sub {WrapperName(name)}(control As IRibbonControl)");
                    sb.AppendLine($"    Call {name}");
                    sb.AppendLine("End Sub");
                    sb.AppendLine();
                }

                module.CodeModule.AddFromString(sb.ToString());
                return true;
            }
            catch (Exception ex)
            {
                error = "Không ghi được macro wrapper vào workbook: " + ex.Message;
                return false;
            }
        }

        public static string WrapperName(string originalMacroName) => originalMacroName + WrapperSuffix;

        private static void RemoveModuleIfExists(dynamic components)
        {
            try
            {
                foreach (dynamic component in components)
                {
                    string name;
                    try { name = (string)component.Name; } catch { continue; }
                    if (string.Equals(name, WorkbookVbaMacroReader.CallbacksModuleName, StringComparison.OrdinalIgnoreCase))
                    {
                        components.Remove(component);
                        break;
                    }
                }
            }
            catch { /* best effort */ }
        }
    }
}
