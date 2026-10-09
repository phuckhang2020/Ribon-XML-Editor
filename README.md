# Ribon Studio

Ứng dụng Windows để tùy chỉnh Ribbon bên trong workbook Excel `.xlsm` và add-in `.xlam`: sắp xếp nút theo tab/group, gán macro VBA, chọn biểu tượng Office hoặc ảnh riêng và chỉnh tooltip.

Ngoài nút macro, có thể bật `startFromScratch`, chọn tab Excel cần hiển thị, đặt vị trí tab mới và thêm nhãn hướng dẫn bằng các dòng `labelControl`.
Tùy chỉnh Ribbon được nhúng trong gói OOXML của đúng workbook đang mở bằng ứng dụng; không sửa Ribbon toàn cục hoặc các workbook khác.
Giao diện hỗ trợ tiếng Việt và tiếng Anh; nhấn biểu tượng bánh răng trong cửa sổ chính để đổi ngôn ngữ. Lựa chọn được lưu riêng trong hồ sơ người dùng Windows, không ghi vào workbook.

Khi chọn **Tạo tab mới**, đặt tên tab và chọn tab Excel đứng trước nó; tên group được nhập riêng bên dưới. Với tab có sẵn, có thể tạo group mới, thêm vào group tùy chỉnh hiện có và chọn vị trí group trong tab đó. Dùng **Thêm nhãn** để tạo hoặc chỉnh sửa group chỉ chứa văn bản hướng dẫn.

## Cấu trúc mã nguồn

- `UI/` — cửa sổ và giao diện WPF.
- `Ribbon/` — catalog tab Excel, đọc và tạo Ribbon XML.
- `Packaging/` — đọc/ghi cấu hình bên trong gói workbook OOXML.
- `Macros/` — tích hợp Excel COM, dò macro và tạo callback.
- `Config/` — mô hình dữ liệu cấu hình Ribbon.
- `Localization/` — nội dung giao diện tiếng Việt/tiếng Anh và lưu lựa chọn ngôn ngữ.
- `Assets/` — tài nguyên giao diện, bao gồm icon ứng dụng.

Lịch sử thay đổi được ghi trong [CHANGELOG.md](./CHANGELOG.md).

## Tải ứng dụng

Tải gói Windows self-contained tại [Ribon Studio v1.0.0](https://github.com/phuckhang2020/Ribon-XML-Editor/releases/tag/v1.0.0). Giải nén `RibonStudio-v1.0.0-win-x64.zip` và chạy `RibbonXmlEditor.exe`. Ứng dụng vẫn yêu cầu Microsoft Excel desktop.

## Yêu cầu

- Windows 10/11
- Microsoft Excel desktop đã cài đặt
- .NET 10 Desktop Runtime (nếu chạy bản framework-dependent)
- Bật **Trust access to the VBA project object model** trong Excel Trust Center để quét macro và tạo callback

Workbook cần đóng trong mọi cửa sổ Excel khác khi mở/chỉnh sửa/lưu bằng ứng dụng. Sau khi lưu, mở lại workbook trong Excel để kiểm tra Ribbon.

## Build

```powershell
dotnet build RibbonXmlEditor.csproj
```

## Publish

Bản framework-dependent:

```powershell
dotnet publish RibbonXmlEditor.csproj -c Release -r win-x64
```

Bản self-contained (không cần cài riêng .NET Desktop Runtime trên máy đích):

```powershell
dotnet publish RibbonXmlEditor.csproj -c Release -r win-x64 --self-contained true
```

Ứng dụng dùng Excel COM automation; vì vậy publish self-contained không thay thế yêu cầu cài Excel desktop.
