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

## Tải và chạy ứng dụng

**Bản phát hành mới nhất: Ribon Studio v1.0.0**

- [Tải ứng dụng Windows 64-bit](https://github.com/phuckhang2020/Ribon-XML-Editor/releases/download/v1.0.0/RibonStudio-v1.0.0-win-x64.zip) (`RibonStudio-v1.0.0-win-x64.zip`)
- [Tải SHA-256 checksum](https://github.com/phuckhang2020/Ribon-XML-Editor/releases/download/v1.0.0/RibonStudio-v1.0.0-win-x64.zip.sha256)
- [Xem ghi chú phát hành](https://github.com/phuckhang2020/Ribon-XML-Editor/releases/tag/v1.0.0)

1. Tải file ZIP và giải nén **toàn bộ nội dung** vào một thư mục trên máy Windows 10/11 64-bit.
2. Mở thư mục vừa giải nén và chạy `RibbonXmlEditor.exe`. Đây là bản self-contained, không cần cài riêng .NET Runtime.
3. Microsoft Excel desktop phải được cài đặt để ứng dụng kết nối Excel, quét macro và tạo callback.
4. Để quét macro và tạo callback VBA, bật **Trust access to the VBA project object model** trong Excel Trust Center.

Có thể kiểm tra SHA-256 của file tải về bằng PowerShell:

```powershell
Get-FileHash .\RibonStudio-v1.0.0-win-x64.zip -Algorithm SHA256
```

Đối chiếu giá trị `Hash` với nội dung file `.sha256` đi kèm release.

Trước khi mở, chỉnh sửa hoặc lưu workbook, hãy đóng workbook đó trong các cửa sổ Excel khác. Sau khi lưu, mở lại workbook trong Excel để xem Ribbon đã cập nhật. Tùy chỉnh Ribbon chỉ được nhúng trong workbook đã chọn.

## Build

Yêu cầu môi trường phát triển: Windows và .NET 10 SDK. Excel desktop cần có mặt để chạy các tính năng tích hợp Excel COM.

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
