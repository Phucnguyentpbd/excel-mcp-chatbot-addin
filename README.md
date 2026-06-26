# Excel MCP Chatbot Add-in

Add-in Excel dạng task pane để chat trực tiếp với workbook đang mở và gọi các tool từ `excel-mcp-server`.

## Hình ảnh minh họa

Chatbot chạy trực tiếp trong Excel, có thể đọc workbook hiện tại, tạo pivot/chart và phản hồi ngay trong task pane.

![Excel MCP Chatbot thao tác workbook, pivot table và chart](docs/images/excel-mcp-chatbot-workbook.png)

## Tính năng chính

- Mở chatbot ngay trong Excel bằng nút `Open Chatbot` trên ribbon.
- Đọc context workbook hiện tại qua Office.js: workbook, sheet, active cell, selected range.
- Kết nối MCP Excel tools, gồm các tool thao tác file Excel và `run_python` để xử lý linh hoạt.
- Cấu hình provider/model trực tiếp trong giao diện chatbot, không cần sửa file config thủ công.
- Hỗ trợ nhiều provider OpenAI-compatible như 9Router, Gemini, OpenRouter, GLM, Alibaba/Qwen, DeepSeek, Groq...
- Dropdown model tự lấy model từ các provider đã cấu hình; riêng 9Router chỉ hiện combo routing như `9router/FREE`, `9router/coder`, `9router/codex`.
- Paste screenshot trực tiếp vào ô chat bằng `Ctrl+V`, không cần lưu file rồi upload.
- Ảnh gửi đi được gắn theo đúng message user trong lịch sử hội thoại.
- Có session chat local, nút `New`, `Sessions`, copy message, edit lại message user.
- Server local chạy HTTPS tại `https://localhost:3100`.
- Có script autostart để server tự chạy khi đăng nhập Windows.

## Cấu trúc nhanh

```text
excel-chatbot-addin/
  public/                    Giao diện task pane Office Add-in
  server/                    Express server, provider store, chat service
  scripts/                   Script smoke test MCP
  manifest.xml               Office Add-in manifest
  setup-excel-sideload.ps1   Cấu hình trusted shared-folder catalog cho Excel
  start-excel-chatbot.ps1    Chạy server nền tại localhost:3100
  stop-excel-chatbot.ps1     Dừng server nền
  install-autostart-task.ps1 Tạo Scheduled Task tự chạy server khi đăng nhập
```

## Yêu cầu

- Windows desktop Excel.
- Node.js 20+.
- Repo `excel-mcp-server` nằm cùng workspace:

```text
C:\Users\PHUC\Documents\Codex\2026-06-24\se\work\excel-mcp-server
```

- Office localhost certificate đã được trust. Nếu thiếu cert, chạy:

```powershell
npx office-addin-dev-certs install
```

## Cài đặt

```powershell
cd C:\Users\PHUC\Documents\Codex\2026-06-24\se\work\excel-chatbot-addin
npm install
```

Nếu Excel báo add-in bị chặn do certificate hết hạn hoặc không hợp lệ:

```powershell
.\renew-office-cert.ps1
.\start-excel-chatbot.ps1
```

## Chạy server

Chạy nền bằng helper script:

```powershell
cd C:\Users\PHUC\Documents\Codex\2026-06-24\se\work\excel-chatbot-addin
.\start-excel-chatbot.ps1
```

Kiểm tra server:

```powershell
curl.exe -k https://localhost:3100/health
```

Dừng server:

```powershell
.\stop-excel-chatbot.ps1
```

## Tự chạy sau khi restart máy

Đã có script tạo Windows Scheduled Task:

```powershell
cd C:\Users\PHUC\Documents\Codex\2026-06-24\se\work\excel-chatbot-addin
.\install-autostart-task.ps1
```

Task được tạo tên:

```text
ExcelMcpChatbotServer
```

Sau khi đăng nhập Windows, server sẽ tự chạy nền. Chỉ cần mở Excel rồi bấm `Open Chatbot`.

## Sideload vào Excel

Chạy script cấu hình trusted shared-folder catalog:

```powershell
cd C:\Users\PHUC\Documents\Codex\2026-06-24\se\work\excel-chatbot-addin
.\setup-excel-sideload.ps1
```

Sau đó trong Excel:

1. Restart Excel.
2. Mở workbook bất kỳ.
3. Vào `Home > Add-ins > Advanced`.
4. Chọn `SHARED FOLDER`.
5. Chọn `Excel MCP Chatbot` rồi bấm `Add`.
6. Dùng nút `Open Chatbot` trong group `Excel AI`.

## Cấu hình provider/model

Trong task pane:

1. Bấm `Provider`.
2. Chọn provider hoặc nhập custom provider.
3. Chọn API mode:
   - `OpenAI Responses API`: chỉ dùng cho OpenAI chính thức.
   - `OpenAI-compatible Chat Completions`: dùng cho Gemini, OpenRouter, Alibaba/Qwen, GLM, DeepSeek, Groq, 9Router...
4. Nhập `Base URL`, `API key`, `Model`.
5. Bấm `Save`.

Lưu ý:

- API key thật chỉ nhập trong giao diện local, không commit vào repo.
- Nếu ô key hiển thị dạng có `...` hoặc `*`, đó là key đã che, app sẽ không dùng chuỗi che đó để gọi API.
- Với Gemini nên dùng endpoint OpenAI-compatible:

```text
https://generativelanguage.googleapis.com/v1beta/openai
```

## Cách dùng chat

- Nhập yêu cầu vào ô chat rồi bấm `Send`.
- Paste ảnh/screenshot trực tiếp bằng `Ctrl+V`.
- Bấm `Img` nếu muốn chọn file ảnh từ máy.
- Bấm icon copy ở bubble để copy nội dung message.
- Bấm icon edit ở bubble user để đưa prompt cũ và ảnh cũ về composer rồi sửa/gửi lại.
- Bấm `Sessions` để mở lại các phiên chat đã lưu local.

## Ghi chú bảo mật

Các file sau không nên commit:

- `.provider-settings.json`: có thể chứa API key.
- `.certs/`: có private key HTTPS local.
- `.server-*.log`, `.server.pid`.
- `node_modules/`.

Các mục này đã được đưa vào `.gitignore`.

## Troubleshooting

Nếu mở Excel không thấy add-in:

```powershell
.\setup-excel-sideload.ps1
```

Sau đó restart Excel và vào lại `Home > Add-ins > Advanced > SHARED FOLDER`.

Nếu task pane mở nhưng không load:

```powershell
.\start-excel-chatbot.ps1
curl.exe -k https://localhost:3100/health
```

Nếu server đang chạy nhưng Excel vẫn giữ UI cũ, restart Excel để Office bỏ cache add-in cũ.

Nếu muốn kiểm tra manifest:

```powershell
npx office-addin-manifest validate manifest.xml
```
