# Tự động hóa báo cáo thị trường tuần

Chương trình giữ nguyên mẫu PowerPoint hiện tại, tự lấy dữ liệu từ các file Excel đã xử lý, cập nhật bảng/biểu đồ/ngày báo cáo, tạo gói dữ liệu cho phần nhận định và xuất PPTX/PDF sau khi người dùng duyệt.

Mẫu mặc định là file gốc `MBS Dau Tu - BC Thi truong Tuan - 03.08.2026.pptx`. Chương trình giữ nguyên thiết kế và cấu trúc 11 slide của file này, chỉ tự động cập nhật dữ liệu, hình ảnh và nội dung vào đúng vị trí.

## Cài đặt lần đầu trên Windows 11

Yêu cầu:

- Microsoft Excel và Microsoft PowerPoint bản desktop đã được cài đặt và kích hoạt. Bản Office trên trình duyệt không đủ để chạy chương trình.
- Python 3.12 hoặc mới hơn. Khi cài Python từ python.org, chọn `Add python.exe to PATH`.
- Kết nối Internet trong lần cài đầu để tải các thư viện Python.

Nhấp đúp `0_Cai_dat_Windows.bat`. Chương trình sẽ tạo môi trường `.venv`, cài đúng thư viện cần thiết và kiểm tra kết nối tự động với Excel/PowerPoint. Chỉ cần chạy lại file này khi đổi máy hoặc cần sửa môi trường Python.

## Quy trình chiều thứ Sáu trên Windows 11

1. Tải các file đầu vào từ FiinPro và đặt tất cả vào thư mục `incoming`.
2. Nếu có ảnh biểu đồ VNINDEX, lưu một file PNG với tên bất kỳ vào `incoming`. Nếu có nhiều ảnh, chương trình dùng file PNG mới nhất; tên `vnindex_chart.png` vẫn được ưu tiên.
3. Nếu có thông tin vĩ mô hoặc bối cảnh thị trường muốn GPT sử dụng, ghi vào `inputs/market_context.txt`.
4. Nhấp đúp `1_Tao_ban_nhap_Windows.bat`. Chương trình sẽ:
   - nhận diện và sao chép file vào đúng thư mục nghiệp vụ;
   - sao lưu bộ dữ liệu hiện tại vào `archive`;
   - chạy `process_data.py all`;
   - tạo PowerPoint nháp và bộ dữ liệu/prompt nhận định trong `output_reports/<ngày>/review`.
5. Kiểm tra bản nháp. Nếu đã cấu hình OpenAI, sửa trực tiếp `review/commentary.json`. Nếu chưa, đưa nội dung `review/review_prompt.txt` vào ChatGPT rồi lưu JSON trả về thành `review/commentary.json`.
6. Nhấp đúp `2_Xuat_bao_cao_Windows.bat` để chèn nhận định đã duyệt và xuất PowerPoint/PDF cuối cùng.

Chương trình không xóa file trong `incoming`; sau khi hoàn tất tuần, có thể tự chuyển chúng sang nơi lưu trữ riêng.

## Cấu hình OpenAI (tùy chọn)

Thêm dòng sau vào file `.env` trong thư mục dự án, không gửi khóa qua chat:

```text
OPENAI_API_KEY=khóa_của_bạn
```

Có thể đổi model bằng `OPENAI_MODEL=...`. Nếu không có khóa, quy trình vẫn tạo prompt hoàn chỉnh để dùng thủ công trong ChatGPT.

## Quy tắc đầu ra cố định

- Báo cáo ngày 24/08/2026 là chuẩn văn phong chính thức. File `inputs/commentary_reference_2026-08-24.json` lưu nguyên văn phần nhận định đã duyệt để mọi kỳ sau giữ cùng cấu trúc, nhịp câu, thuật ngữ và mức độ chi tiết.
- Nhận định Slide 3 luôn gồm ba đoạn theo cùng một cấu trúc: diễn biến chỉ số và thanh khoản; bối cảnh vĩ mô/quốc tế và triển vọng; đóng góp của nhóm Vingroup. Prompt quy định câu mở đầu, thứ tự số liệu, độ dài, thuật ngữ và cách làm tròn để giữ văn phong nhất quán giữa các kỳ.
- Slide 8 tự sao chép nguyên văn hai đoạn đầu của Slide 3. Đoạn thứ ba về đóng góp của Vingroup không được chèn vào Slide 8.
- Mọi chữ mà chương trình tạo hoặc cập nhật trong PowerPoint, bảng Excel trung gian và biểu đồ Excel đều dùng Arial. Trước khi lưu/xuất PDF, toàn bộ chữ nhìn thấy trong các slide được chuẩn hóa lại về Arial.
- Đóng góp Vingroup chỉ gồm VIC, VHM, VRE và VPL. Chương trình chỉ chấp nhận file có cặp trường `Giá đóng cửa` và `Số CP lưu hành`, tự sửa các ngày trùng trong lịch sử và chặn xuất bản nếu số liệu bất hợp lý.
- Bảng `Danh mục MBS` được đánh lại STT từ 1 đến 70 sau khi sắp xếp.
- `Lịch sự kiện` dịch các sự kiện vĩ mô đã lọc sang tiếng Việt và gộp mọi sự kiện cùng ngày vào một dòng ngày duy nhất.
- `NN bán ròng` chỉ bổ sung dữ liệu mới; nội dung chữ và khối nhãn hiện có trong workbook được giữ nguyên.

## Các lệnh hữu ích

Trên Windows PowerShell:

```text
.\.venv\Scripts\python.exe weekly_report.py doctor
.\.venv\Scripts\python.exe weekly_report.py route
.\.venv\Scripts\python.exe weekly_report.py process --as-of YYYY-MM-DD
.\.venv\Scripts\python.exe weekly_report.py prepare --as-of YYYY-MM-DD
.\.venv\Scripts\python.exe weekly_report.py all --as-of YYYY-MM-DD
.\.venv\Scripts\python.exe weekly_report.py finalize
```

Trên macOS:

```text
python3 weekly_report.py route
python3 weekly_report.py process --as-of YYYY-MM-DD
python3 weekly_report.py prepare --as-of YYYY-MM-DD
python3 weekly_report.py all --as-of YYYY-MM-DD
python3 weekly_report.py finalize
```

`--as-of` là ngày giao dịch cuối tuần, thường là thứ Sáu. Dùng `--force` với `prepare` hoặc `all` khi cần tạo lại bản nháp của cùng một tuần.

Lệnh `finalize` luôn yêu cầu `review/commentary.json` đã được duyệt. Tùy chọn cũ `--keep-existing-commentary` bị từ chối vì có thể vô tình xuất phần nhận định của kỳ mẫu.

## Quy tắc nhận diện file đầu vào

- `OneDay`, `OneWeek`, `OneMonth` → dữ liệu ngành.
- `Top giá trị ròng...` của tổ chức trong nước → tự doanh; hai file Mua/Bán được ghép thành workbook dùng cho PowerPoint.
- Các file `Top giá trị ròng...` khác và `Room_*.pdf` → giao dịch nước ngoài.
- Một file `Chi_so_&_Nganh...` → được sao chép đồng thời sang thanh khoản và đóng góp Vingroup; không cần nhân đôi thủ công.
- `De_Doanh_Nghiep...` → danh mục MBS.
- `Du_lieu_Giao_dich_Doanh_nghiep...xlsx` → tự phân loại theo mã cổ phiếu trong file.
- `Giao Dich To Chuc Trong Nuoc_Top gia tri rong_Mua/Ban...xlsx` → bảng Top 10 tự doanh; mô-đun chỉ sao chép dữ liệu từ hai file này và không còn đọc PDF hay tổng hợp lịch sử.
- File không nhận diện được sẽ được báo rõ và không bị di chuyển.

## Lưu ý

- Bản nháp luôn dừng ở bước duyệt nhận định; chương trình không tự xuất báo cáo cuối khi chưa có sự kiểm tra của người dùng.
- Chương trình chọn nguồn theo ngày trong tên file và ngày `--as-of`, không chỉ theo thời điểm file được sửa. Nếu nguồn bắt buộc bị thiếu, quá cũ, nằm sau ngày báo cáo hoặc cặp Mua/Bán không cùng kỳ, lỗi được ghi vào `run_state.json` và bước xuất bản bị chặn.
- Nếu Excel đang bận, chương trình chuyển sang bộ đọc dự phòng. Riêng bảng danh mục, các chỉ tiêu giá/hiệu suất/định giá được tính trực tiếp từ sheet `Cập nhật giá`, nên không bị trống vì công thức chưa có giá trị lưu sẵn.
- Trên Windows, PowerPoint desktop trực tiếp xuất PDF; không cần LibreOffice, Node.js hoặc bộ công cụ Codex.
- Trên macOS, nếu LibreOffice không xuất được PDF, chương trình dùng bộ render dự phòng khi môi trường Codex có sẵn.
- Trong lúc chạy, không nên chỉnh đúng file Excel hoặc PowerPoint đang được chương trình mở. Chương trình cố gắng giữ nguyên các file Office khác đang mở, nhưng đóng file đích trước khi chạy vẫn là cách an toàn nhất.
