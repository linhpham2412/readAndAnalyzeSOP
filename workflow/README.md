# Workflow kiểm tra nhập hàng nhượng quyền theo tuần

## Chạy hàng tuần
```powershell
pip install pandas openpyxl xlrd        # lần đầu
python workflow/kiem_tra_nhap_hang.py "19 - 25.10.xls"            # báo cáo tổng
python workflow/kiem_tra_nhap_hang.py "19 - 25.10.xls" --per-nq   # thêm 1 file riêng cho từng NQ
python workflow/kiem_tra_nhap_hang.py --validate                  # đối chiếu logic với CHECK_2026.xlsx
```
Kết quả: `output/BaoCao_NhapHang_<tên file>.xlsx` (và thư mục `..._NQ/` nếu dùng `--per-nq`).

## Các bước thay cho thao tác thủ công
| Thủ công | Workflow |
|---|---|
| Pivot Table (Mã KH x Mã HH, tổng SL) | `load_weekly_report` (bỏ dòng Total) |
| Sheet SOP_Banh / SOP_Nuoc / QuyDoi_DonVi / ChiTiet | `compute_nq` (cùng công thức) |
| Sheet KIỂM TRA + tô đỏ theo ngưỡng | `build_findings`, sheet "Ma trận %" |
| Tổng hợp tình hình nhập NVL | sheet "Tổng quan" + "Chi tiết từng NQ" |

Danh mục NQ (Quản lý, Mã NQ, Tên) lấy từ sheet `KIỂM TRA` của `CHECK_2026.xlsx`. Khi có NQ mới, thêm vào sheet đó
(sheet "KH ngoài danh mục" của báo cáo sẽ liệt kê các KH chưa có trong danh mục).

## Quy tắc
- `% = thực nhập * 100 / định mức SOP - 100` (âm = thiếu, dương = dư).
- Ngưỡng thiếu theo từng NVL lấy từ Conditional Formatting cũ: thịt/hộp/túi hotdog/xúc xích -10%; tương cà, mayo, dầu, pate, jambon,
  chà bông, túi dựng ly, trà đen -20%; ly, màng ép, ống hút -30%; sữa bột, đường, cacao, pepsi = 0 là "không nhập".
- Thêm: tổng ly quy đổi từ NVL nước so với ly SOP (sheet SOP_Nuoc) cảnh báo khi < -30% (sửa `TH_NUOC_LY`).
- Tình trạng NQ: `ĐỦ HÀNG`, `THIẾU HÀNG`, `KHÔNG CÓ ĐƠN NHẬP`, `KHÔNG NHẬP BÁNH`.
- Muốn đổi ngưỡng/hệ số: sửa `BANH_ITEMS` và các hằng số đầu file `kiem_tra_nhap_hang.py`.

## Lỗi chữ tiếng Việt trong báo cáo .xls
File `19 - 25.10.xls` **không phải TCVN3**: nó lưu theo Windows-1252 nên chữ Việt ngoài bảng mã này (ă, đ, ơ, ư, ạ, ế, ệ, ố...)
đã bị thay bằng dấu `?` ngay trong file (không thể chuyển ngược bằng bảng mã). Chữ như à á â ê ì í ò ó ô ù ú ý vẫn đúng.
- Mã hàng, mã KH, số lượng không bị ảnh hưởng nên phân tích vẫn chính xác.
- Tên KH được khôi phục tự động bằng `viet_restore.py` + từ điển `tu_dien_phuc_hoi.txt` (thêm từ/cụm từ đúng vào file này khi còn sót `?`).
- Sheet "KH ngoài danh mục" hiển thị cả tên gốc và tên đã khôi phục để đối chiếu. Tên của NQ trong danh mục lấy từ CHECK_2026.xlsx.
- Nếu một ngày file thật sự là TCVN3, script sẽ in cảnh báo.

## Sử dụng file thực thi `KiemTraNhapHang.exe`
1. Mở trực tiếp `KiemTraNhapHang.exe` ở thư mục gốc.
2. **Kéo thả** file báo cáo tuần (`.xls` hoặc `.xlsx`) vào khung giao diện, hoặc bấm nút **Chọn file...** để mở cửa sổ chọn file.
3. (Tùy chọn) Tích chọn *Xuất thêm file riêng cho từng Nhượng Quyền (`--per-nq`)*.
4. Bấm nút **XỬ LÝ VÀ MỞ BÁO CÁO** -> ứng dụng sẽ gọi `.\workflow\kiem_tra_nhap_hang.py "<file đã chọn>"`, xuất báo cáo vào thư mục `output/`, **tự động mở file Excel kết quả** và **tự động đóng chương trình**.
