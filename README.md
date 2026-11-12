# Chuẩn hóa thể thức văn bản (Word + Excel)

Add-in cho Microsoft Word và Excel giúp trình bày văn bản đúng thể thức:

- **Khối Đảng**: theo Hướng dẫn 05-HD/VPTW.
- **Chính quyền, Công an, Quân đội, MTTQ**: theo Nghị định 30/2020/NĐ-CP.
- **Excel**: căn chỉnh trang in A4 cho bảng tính.

Không cần mã kích hoạt. Add-in không gửi thông tin máy hay đơn vị đi đâu.

**Tác giả:** Lê Cường Việt - VPĐU xã Kon Plông

---

## Tải về

👉 **[Tải bộ cài ChuanHoaVB_KonPlong.zip (bản mới nhất)](https://github.com/Bi2702/ChuanHoaVB/releases/latest/download/ChuanHoaVB_KonPlong.zip)**

Hoặc vào trang [Releases](https://github.com/Bi2702/ChuanHoaVB/releases) và tải tệp `ChuanHoaVB_KonPlong.zip`
ở mục **Assets** (không tải mục *Source code*).

**Yêu cầu:** Windows, Microsoft Word và Excel bản cài trên máy (Office 2010 trở lên / Microsoft 365).

## Cài đặt (làm 1 lần)

1. **Bỏ chặn tệp:** chuột phải `ChuanHoaVB_KonPlong.zip` → **Properties** → tích **Unblock** (Bỏ chặn) → **OK**.
2. **Giải nén:** chuột phải tệp .zip → **Extract All** (Giải nén tất cả).
3. **Lưu và đóng hết** Word, Excel.
4. Mở thư mục vừa giải nén, **nhấp đúp `CaiDat.bat`**. Thấy dòng *"Cai dat xong"* là được.
5. Mở lại Word sẽ thấy tab **THỂ THỨC VĂN BẢN**, mở Excel sẽ thấy tab **CHUẨN HÓA IN**.

> Không nhấp đúp mở trực tiếp tệp `.dotm` / `.xlam`, hãy luôn cài bằng `CaiDat.bat`.
>
> Nếu Windows hiện cảnh báo "Windows protected your PC", bấm **More info** → **Run anyway**.

## Sử dụng trên Word

**Bước đầu tiên: nhập thông tin đơn vị** (tab THỂ THỨC VĂN BẢN), chỉ cần làm 1 lần:
Cấp trên, Cơ quan, Viết tắt, Địa danh, Thẩm quyền, Chức vụ, Họ tên người ký.
Thông tin được lưu riêng cho từng khối.

- Tên dài muốn xuống dòng thì gõ dấu `|` (VD: `ỦY BAN NHÂN DÂN | XÃ KON PLÔNG`).
- Khối Đảng: ô *Cấp trên* ghi đảng bộ cấp trên trực tiếp (VD: `ĐẢNG BỘ TỈNH QUẢNG NGÃI`).
- Người ký ký trực tiếp thì để trống *Thẩm quyền*. Ký thay: *Thẩm quyền* `KT. CHỦ TỊCH`, *Chức vụ* `PHÓ CHỦ TỊCH`.
- Cách nhanh: mở 1 văn bản đã đúng thể thức của đơn vị → bấm **Lấy từ văn bản đang mở**.

**Chuẩn hóa một văn bản có sẵn**

1. Mở văn bản, ô **Khối** để *Tự động nhận diện* (khuyên dùng).
2. Bấm **Chuẩn hóa thể thức**. Add-in sửa khổ A4, lề, phông chữ, Quốc hiệu – Tiêu ngữ / tiêu đề Đảng,
   số ký hiệu, ngày tháng, tên loại, trích yếu, Kính gửi, nội dung, chữ ký, Nơi nhận, số trang.
3. Đọc **báo cáo** hiện ra sau khi chuẩn hóa: các điểm cần tự kiểm tra trước khi trình ký.
4. Không ưng ý: bấm **Ctrl+Z** một lần để hoàn tác toàn bộ.

Mẹo: phần đầu văn bản và phần chữ ký nên đặt trong bảng không viền (như văn bản mẫu của add-in).

**Tên cơ quan dài bị rớt dòng:** chọn ở ô **Tên dài**, rồi bấm Chuẩn hóa lại:
- *Hạ cỡ chữ* (khuyên dùng): nới cột tên cơ quan, khối Nhà nước hạ cỡ chữ phần đầu.
- *Lấn lề trái*: giữ cỡ chữ, bảng tiêu đề lấn sang lề trái tối đa 15 mm.

**Tạo văn bản mẫu:** bấm **Tạo văn bản mẫu** và chọn loại (Nghị quyết, Quyết định, Kế hoạch, Báo cáo,
Thông báo, Tờ trình, Công văn, Giấy mời, Biên bản). Chỗ cần điền được tô vàng trong dấu `[ ... ]`.

## Sử dụng trên Excel

Mở bảng tính cần in → tab **CHUẨN HÓA IN** → bấm **Chuẩn hóa trang in**. Add-in tự:

- Bật xuống dòng, giãn hàng để không che chữ; nới cột số bị hiện `####`.
- Chọn A4 dọc hoặc ngang theo độ rộng bảng, vừa 1 trang theo chiều ngang.
- Lặp dòng tiêu đề bảng ở các trang sau, kéo phần người ký về chung trang với cuối bảng.
- Đánh số "Trang x/y" khi từ 2 trang trở lên.

Muốn trả lại như cũ: bấm **Hoàn tác** (trước khi đóng Excel). Sheet đang khóa thì bỏ khóa trước.

## Cập nhật bản mới

Bấm nút **Cập nhật** trên tab của add-in để kiểm tra. Có bản mới thì tải về,
**đóng Word/Excel** rồi chạy lại `CaiDat.bat`. Thông tin đơn vị đã nhập được giữ nguyên.

## Gỡ cài đặt

Đóng Word, Excel rồi nhấp đúp `GoCaiDat.bat` trong thư mục bộ cài.

## Gặp lỗi?

- **Không thấy tab trên Word/Excel:** đóng hẳn Word/Excel rồi chạy lại `CaiDat.bat`; nhớ bước *Unblock* trước khi giải nén.
- **Báo macro bị chặn:** xóa bộ cài cũ, tải lại, làm đủ bước *Unblock* rồi cài lại.
- Hướng dẫn chi tiết có trong tệp `HuongDan.txt` đi kèm bộ cài.
