# Chuẩn hóa thể thức văn bản (Word + Excel)

Add-in cho Microsoft Word và Excel giúp trình bày văn bản đúng thể thức:

- **Khối Đảng**: theo Hướng dẫn 05-HD/VPTW.
- **Chính quyền, Công an, Quân đội, MTTQ**: theo Nghị định 30/2020/NĐ-CP.
- **Excel**: căn chỉnh trang in A4 cho bảng tính.

Không cần mã kích hoạt. Add-in không gửi thông tin máy hay đơn vị đi đâu
(nút Cập nhật chỉ đọc trang phát hành công khai trên GitHub).

Lê Cường Việt - VPĐU xã Kon Plông

---

## Mục lục

- [Tải về](#tải-về)
- [Cài đặt](#cài-đặt-làm-1-lần)
- [Sử dụng trên Word](#sử-dụng-trên-word)
  - [1. Đơn vị ban hành](#1-đơn-vị-ban-hành-làm-1-lần)
  - [2. Chuẩn hóa văn bản có sẵn](#2-chuẩn-hóa-một-văn-bản-có-sẵn)
  - [3. Văn bản chép từ web, PDF, Zalo](#3-văn-bản-chép-từ-web-pdf-zalo)
  - [4. Số trang và phụ lục](#4-số-trang-và-phụ-lục)
  - [5. Tạo mục lục](#5-tạo-mục-lục)
  - [6. Tạo văn bản mẫu](#6-tạo-văn-bản-mẫu)
  - [7. Tùy chọn thể thức](#7-tùy-chọn-thể-thức)
- [Sử dụng trên Excel](#sử-dụng-trên-excel)
- [Cập nhật bản mới](#cập-nhật-bản-mới)
- [Gỡ cài đặt](#gỡ-cài-đặt)
- [Gặp lỗi?](#gặp-lỗi)
- [Các phiên bản](#các-phiên-bản)

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
5. Mở lại Word sẽ thấy tab **THỂ THỨC VĂN BẢN** (và nhóm nút ở tab Home), mở Excel sẽ thấy tab **CHUẨN HÓA IN**.

> Không nhấp đúp mở trực tiếp tệp `.dotm` / `.xlam`, hãy luôn cài bằng `CaiDat.bat`.
>
> Nếu Windows hiện cảnh báo "Windows protected your PC", bấm **More info** → **Run anyway**.

## Sử dụng trên Word

### 1. Đơn vị ban hành (làm 1 lần)

Tab **THỂ THỨC VĂN BẢN**, nhóm **Đơn vị ban hành**. Mỗi khối (Đảng, Chính quyền, Công an, Quân đội, MTTQ)
có tối đa **9 đơn vị**, mỗi đơn vị có thông tin và tùy chọn thể thức riêng.

- Khối Đảng có sẵn 4 đơn vị: **Đảng ủy xã, Văn phòng Đảng ủy, Ban Xây dựng Đảng, Ủy ban Kiểm tra**.
- Chọn đơn vị ở ô **Đơn vị** (có cả ở tab Home), rồi điền: *Tên gọi* (tên ngắn trong danh sách, không in ra),
  *Cấp trên, Cơ quan, Viết tắt, Địa danh, Thẩm quyền, Chức vụ, Họ tên* người ký.
- **Thêm**: chép đơn vị đang chọn thành đơn vị mới để sửa nhanh. **Xóa**: bỏ đơn vị đang chọn.

Cách điền:

| Ô | Ví dụ |
|---|---|
| Cấp trên (Đảng ủy xã) | `ĐẢNG BỘ TỈNH QUẢNG NGÃI` |
| Cấp trên (Văn phòng, các ban của Đảng ủy) | `ĐẢNG ỦY XÃ KON PLÔNG` |
| Cơ quan | `ĐẢNG ỦY XÃ KON PLÔNG`, `BAN XÂY DỰNG ĐẢNG`, `ỦY BAN NHÂN DÂN \| XÃ KON PLÔNG` |
| Viết tắt (trong số, ký hiệu) | `ĐU`, `VPĐU`, `UBND` |
| Thẩm quyền | `T/M BAN THƯỜNG VỤ`, `TM. ỦY BAN NHÂN DÂN`, `KT. CHỦ TỊCH`; để trống nếu ký trực tiếp |
| Chức vụ | `BÍ THƯ`, `TRƯỞNG BAN`, `PHÓ CHỦ TỊCH` |

- Tên dài muốn xuống dòng thì gõ dấu `|`.
- Cách nhanh: mở 1 văn bản đã đúng thể thức của đơn vị → bấm **Lấy từ văn bản đang mở**.

### 2. Chuẩn hóa một văn bản có sẵn

1. Mở văn bản, ô **Khối** để *Tự động nhận diện* (khuyên dùng): có "ĐẢNG CỘNG SẢN VIỆT NAM" là khối Đảng,
   có Quốc hiệu thì xét tên cơ quan để chọn Chính quyền / Công an / Quân đội / MTTQ.
2. Bấm **Chuẩn hóa thể thức**. Add-in sửa khổ A4, lề, phông chữ, Quốc hiệu – Tiêu ngữ / tiêu đề Đảng,
   số ký hiệu, ngày tháng, tên loại, trích yếu, Kính gửi, nội dung, chữ ký, Nơi nhận, số trang;
   gom chữ lẻ cuối đoạn; dọn văn bản chép về; dò từ nghi sai chính tả.
3. Đọc **báo cáo** hiện ra: phần *Đã chuẩn hóa* và các điểm *Cần kiểm tra trước khi trình ký*.
   Từ nghi sai chính tả (*nghành, sử lý, tập chung, của của*...) được gắn **ghi chú bên lề** kèm gợi ý sửa;
   add-in không tự sửa chữ. Sửa xong thì xóa ghi chú: thẻ **Review (Xem lại) → Delete (Xóa)**.
4. Không ưng ý: bấm **Ctrl+Z** một lần để hoàn tác toàn bộ.

Mẹo: phần đầu văn bản và phần chữ ký nên đặt trong bảng không viền (như văn bản mẫu của add-in).

**Tên cơ quan dài bị rớt dòng:** chọn ở ô **Tên dài**, rồi bấm Chuẩn hóa lại:
- *Hạ cỡ chữ* (khuyên dùng): nới cột tên cơ quan; khối Nhà nước hạ Quốc hiệu 12, Tiêu ngữ và ngày tháng 13, tên cơ quan 12.
- *Lấn lề trái*: giữ cỡ chữ, bảng tiêu đề lấn sang lề trái tối đa 15 mm.

**Ngày tháng** được đặt ngay dưới dòng "ĐẢNG CỘNG SẢN VIỆT NAM" (hoặc Quốc hiệu - Tiêu ngữ); ô bên phải
dòng số ký hiệu để trống.

**Kính gửi / Kính trình nhiều nơi:** gõ mỗi nơi nhận một dòng, bắt đầu bằng dấu `-`:

```
Kính gửi: - Chi bộ thôn 1;
          - Chi bộ thôn 2;
          - Chi bộ Trường Tiểu học và THCS.
```

Khi chuẩn hóa, các dòng được canh thẳng hàng với nơi nhận đầu tiên (gõ "Kính gửi:" riêng một dòng, hay dùng
phím Tab sau dấu ":" đều được); tên dài xuống dòng thì thẳng với chữ sau dấu `-`.
Nếu "Kính gửi" được đặt trong bảng tiêu đề, add-in tự đưa ra ngoài bảng rồi canh như trên.

**Các trường hợp khác khi chuẩn hóa:**
- Danh sách **đánh số tự động** (1., 2., gạch đầu dòng) trong nội dung được đổi thành gõ tay để thẳng lề như đoạn văn.
- Dòng ghi chú trong ngoặc đơn có chữ "kèm theo", ví dụ *(có danh sách trích ngang kèm theo)*: giữ canh giữa, in nghiêng.

### 3. Văn bản chép từ web, PDF, Zalo

Khi chuẩn hóa, phần nội dung được dọn tự động:

- Ngắt dòng bằng **Shift+Enter** giữa câu (làm chữ giãn rộng khi căn đều) → nối lại thành câu liền.
- Mỗi dòng bị tách thành một đoạn riêng → nối lại (khi dòng trên chưa kết câu và dòng dưới bắt đầu bằng chữ thường).
- **Dòng trống** chen giữa các đoạn → bỏ (khoảng cách đoạn đã do giãn đoạn đảm nhiệm).
- Khoảng trắng thừa, dấu cách cứng, tab, dấu cách trước dấu câu → bỏ.

Phần đầu, tên loại, trích yếu, chữ ký và mục lục được giữ nguyên.

### 4. Số trang và phụ lục

- Số trang bằng chữ số Ả-rập, **ở giữa đầu trang, cách mép trên 10 mm**, trang thứ nhất không hiện số.
- **Phụ lục đánh số riêng:** mỗi dòng tiêu đề bắt đầu bằng `PHỤ LỤC` nằm sau phần *Nơi nhận* được tách sang
  trang mới và đánh số lại từ 1.
- Số trang cũ ở chân trang (VD "Trang 1") được bỏ.

### 5. Tạo mục lục

Đặt con trỏ vào chỗ muốn chèn (một dòng trống) → bấm **Tạo mục lục**. Add-in tự nhận các đề mục
*Phần, Chương, Mục, I., Điều 1., 1. (in đậm), PHỤ LỤC* và lập mục lục có số trang, không đổi chữ hay định dạng
nội dung. Sửa văn bản xong, bấm lại nút để cập nhật đề mục và số trang.

### 6. Tạo văn bản mẫu

Bấm **Tạo văn bản mẫu** (hoặc *Mẫu khối Đảng / Nhà nước / MTTQ* khi để Tự động nhận diện) và chọn loại.
Thông tin lấy từ **đơn vị đang chọn**; chỗ cần điền được tô vàng trong dấu `[ ... ]`, số và ngày tháng để trống cho văn thư.

| Khối | Mẫu |
|---|---|
| Đảng (18) | Nghị quyết, Quyết định, Kế hoạch, Báo cáo, Thông báo, Tờ trình, Công văn, Giấy mời, Biên bản, Quy định, Quy chế, Chỉ thị, Kết luận, Hướng dẫn, Chương trình, Thông tri, Đề án, Giấy giới thiệu |
| Nhà nước (20) | Quyết định, Kế hoạch, Báo cáo, Thông báo, Tờ trình, Công văn, Giấy mời, Biên bản, Quyết định ban hành / phê duyệt văn bản, Quy định kèm theo quyết định, Quy chế kèm theo quyết định, Hướng dẫn, Chương trình, Quyết định ủy quyền, Giấy giới thiệu, Công điện, Phiếu chuyển, Nghị quyết (cá biệt), Chỉ thị, Giấy nghỉ phép |
| MTTQ (12) | Kế hoạch, Báo cáo, Thông báo, Công văn, Giấy mời, Tờ trình, Quyết định, Chương trình phối hợp, Hướng dẫn, Nghị quyết hội nghị, Biên bản, Giấy giới thiệu |

### 7. Tùy chọn thể thức

Nhóm **Tùy chọn: cỡ chữ** và **Tùy chọn: trang, đoạn** trên tab THỂ THỨC VĂN BẢN, lưu riêng cho **đơn vị đang chọn**.
Chỉ chọn được trong khoảng quy định cho phép; *Mặc định* là thông số chuẩn của khối.

| Tùy chọn | Khối Đảng | Khối Nhà nước |
|---|---|---|
| Cỡ chữ nội dung | 14 – 15 | 13 – 14 |
| Tên loại | 15 – 16 | 13 – 14 |
| Trích yếu | 14 – 15 | 13 – 14 |
| Quốc hiệu / tên cơ quan | cố định 15 / 14 | 12 – 13 |
| Lề trên, dưới / trái / phải | 20–25 / 30–35 / 15–20 mm | như khối Đảng |
| Lùi đầu dòng | 1 cm hoặc 1,27 cm | như khối Đảng |
| Giãn dòng | chính xác 18 – 22 pt | dòng đơn đến 1,5 dòng |
| Giãn đoạn | 6, 9, 12 pt | như khối Đảng |

- **Gom chữ lẻ cuối đoạn** (bật sẵn): đoạn có dòng cuối chỉ còn 1 chữ thì co nhẹ khoảng cách chữ (tối đa 0,5 pt) để kéo chữ đó lên.
- **Về mặc định**: bỏ hết tùy chọn của đơn vị đang chọn.
- **Hiện ở tab Home**: bỏ tích nếu muốn tab Home gọn hơn.

## Sử dụng trên Excel

Mở bảng tính cần in → tab **CHUẨN HÓA IN** → bấm **Chuẩn hóa trang in**. Add-in tự:

- Bật xuống dòng, giãn hàng để không che chữ (kể cả ô gộp); nới cột số bị hiện `####`.
- Chọn A4 dọc hoặc ngang theo độ rộng bảng, vừa 1 trang theo chiều ngang.
- Lặp dòng tiêu đề bảng ở các trang sau, kéo phần người ký về chung trang với cuối bảng.
- Đánh số "Trang x/y" khi từ 2 trang trở lên; cảnh báo nếu chữ in ra quá nhỏ.

Muốn trả lại như cũ: bấm **Hoàn tác** (trước khi đóng Excel). Sheet đang khóa thì bỏ khóa trước.

## Cập nhật bản mới

- Bấm nút **Cập nhật** để kiểm tra ngay. Mỗi tuần add-in cũng tự kiểm tra một lần khi bấm *Chuẩn hóa thể thức*
  (không có mạng thì bỏ qua, không báo lỗi).
- Có bản mới: bấm **Yes** để tải, rồi **đóng hết Word và Excel**. Bản mới tự cài và báo khi xong; mở lại là dùng được.
- Tệp tải về được kiểm tra mã SHA-256 do GitHub công bố; sai mã thì không cài.
- Thông tin đơn vị và tùy chọn thể thức được giữ nguyên.

> **Đang dùng bản 1.0.0:** bản này chưa tự cài được. Tải bộ cài mới ở mục [Tải về](#tải-về),
> đóng Word/Excel rồi chạy lại `CaiDat.bat` một lần; từ bản 1.1.0 trở đi sẽ tự cập nhật.

## Gỡ cài đặt

Đóng Word, Excel rồi nhấp đúp `GoCaiDat.bat` trong thư mục bộ cài.

## Gặp lỗi?

- **Không thấy tab trên Word/Excel:** đóng hẳn Word/Excel rồi chạy lại `CaiDat.bat`; nhớ bước *Unblock* trước khi giải nén.
- **Báo macro bị chặn:** xóa bộ cài cũ, tải lại, làm đủ bước *Unblock* rồi cài lại.
- **Không thấy nhóm nút ở tab Home:** vào tab THỂ THỨC VĂN BẢN, tích lại ô *Hiện ở tab Home*.
- **Cập nhật tự động không chạy:** tải bộ cài mới ở mục [Tải về](#tải-về) và cài lại bằng `CaiDat.bat`.
- Hướng dẫn chi tiết có trong tệp `HuongDan.txt` đi kèm bộ cài.

## Các phiên bản

**1.1.3**
- Chuẩn hóa dò thêm từ nghi sai chính tả (sai cấu tạo âm tiết, cụm từ hay viết sai, gõ lặp từ), ghi chú bên lề kèm gợi ý sửa.
- Văn bản được đặt ngôn ngữ tiếng Việt, Word không gạch đỏ theo từ điển tiếng Anh.

**1.1.2**
- Kính gửi đặt trong bảng tiêu đề được đưa ra ngoài và canh thẳng hàng.
- Ngày tháng chuyển lên dưới Tiêu ngữ không còn dính đường kẻ; giữ khoảng trống điền ngày.
- Đánh số tự động đổi thành gõ tay cho thẳng lề; dòng "(... kèm theo)" giữ canh giữa.

**1.1.1**
- Ngày tháng nằm ngay dưới tiêu đề (cả khối Đảng, Nhà nước, MTTQ).
- Kính gửi / Kính trình nhiều nơi canh thẳng hàng; mẫu Công văn, Tờ trình có sẵn dạng gửi nhiều nơi.

**1.1.0**
- Nhiều đơn vị ban hành cho mỗi khối (tối đa 9); khối Đảng có sẵn Đảng ủy xã, Văn phòng, Ban Xây dựng Đảng, Ủy ban Kiểm tra.
- Tùy chọn thể thức theo từng đơn vị, chỉ trong khoảng quy định.
- 50 mẫu văn bản: Đảng 18, Nhà nước 20, MTTQ 12.
- Số trang theo hướng dẫn: giữa đầu trang, cách mép trên 10 mm, phụ lục đánh số riêng.
- Nút **Tạo mục lục**.
- Dọn văn bản chép từ web, PDF, Zalo (ngắt dòng, dòng trống, khoảng trắng); gom chữ lẻ cuối đoạn.
- Tự kiểm tra và tự cài bản mới; tùy chọn hiện nhóm nút ở tab Home.

**1.0.0** — Bản đầu tiên.
