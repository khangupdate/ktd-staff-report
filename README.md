# KTD Staff Report — báo cáo nhân sự tạm thời

Web tĩnh **GitHub Pages** để lập báo cáo nhân sự Khiết Tâm Đường từ hai file Excel kỳ **15/09–14/10/2026**.

**Mở web:** https://khangupdate.github.io/ktd-staff-report/

## Cách dùng

1. Truy cập link trên.
2. Nhấn **Nhập 2 file Excel**.
3. Chọn hai tệp CHẤM CÔNG KTD.xlsx và KIỂM KHO.xlsx.
4. Nhấn **Tạo báo cáo**; chọn Quỳnh, Kiên, Tính, Thương hoặc Tổng quan.
5. Dùng **In / Lưu PDF** nếu cần.

Ứng dụng không có backend, database, Google Sheets API, BMT hoặc ERP. Toàn bộ dữ liệu Excel được đọc ngay trong trình duyệt và dữ liệu tổng hợp được lưu trong localStorage trên thiết bị đang xem.

## Giới hạn quan trọng

- **Tuyệt đối không commit Excel, thông tin khách hàng, lương hay dữ liệu nhân viên lên repo public.**
- Repo public chỉ chứa mã nguồn. Một người khác truy cập web không nhìn thấy dữ liệu đã nhập trên trình duyệt của bạn.
- Ngày, giờ làm trong ca (09:00–18:00) và giờ sau 18h lấy từ file chấm công. Phần sau 18h chưa tính lương.
- Đơn giá theo file: mức ở C3 của tab chấm công (18.700đ/giờ) và riêng Tính 27.300đ/giờ. Hiển thị **ước tính tiền theo giờ, không phải bảng lương cuối**.
- Ca dựa trên tab 13 - CĂN CỨ KTV của kiểm kho và lọc đúng kỳ. Các ca *tạm tính* khác với ca *đã duyệt*. Không lấy một dòng xuất Albumin làm chứng minh cho toàn bộ hoạt động trị liệu.
- Tab kiểm kho này ghi nhận tới 10/10/2026 trong snapshot; dữ liệu kỳ 11–14/10 có thể chưa đủ. Danh mục nhân sự trong kho và trong chấm công chưa hoàn toàn nhất quán.
- Font và thư viện JS đọc Excel tải qua CDN lúc mở trang, nhưng **không tự tải file Excel lên bên ngoài**.
- Trên thiết bị dùng chung, nhấn **Xóa dữ liệu đã lưu trên thiết bị** sau khi xem.

## Chuyển repo sang Private khi không dùng nữa

GitHub → repo → Settings → General → Danger Zone → Change repository visibility → Private.

**Lưu ý:** GitHub Pages có thể ngừng hoạt động sau khi repo chuyển private, tùy gói GitHub. Nội dung từng public có thể bị sao chép/cache, kể cả sau khi chuyển private.
