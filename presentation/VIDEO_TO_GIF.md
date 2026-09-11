# Chuyển video thành GIF cho PowerPoint

## Cách dễ nhất

1. Nhấp đúp **Video_to_GIF.cmd** trong thư mục `presentation`.
2. Chọn video. Giữ **Ctrl** hoặc **Shift** để chọn nhiều video cùng lúc.
3. Chờ thông báo **Done**. GIF được lưu tự động vào `presentation/assets`.

Bạn cũng có thể **kéo thả một hoặc nhiều video lên Video_to_GIF.cmd**.
Không cần sửa script hay di chuyển video gốc. Khi chuyển lại cùng một video,
tên GIF mới có thêm `_2`, `_3`… để giữ các bản trước đó.

Mặc định: **nhanh 5 lần**, rộng tối đa **800 px**, khoảng **8 khung hình/giây**,
**256 màu cho từng khung hình**, có dithering để giảm các mảng màu bị phân bậc,
giữ tỷ lệ khung hình, chuyển toàn bộ video và **lặp vô hạn**.
GIF không có âm thanh. Tên đầu ra là `<tên video>_5x.gif`.

Mặc định màu đã được cải thiện sau bản xuất đầu tiên: bảng màu được học riêng
cho từng khung hình ở đúng kích thước đầu ra để giữ màu đường biểu đồ tốt hơn.
Dithering làm chuyển màu mượt hơn nhưng có thể thấy hạt nhỏ khi phóng to.
Các GIF đã xuất trước đó không tự thay đổi; cần chuyển lại để áp dụng.

## Hai GIF đã xuất từ video platoon

Cả hai nằm trong `presentation/assets`, đều nhanh **5×**, dài **16,16 giây**
so với **80,8 giây** của video gốc và lặp vô hạn:

| Tệp | Kích thước | Dung lượng | So với video gốc 14,82 MB |
| --- | --- | --- | --- |
| `Distributed Observer Controller Co-Design for String Stability in Vehicle Platoons (1)_5x.gif` | 800 × 450 | 5,81 MB | Nhẹ hơn 60,8% |
| `Distributed Observer Controller Co-Design for String Stability in Vehicle Platoons (1)_5x_960px.gif` | 960 × 540 | 10,33 MB | Nhẹ hơn 30,3% |

Chọn bản đầu để tiết kiệm dung lượng, hoặc bản 960 px nếu cần hình và chữ rõ hơn.

## Đổi thông số khi cần

Mở PowerShell tại thư mục gốc dự án rồi chạy:

```powershell
# Đổi video: chỉ thay đường dẫn trong dấu nháy.
py -3 .\presentation\tools\video_to_gif.py "C:\duong-dan\video-khac.mp4"

# Bản rõ hơn: tăng kích thước và số khung hình.
py -3 .\presentation\tools\video_to_gif.py "C:\duong-dan\video-khac.mp4" --width 960 --fps 10

# Bản nhẹ với cấu hình cũ (màu kém hơn).
py -3 .\presentation\tools\video_to_gif.py "C:\duong-dan\video-khac.mp4" --colors 96 --palette global --dither none

# Video Distributed Trust Estimation: nhanh 2 lần, màu cải thiện.
py -3 .\presentation\tools\video_to_gif.py "C:\Users\Quang Huy Nugyen\Videos\Distributed Trust Estimation.mp4" --speed 2

# Đổi tốc độ: ví dụ nhanh 3 lần.
py -3 .\presentation\tools\video_to_gif.py "C:\duong-dan\video-khac.mp4" --speed 3
```

Có thể đặt thông số mà không truyền đường dẫn video để mở hộp thoại chọn tệp:

```powershell
py -3 .\presentation\tools\video_to_gif.py --width 960 --fps 10
```

GIF không phải lúc nào cũng nhẹ hơn MP4. Công cụ in dung lượng thực tế sau mỗi
lần xuất; giảm `--width`, `--fps` hoặc `--colors` nếu cần. Giảm kích thước cũng
làm chữ và biểu đồ nhỏ khó đọc hơn.

## Môi trường

Công cụ dùng Python 3, OpenCV và Pillow, đã có trên máy này. Nếu chuyển sang
máy khác có Python, cài các thư viện bằng:

```powershell
py -3 -m pip install opencv-python Pillow
```

Thời lượng nguồn được tính theo số khung hình / tốc độ khung hình mà OpenCV
đọc được; phù hợp với video tốc độ khung hình cố định như video hiện tại.
Với video có tốc độ khung hình biến thiên, thời gian có thể sai lệch.
Thời lượng GIF được làm tròn đến 0,01 giây theo định dạng GIF.
Các khung hình xuất được giữ trong RAM, nên video rất dài hoặc độ phân giải
lớn sẽ cần nhiều bộ nhớ; hãy giảm kích thước và số khung hình khi cần.
