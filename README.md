# ggtrend_pj

## CÔNG VIỆC GIAI ĐOẠN 1

Bạn hãy thực hiện các bài tập dưới đây, kết quả gửi về đội ngũ Mentor theo lưu ý bên dưới:<br>
•	File kết quả là 01 file nén với tên: DAK27_Hoten_GĐ1.rar.<br>
File kết quả bao gồm:<br>
  o	File vn_trend_2020.xls. 
  o	Source Code của các bạn để sinh ra file vn_trend_2020.xls

Bạn hãy thực hiện xử lý các trường hợp sau
1.	Viết chương trình để tự động lấy thông tin xu hướng của các từ khóa trong file key_trends.xls, xu hướng lấy trong các phạm vi sau:
a.	Lấy tại lãnh thổ Việt Nam
b.	Lấy trong thời gian 1/1/2020 – 31/12/2020
2.	Lưu kết quả thu thập được vào 1 file tên là vn_trend_2020.xls.<br>
a.	Lưu kết quả thu được vào nhiều sheet, tên mỗi sheet chính là tên tiêu đều, ví dụ như sau:

 ![image](https://user-images.githubusercontent.com/75520765/135628216-2ac86f5c-f4c8-43bc-8b30-71ce2ced5d5f.png)
 
b.	Nội dung trong 1 sheet sẽ là thông tin xu hướng lấy được từ google trends của các từ khóa tương ứng . Ví dụ như sau (bạn không cần phần màu trong ảnh, chỉ cần đúng nội dung là được)

## CÔNG VIỆC GIAI ĐOẠN 2

Bạn hãy thực hiện các bài tập dưới đây, kết quả gửi về đội ngũ Mentor theo lưu ý bên dưới:<br>
•	File kết quả là 01 file nén với tên: DAK27_Hoten_GĐ2.rar. File kết quả bao gồm:<br>
    o	File vn_trending_result.xls. 
    o	Quy trình cài đặt PostgreSQL bạn đã thực hiện (chụp ảnh các bước cài đặt) và các câu lệnh tạo user, tạo bảng vn_trending. Lưu file này là quytrinh_db.docx.
    o	Source code của bạn 

Bạn hãy thực hiện xử lý các trường hợp sau<br>
1.	Cài đặt Cơ sở dữ liệu PostgreSQL phiên bản 12 với mục đích lưu trữ các dữ liệu Trending thu thập được.
2.	Thực hiện tạo bảng tên vn_trending theo cấu trúc các trường như sau:
STT	Tên cột	Kiểu dữ liệu và mục đích<br>
1	Id	Sinh tự động giá trị cho trường này. Đây là trường lưu giá trị dạng số
2	keyword	kiểu ký tự, lưu keyword tìm kiếm
3	date	kiểu date, lưu ngày tìm kiếm của keyword
4	Value	Kiểu số, lưu số lần tìm kiếm trong ngày của keyword
5	trend_type	lưu trữ các nhóm tin tức: news, film, songs,... trong file key_trends.xls
3.	Hãy phát triển tiếp chương trình bạn đã xây dựng ở giai đoạn 1. 
Chương trình của bạn lúc này sẽ lấy trực tiếp từ Google Trend và đưa vào bảng trong Database vừa tạo bên trên.
4.	Bạn thực hiện phân tích dữ liệu sau và lưu ra file excel tên là vn_trending_result.xls.
-	Hãy lấy TOP 10 từ khóa được tìm kiếm nhiều nhất trong danh sách trên.

## Dự án giai đoạn 3 

Bạn hãy thực hiện các bài tập dưới đây, kết quả gửi về đội ngũ Mentor theo lưu ý bên dưới:<br>
•	File kết quả là 01 file nén với tên: DAK27_Hoten_GĐ3.rar. File kết quả bao gồm:
    o	File vn_trending_search_keyword_2020.xlsx, top_search_key_2020.png, top_search_key_2019.png
    o	Source code của bạn với định dạng file .py  (lưu ý phải đúng định dạng này)

1. Báo cáo các từ khóa search 2020 tại Việt Nam, nhóm theo loại từ khóa tìm kiếm (News, film, song, person, ...), xuất file excel vn_trending_search_keyword_2020.xlsx
	Gồm các cột:<br>
  
![image](https://user-images.githubusercontent.com/75520765/135628766-f32e7f59-1419-4ba9-895a-4916ce8ea288.png)

2. Vẽ biểu đồ line chart top 5 trending các từ khóa tìm kiếm nhiều nhất 2020, xuất file ảnh top_search_key_2020.png

![image](https://user-images.githubusercontent.com/75520765/135628808-945e059a-6aa0-4bc8-b013-100bba66366d.png)
 
3. Vẽ biểu đồ bar chart top 5 trending các từ khóa tìm kiếm nhiều nhất 2019, xuất file top_search_key_2019.png

![image](https://user-images.githubusercontent.com/75520765/135628842-e3ad5823-a487-4b07-8feb-25e62d44cfbe.png)



