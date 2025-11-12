CHECK_PROMPT = lambda MONTH_HTML, SCHEDULE: f"""
Bạn là một chuyên gia kiểm tra lịch làm việc và xử lý DOM HTML bằng Playwright.
Mục tiêu: kiểm tra **tình trạng hiển thị** của tháng trong `MONTH_HTML` so với `SCHEDULE`. Nếu có sai lệch, xuất **chuỗi code Python Playwright** để click thay đổi từng ô ngày sao cho trạng thái **trùng với SCHEDULE**. Nếu không có sai lệch, trả về JSON: {{ "code": "" }}.

I. Dữ liệu đầu vào
1) HTML (MONTH_HTML) chứa bảng lịch (một tháng).
{MONTH_HTML}

2) Object lịch (SCHEDULE) chứa 3 nhóm: red_days, blue_days, black_days.
{SCHEDULE}


II. Luận lý và quy tắc rõ ràng (bắt buộc tuân thủ)
1. Trạng thái trong HTML (hiện tại):
   - Nếu ô `<td>` chứa class `pink_holiday` → trạng thái hiện tại = `red`
   - Nếu chứa class `blue_holiday` → trạng thái hiện tại = `blue`
   - Nếu chỉ có class `pointable` (không có `pink_holiday` và không có `blue_holiday`) → trạng thái hiện tại = `black`

2. Trạng thái mục tiêu (theo SCHEDULE):
   - Nếu ngày có trong `black_days[year][month]` → target = `black`
   - else if ngày có trong `red_days[year][month]` → target = `red`
   - else if ngày có trong `blue_days[year][month]` → target = `blue`
   - **Nếu một ngày xuất hiện trong nhiều nhóm của SCHEDULE, áp dụng ưu tiên:** `black` > `red` > `blue`.

3. Chu trình trạng thái khi click (1 click đổi sang trạng thái tiếp theo):
   - Vòng lặp trạng thái theo thứ tự: `red` (pink_holiday) → `blue` (blue_holiday) → `black` (pointable) → `red` → ...
   - Từ trạng thái hiện tại và target, tính số lần click cần thiết (0,1,2).

4. Selector để tìm ô ngày **không phụ thuộc** vào class hiện tại:
   - Dùng selector chung theo text: `current_calendar.locator('td.pointable:text-is("**date**")'), đảm bảo tìm đúng ô có nội dung số ngày.

III. Yêu cầu xuất code Playwright khi có sai lệch
- Xuất **một string** chứa **chương trình Python Playwright** (sync API) để click từng ô sai đủ số lần.
- Mẫu click cho một lần:
```python
cell = current_calendar.locator('td.pointable:text-is("**date**")')
cell.click()
```
"""