from playwright.sync_api import sync_playwright
import os
import re
import json
import sys
import pandas as pd
from collections import defaultdict
from itertools import islice
from datetime import datetime
from prompt import CHECK_PROMPT
from bs4 import BeautifulSoup
import ctypes
from openpyxl import Workbook, load_workbook
import os
import time

def group_dates_by_year_month(date_list):
    result = defaultdict(lambda: defaultdict(list))
    for date_str in date_list:
        year, month, _ = date_str.split('-')
        month = int(month)
        year = int(year)
        result[year][month].append(date_str)
    result = {year: dict(months) for year, months in result.items()}
    return result

def get_holidays_weekends(df):
    result = {}
    for category in df['種別'].unique():
        result[category] = df.loc[df['種別'] == category, '日付'].tolist()
    return result

def circled_number_to_int(c):
    return ord(c) - ord('①') + 1

def parse_pattern_sheet_numeric(df):
    result = {}
    num_cols = df.shape[1]
    for col in range(0, num_cols, 2):
        pattern_name = df.iloc[0, col]
        if pd.isna(pattern_name):
            continue
        pattern_number = circled_number_to_int(pattern_name[0])
        pattern_dict = {"出勤": [], "休日": []}
        day_type_row = df.iloc[1, col]
        dates = df.iloc[2:, col].dropna()
        dates = [d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d) for d in dates]
        pattern_dict[day_type_row].extend(dates)
        if col + 1 < num_cols:
            adj_day_type = df.iloc[1, col + 1]
            if pd.notna(adj_day_type):
                adj_dates = df.iloc[2:, col + 1].dropna()
                adj_dates = [d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d) for d in adj_dates]
                pattern_dict[adj_day_type].extend(adj_dates)
        result[pattern_number] = pattern_dict
    return result

def extract_td_tags(html):
    soup = BeautifulSoup(html, 'html.parser')
    td_tags = soup.find_all('td')
    td_strings = [str(td) for td in td_tags if td.get_text(strip=True)]
    return td_strings

def check_calendar_days(calendar_html, schedule, current_year, month):
    results = []
    for td_html in calendar_html:
        soup = BeautifulSoup(td_html, 'html.parser')
        td = soup.find('td')
        day_text = td.text.strip()
        if not day_text.isdigit():
            results.append(False)
            continue
        date_str = f"{current_year}-{month:02d}-{int(day_text):02d}"
        classes = td.get('class', [])
        if "pink_holiday" in classes:
            results.append(date_str in schedule[current_year][month]['red_days'])
        elif "blue_holiday" in classes:
            results.append(date_str in schedule[current_year][month]['blue_days'])
        elif "pointable" in classes:
            results.append(date_str in schedule[current_year][month]['black_days'])
        else:
            results.append(False)
    return results

def find_day_category(day_str, schedule):
    for key, days in schedule.items():
        if day_str in days:
            return key
    return None

def get_click_count(td_html: str, target_status: str):
    if "pink_holiday" in td_html:
        current_status = "red_days"
    elif "blue_holiday" in td_html:
        current_status = "blue_days"
    else:
        current_status = "black_days"

    states = ["red_days", "blue_days", "black_days"]

    # Tính số bước để đến trạng thái mong muốn (theo vòng xoay)
    current_idx = states.index(current_status)
    target_idx = states.index(target_status)

    clicks = (target_idx - current_idx) % len(states)
    if clicks == 1:
        return day_cell.click()
    elif clicks == 2:
        return day_cell.dblclick()
    else:
        return "Error!"

def extract_info_from_info_sheet(df):
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    info_map = df.set_index('name')['value'].to_dict()

    def get(key):
        v = info_map.get(key)
        if pd.isna(v):
            return None
        return str(v).strip()
    
    url = get('URL')
    admin_id = get('管理者ID')
    password = get('password')

    raw_year = get('year')
    years = None
    if raw_year:
        parts = [p.strip() for p in raw_year.replace('，', ',').split(',')]
        parts = [p for p in parts if p]
        parsed = []
        for p in parts:
            try:
                parsed.append(int(p))
            except ValueError:
                parsed.append(p)
        years = parsed

    return {
        'URL': url,
        '管理者ID': admin_id,
        'password': password,
        'year': years
    }

def extract_weekend_days(tr_html_list, year_month):
    result = {"blue_days": [], "red_days": [], "black_days": []}
    table = [] 
    for tr_html in tr_html_list:
        soup = BeautifulSoup(tr_html, "html.parser")
        tds = soup.find_all("td")
        row = []
        for td in tds:
            text = td.get_text(strip=True)
            row.append(text)
        table.append(row)
    for weekday in table[1:]:
        for index, value in enumerate(weekday):
            if value != '' and index == 0:
                result["red_days"].append(f"{year_month}{int(weekday[index]):02d}")
            if value != '' and index == 6:
                result["blue_days"].append(f"{year_month}{int(weekday[index]):02d}")
            if value != '' and index != 0 and index != 6:
                result["black_days"].append(f"{year_month}{int(weekday[index]):02d}")
    return result

def update_day_patterns(A, B):
    holidays = set(A.get("red_days", []))
    leave_days = set(A.get("blue_days", []))
    work_days = set(A.get("black_days", []))
    sundays = set(B.get("red_days", []))
    saturdays = set(B.get("blue_days", []))
    normal_days = set(B.get("black_days", []))
    A["red_days"] = sorted((sundays | holidays) - work_days)
    A["blue_days"] = sorted((saturdays | leave_days) - work_days - holidays)
    A["black_days"] = sorted((normal_days | work_days)- leave_days - holidays) 
    return A

def generate_year_click_code(target_year: int, current_year: int):
    # if not (2000 <= target_year <= 2100) or not (2000 <= current_year <= 2100):
    #     return "# Lỗi: year ngoài giới hạn 2000-2100"
    diff = target_year - current_year
    if diff == 0:
        return 0
    if diff > 0:
        clicks = "\n".join(["button.nth(1).click()" for _ in range(diff)])
    elif diff < 0:
        clicks = "\n".join(["button.first.click()" for _ in range(abs(diff))])
    return clicks

def create_sheet_result_output(file_path: str,years):
    # Mở file để append
    wb = load_workbook(file_path)
    # Nếu sheet không tồn tại → tạo mới
    if "result" not in wb.sheetnames:
        wb.create_sheet("result")
    else:
        del wb["result"]
        wb.create_sheet("result")
    ws = wb["result"]
    ws.append(["勤務地名", "勤務地ID", "パターン選択"] + years)
    # Lưu file
    wb.save(file_path)

def add_row_sheet_result_output(file_path,rows):
    wb = load_workbook(file_path)
    ws = wb["result"]
    ws.append(rows)
    wb.save(file_path)

def add_block_sheet_result_output(file_path, row, column, value):
    wb = load_workbook(file_path)
    ws = wb["result"]
    ws.cell(row=row, column=column, value=value)
    wb.save(file_path)

if __name__ == '__main__':
    file_path = "ISUZU_template.xlsx"
    error = ""
    try:
        sheet1 = pd.read_excel(file_path, sheet_name='情報')
        sheet2 = pd.read_excel(file_path, sheet_name='法定休日')
        sheet3 = pd.read_excel(file_path, sheet_name='一般休日')
        sheet4 = pd.read_excel(file_path, sheet_name='パターン内容', parse_dates=True)
    except Exception as e:
        print(f"Không thể mở file {file_path}.")

    #Sheet 1
    information = extract_info_from_info_sheet(sheet1)
    create_sheet_result_output(file_path = file_path, years = information['year'])

    #Sheet 2
    leave_schedule = group_dates_by_year_month(get_holidays_weekends(sheet2)["祝日"])

    #Sheet 3
    sheet3['パターン選択'] = sheet3['パターン選択'].apply(circled_number_to_int)
    areas = sheet3.set_index('勤務地ID')['パターン選択'].to_dict()
    kinmuchi = sheet3.set_index('勤務地ID')['勤務地名'].to_dict()

    #Sheet 4
    butan_list = parse_pattern_sheet_numeric(sheet4)

    #Final area list
    final_area_list = {k: butan_list[v] for k, v in areas.items()}

    # final_area_list = dict(islice(final_area_list.items(), 5))

    master_schedule = {}

    for area_id, pattern_data in final_area_list.items():
        master_schedule[area_id] = defaultdict(lambda: defaultdict(dict))
        for year in leave_schedule.keys():
            for month in range(1, 13):
                leave_days = [
                        day for day in pattern_data.get("休日", [])
                        if datetime.strptime(day, "%Y-%m-%d").year == year and
                        datetime.strptime(day, "%Y-%m-%d").month == month
                    ]
                work_days = [
                        day for day in pattern_data.get("出勤", [])
                        if datetime.strptime(day, "%Y-%m-%d").year == year and
                        datetime.strptime(day, "%Y-%m-%d").month == month
                    ]
                
                holidays = leave_schedule.get(year, {}).get(month, [])

                red_days = sorted(set(holidays))

                blue_days = sorted(set(leave_days))

                black_days = sorted(set(work_days))

                master_schedule[area_id][year][month] = {
                    "red_days": red_days,
                    "blue_days": blue_days,
                    "black_days": black_days
                }

    master_schedule = {k: dict(v) for k, v in master_schedule.items()}

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False, slow_mo=500,args=["--kiosk"])
        page = browser.new_page(viewport={"width": 1920, "height": 1080})

        page.goto(information['URL'], wait_until='domcontentloaded')
        page.fill('input[name="presentation_id"]', information['管理者ID'])
        page.fill('input[name="password"]', information['password'])
        page.click('button:has-text("ログイン")')
        page.wait_for_load_state("domcontentloaded")
        page.goto(information['URL'] + "calendar", wait_until="domcontentloaded")

        first_loop = True
        row = 2

        #TODO: Thêm bước xác nhận nhập account có đúng không ở đây

        for area_ID, schedule in master_schedule.items():
            # area_ID = "S02000"
            column = 4
            add_row_sheet_result_output(file_path, [kinmuchi.get(area_ID), area_ID, areas.get(area_ID)])
            if not first_loop:
                page.locator('a.modal-open.btn_gray', has_text="変更").click()
            else:
                first_loop = False
            input_locator = page.locator('div.search_setting:has(span:text-is("勤務地ID")) input')
            input_locator.fill(area_ID)
            input_locator.press("Enter")
            try:
                second_button = page.locator('a.ss_size.s_height.btn_gray:text-is("選択")').nth(1)
            except Exception as e:
                add_block_sheet_result_output(file_path =file_path, row = row, column = column, value = f"Không có mã kinmuchi {area_ID}")
                print(f"Không có mã kinmuchi {area_ID}")
                first_loop = True
                row += 1
                continue
            second_button.click()
            
            for target_year in information['year']:
                # target_year = 2026
                year_text = page.locator('section.ll_font').inner_text()
                current_year = int(year_text.replace("年", "").strip())
                button = page.locator('section.ico_position img.ico_ico_arrow')

                clicks = generate_year_click_code(target_year = target_year, current_year = current_year)
                if clicks != 0:
                    exec(clicks)
                page.wait_for_load_state("domcontentloaded")
                time.sleep(2)
                error_months = []

                for month in range(1,13):
                    # month = 1
                    current_calendar = page.locator('section.caeru_calendar_wrapper').filter(has=page.locator(f'span:text-is("{month}月")'))
                    tr_elements = current_calendar.locator('table tr')
                    # Lấy nội dung HTML từng <tr>
                    tr_html_list = [tr_elements.nth(i).inner_html() for i in range(1,tr_elements.count())]
                    year_month = f"{target_year}-{month:02d}-"
                    weekends = (extract_weekend_days(tr_html_list,year_month))
                    schedule[target_year][month] = update_day_patterns(schedule[target_year][month], weekends)

                    # with open("master_schedule.txt", "w", encoding="utf-8") as f:
                    #     json.dump(schedule[target_year][month], f, ensure_ascii=False, indent=4)

                    calendar_html = current_calendar.inner_html()
                    td_list = extract_td_tags(calendar_html)[7:]

                    if "pink_holiday" not in "".join(td_list) and "pink_holiday" not in "".join(td_list):
                        for type_day, list_type_days in schedule[target_year][month].items():
                            try: 
                                for leaves in list_type_days:
                                    day = leaves[-2:]
                                    day_cell = current_calendar.locator(f'td.pointable:text-is("{int(day)}")')
                                    if type_day == "red_days":
                                        day_cell.click()
                                    elif type_day == "blue_days":
                                        day_cell.dblclick()
                            except Exception as e:
                                print(f"Không có {type_day} ở tháng thứ {month}.Error: {e}")

                    calendar_html = current_calendar.inner_html()
                    td_list = extract_td_tags(calendar_html)[7:]
                    result = check_calendar_days(td_list, schedule, target_year, month)

                    if False in result:
                        print(f"Tháng {month}: Error!")
                        for index, day in enumerate(result):
                            if day == False:
                                print(f"Lỗi ở ngày {index+1}")
                                current_error_day_status = td_list[index]
                                date_str = f"{target_year}-{month:02d}-{int(index+1):02d}"
                                target = find_day_category(date_str, schedule[target_year][month])
                                day_cell = current_calendar.locator(f'td.pointable:text-is("{int(index+1)}")')
                                get_click_count(current_error_day_status, target)
                                print(f"Đã sửa ngày {index+1} thành {target}")
                        
                    else:
                        print(f"Tháng {month}: OK!")

                    save_button = current_calendar.locator('a.btn_greeen:has-text("保存")')
                    save_button.click()
                    page.wait_for_load_state("domcontentloaded")
                    page.pause()
                
                if len(error_months) != 0:
                    add_block_sheet_result_output(file_path =file_path, row = row, column = column, value = "Error with " + ", ".join(map(str, error_months)))
                else:
                    add_block_sheet_result_output(file_path =file_path, row = row, column = column, value = "Done 12 months without error !")
                    
                column += 1
                # page.pause()
            row += 1


