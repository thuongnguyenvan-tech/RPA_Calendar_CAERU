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

if __name__ == '__main__':
    ADMIN = "admin"
    PASSWORD = "solomon"
    URL = "https://test-newgps.caeru.biz/isuzu/"

    file_path = "ISUZU_template.xlsx"
    sheet1 = pd.read_excel(file_path, sheet_name='法定休日')
    sheet2 = pd.read_excel(file_path, sheet_name='一般休日')
    sheet3 = pd.read_excel(file_path, sheet_name='パターン内容', parse_dates=True)

    current_year = 2025

    #Sheet 1
    leave_schedule = group_dates_by_year_month(get_holidays_weekends(sheet1)["祝日"])
    sunday_schedule = group_dates_by_year_month(get_holidays_weekends(sheet1)["日曜"])
    saturday_schedule = group_dates_by_year_month(get_holidays_weekends(sheet1)["土曜"])
    normal_schedule = group_dates_by_year_month(get_holidays_weekends(sheet1)["平日"])

    #Sheet 2
    sheet2['パターン選択'] = sheet2['パターン選択'].apply(circled_number_to_int)
    areas = sheet2.set_index('勤務地ID')['パターン選択'].to_dict()

    #Sheet 3
    butan_list = parse_pattern_sheet_numeric(sheet3)

    #Final area list
    final_area_list = {k: butan_list[v] for k, v in areas.items()}

    # final_area_list = dict(islice(final_area_list.items(), 5))

    master_schedule = {}

    for area_id, pattern_data in final_area_list.items():
        master_schedule[area_id] = defaultdict(lambda: defaultdict(dict))

        for year in leave_schedule.keys():
            for month in range(1, 13):
                normal_days = normal_schedule.get(year, {}).get(month, [])
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
                
                saturdays = saturday_schedule.get(year, {}).get(month, [])
                sundays = sunday_schedule.get(year, {}).get(month, [])
                holidays = leave_schedule.get(year, {}).get(month, [])

                red_days = sorted((set(sundays) | set(holidays)) - set(work_days))

                blue_days = sorted((set(saturdays) | set(leave_days)) - set(work_days))

                common_days = list(set(leave_days) & set(normal_days))

                black_days = sorted(set(normal_days) - set(common_days) | set(work_days))

                master_schedule[area_id][year][month] = {
                    "red_days": red_days,
                    "blue_days": blue_days,
                    "black_days": black_days
                }
                
    master_schedule = {k: dict(v) for k, v in master_schedule.items()}

    with open("master_schedule.txt", "w", encoding="utf-8") as f:
        json.dump(master_schedule, f, ensure_ascii=False, indent=4)

    with sync_playwright() as playwright:
        browser = playwright.firefox.launch(headless=False, slow_mo=500,args=["--start-maximized"])
        page = browser.new_page(viewport={"width": 1920, "height": 1080})
        page.goto(URL, wait_until='domcontentloaded')
        page.fill('input[name="presentation_id"]', ADMIN)
        page.fill('input[name="password"]', PASSWORD)
        page.click('button:has-text("ログイン")')
        page.wait_for_load_state("domcontentloaded")
        page.goto(URL + "calendar", wait_until="domcontentloaded")

        # #Click AREA tổng quát
        # page.click('p.button > a.btn_gray:nth-of-type(1)')
        # page.wait_for_timeout(3000)

        master_schedule = dict(islice(master_schedule.items(), 5))

        first_loop = True
        for area_ID, schedule in master_schedule.items():
            if not first_loop:
                page.locator('a.modal-open.btn_gray', has_text="変更").click()
            else:
                first_loop = False
            input_locator = page.locator('div.search_setting:has(span:text-is("勤務地ID")) input')
            input_locator.fill(area_ID)
            input_locator.press("Enter")
            second_button = page.locator('a.ss_size.s_height.btn_gray:text-is("選択")').nth(1)
            second_button.click()

            for month in range(1,13):
                for type_day, list_type_days in schedule[current_year][month].items():
                    try: 
                        current_calendar = page.locator('section.caeru_calendar_wrapper').filter(has=page.locator(f'span:text-is("{month}月")'))
                        for leaves in list_type_days:
                            day = leaves[-2:]
                            day_cell = current_calendar.locator(f'td.pointable:text-is("{int(day)}")')
                            if type_day == "red_days":
                                day_cell.click()
                            elif type_day == "blue_days":
                                day_cell.dblclick()
                    except Exception as e:
                        print(f"Không có {type_day} ở tháng thứ {month}")

                # page.pause()
                
                calendar_html = current_calendar.inner_html()
                td_list = extract_td_tags(calendar_html)[7:]

                result = check_calendar_days(td_list, schedule, current_year, month)

                if False in result:
                    print(f"Tháng {month}: Error!")
                    for index, day in enumerate(result):
                        if day == False:
                            print(f"Lỗi ở ngày {index+1}")
                            current_error_day_status = td_list[index]
                            # print(current_error_day_status)
                            date_str = f"{current_year}-{month:02d}-{int(index+1):02d}"
                            target = find_day_category(date_str, schedule[current_year][month])
                            day_cell = current_calendar.locator(f'td.pointable:text-is("{int(index+1)}")')
                            get_click_count(current_error_day_status, target)
                            print(f"Đã sửa ngày {index+1} thành {target}")
                            # print(target)
                else:
                    print(f"Tháng {month}: OK!")

                #Fix calendar by AI
                # temp_result = get_code_output(CHECK_PROMPT(td_list, schedule[current_year][month]))
                # CODE = get_result_from_json(temp_result)
                # print(f"CODE: {CODE}")
                page.pause()
                # Tìm và click vào nút 保存 bên trong tháng 1
                save_button = current_calendar.locator('a.btn_greeen:has-text("保存")')
                save_button.click()

