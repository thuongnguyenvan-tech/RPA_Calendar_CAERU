from playwright.sync_api import sync_playwright
import openai
import os
import re
import json
import sys
import pandas as pd
from collections import defaultdict

openai.api_key = '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'

def group_dates_by_year_month(date_list):
    """
    Nhóm các ngày theo năm và tháng (tháng là số nguyên).
    Input: list các chuỗi dạng 'YYYY-MM-DD'
    Output: dict { 'YYYY': { 1: [...], 2: [...], ... } }
    """
    result = defaultdict(lambda: defaultdict(list))
    
    for date_str in date_list:
        year, month, _ = date_str.split('-')
        month = int(month)  # chuyển tháng thành số nguyên
        result[year][month].append(date_str)
    
    # Chuyển defaultdict -> dict thường
    result = {year: dict(months) for year, months in result.items()}
    return result

def get_code_output(prompt, model="gpt-4o"):
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0,
        response_format={"type": "json_object"},
    )
    return response.choices[0].message["content"]

def get_result_from_json(result):
    temp = json.loads(result)
    return temp.get("code", temp.get("satisfy", ""))

def split_playwright_commands(code: str):
    parts = re.split(r';(?=(?:[^\'"]*[\'"][^\'"]*[\'"])*[^\'"]*$)', code)
    commands = [p.strip() for p in parts if p.strip()]
    return commands
def get_holidays_weekends(df):
    # Lọc các hàng không phải '平日'
    filtered_df = df[df['種別'] != '平日']
    
    # Khởi tạo dict kết quả
    result = {}
    
    # Lặp qua từng loại '種別' và lấy danh sách ngày
    for category in filtered_df['種別'].unique():
        result[category] = filtered_df.loc[filtered_df['種別'] == category, '日付'].tolist()
    
    return result

if __name__ == '__main__':
    file_path = "ISUZU_template.xlsx"
    sheet1 = pd.read_excel(file_path, sheet_name='法定休日')
    sheet2 = pd.read_excel(file_path, sheet_name='一般休日')
    sheet3 = pd.read_excel(file_path, sheet_name='パターン内容')
    URL = "https://test-newgps.caeru.biz/thuongAI/login"
    current_year = 2025
    init_code = f'page = browser.new_page(viewport={{"width": 1920, "height": 1080}});page.goto("{URL}", wait_until=\'domcontentloaded\');'

    #Sheet 1
    leave_schedule = group_dates_by_year_month(get_holidays_weekends(sheet1)["祝日"])
    sunday_schedule = group_dates_by_year_month(get_holidays_weekends(sheet1)["日曜"])
    saturday_schedule = group_dates_by_year_month(get_holidays_weekends(sheet1)["土曜"])

    # with sync_playwright() as playwright:
    #     browser = playwright.firefox.launch(headless=False, slow_mo=500,args=["--start-maximized"])
    #     page = browser.new_page()
    #     page.goto("https://test-newgps.caeru.biz/thuongAI/login", wait_until='domcontentloaded')
    #     page.fill('input[name="presentation_id"]', 'admin')
    #     page.fill('input[name="password"]', '1')
    #     page.click('button:has-text("ログイン")')
    #     page.wait_for_load_state("domcontentloaded")
    #     page.goto("https://test-newgps.caeru.biz/thuongAI/calendar", wait_until="domcontentloaded")

    #     #Click AREA tổng quát
    #     page.click('p.button > a.btn_gray:nth-of-type(1)')
    #     page.wait_for_timeout(3000)

    #     for month in range(1,13):
    #         current_calendar = page.locator('section.caeru_calendar_wrapper').filter(
    #             has=page.locator(f'span:text-is("{month}月")')
    #         )

    #         # # Chọn tất cả ô thứ Bảy (土) và chủ nhật có class pointable
    #         # saturdays = current_calendar.locator('td.pointable.saturday')
    #         # sundays = current_calendar.locator('td.pointable.sunday')

    #         # # Click ô thứ Bảy  và chủ nhật đầu tiên
    #         # saturdays.first.dblclick()
    #         # sundays.first.click()

    #         # # CLick Save
    #         # save_button = current_calendar.locator('a.btn_greeen:has-text("保存")')
    #         # save_button.click()
    #         try:
    #             for leave_day in leave_schedule[str(current_year)][month]:
    #                 day = leave_day[-2:]
    #                 day_cell = current_calendar.locator(f'td.pointable:text-is("{int(day)}")')
    #                 day_cell.click()
    #         except Exception as e:
    #             print(f"Không có ngày lễ ở tháng thứ {month}")

    #         try:
    #             for sundays in sunday_schedule[str(current_year)][month]:
    #                 day = sundays[-2:]
    #                 day_cell = current_calendar.locator(f'td.pointable:text-is("{int(day)}")')
    #                 day_cell.click()
    #         except Exception as e:
    #             print(f"Lỗi SUNDAY ở tháng thứ {month}")

    #         try:
    #             for saturdays in saturday_schedule[str(current_year)][month]:
    #                 day = saturdays[-2:]
    #                 day_cell = current_calendar.locator(f'td.pointable:text-is("{int(day)}")')
    #                 day_cell.dblclick()
    #         except Exception as e:
    #             print(f"Lỗi SATURDAY ở tháng thứ {month}")
    #     page.pause()
        # result = page.locator("body").evaluate("el => el.innerHTML")
    # print(result)

    

    # print(leave_schedule)

