from playwright.sync_api import sync_playwright
import openai
import os
import re
import json
import sys
import pandas as pd
from collections import defaultdict
from itertools import islice
openai.api_key = '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'

def group_dates_by_year_month(date_list):
    result = defaultdict(lambda: defaultdict(list))
    for date_str in date_list:
        year, month, _ = date_str.split('-')
        month = int(month)
        result[year][month].append(date_str)
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
    filtered_df = df[df['種別'] != '平日']
    result = {}
    for category in filtered_df['種別'].unique():
        result[category] = filtered_df.loc[filtered_df['種別'] == category, '日付'].tolist()
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

    #Sheet 2
    sheet2['パターン選択'] = sheet2['パターン選択'].apply(circled_number_to_int)
    areas = sheet2.set_index('勤務地ID')['パターン選択'].to_dict()

    #Sheet 3
    butan_list = parse_pattern_sheet_numeric(sheet3)

    #Final area list
    final_area_list = {k: butan_list[v] for k, v in areas.items()}

    # final_area_list = dict(islice(final_area_list.items(), 5))

    final_data = {}

    for area_id, pattern_data in final_area_list.items():
        work_days = pattern_data["出勤"]
        leave_days = pattern_data["休日"]

        red_days = list(set(get_holidays_weekends(sheet1)["祝日"] + get_holidays_weekends(sheet1)["日曜"]))
        blue_days = list(set(get_holidays_weekends(sheet1)["土曜"] + leave_days))
        black_days = work_days
        
        final_data[area_id] = {
            "red_days": group_dates_by_year_month(red_days),
            "blue_days": group_dates_by_year_month(blue_days),
            "black_days": group_dates_by_year_month(black_days)
        }

    with sync_playwright() as playwright:
        browser = playwright.firefox.launch(headless=False, slow_mo=500,args=["--start-maximized"])
        page = browser.new_page(viewport={{"width": 1920, "height": 1080}})
        page.goto(URL, wait_until='domcontentloaded')
        page.fill('input[name="presentation_id"]', ADMIN)
        page.fill('input[name="password"]', PASSWORD)
        page.click('button:has-text("ログイン")')
        page.wait_for_load_state("domcontentloaded")
        page.goto(URL + "calendar", wait_until="domcontentloaded")

        for area_ID, list_type_day in final_data.items():
        #     page.locator('a.modal-open.btn_gray', has_text="変更").click()
            input_locator = page.locator('div.search_setting:has(span:text-is("勤務地ID")) input')
            input_locator.fill(area_ID)
            input_locator.press("Enter")
            second_button = page.locator('a.ss_size.s_height.btn_gray:text-is("選択")').nth(1)
            second_button.click()
            for type_day, schedule in list_type_day.items():
                if type_day == "red_days":
                    for month, days in schedule[str(current_year)].items():
                        current_calendar = page.locator('section.caeru_calendar_wrapper').filter(has=page.locator(f'span:text-is("{month}月")')) 
                        for leaves in days:
                            day = leaves[-2:]
                            day_cell = current_calendar.locator(f'td.pointable:text-is("{int(day)}")')
                            day_cell.click()
                elif type_day == "blue_days":
                    for month, days in schedule[str(current_year)].items():
                        current_calendar = page.locator('section.caeru_calendar_wrapper').filter(has=page.locator(f'span:text-is("{month}月")')) 
                        for leaves in days:
                            day = leaves[-2:]
                            day_cell = current_calendar.locator(f'td.pointable:text-is("{int(day)}")')
                            day_cell.dblclick()
                else:
                    pass