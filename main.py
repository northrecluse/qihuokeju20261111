from datetime import datetime, timedelta
import json
import os
import random
import re
import time
import urllib.parse
import chinese_calendar as chinesecalendar
import ddddocr
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment
from playwright.sync_api import sync_playwright

COOKIE_FILE = 'cookies.json'
OUTPUT_ROOT = '龙虎榜数据_长期历史汇总'


def get_current_time_str():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def get_trade_days_in_range(start_str, end_str):
    try:
        start_date = datetime.strptime(start_str.strip(), '%Y%m%d')
        end_date = datetime.strptime(end_str.strip(), '%Y%m%d')

        trade_days = []
        curr = start_date
        while curr <= end_date:
            if chinesecalendar.is_workday(curr) and curr.weekday() < 5:
                trade_days.append(curr.strftime('%Y-%m-%d'))
            curr += timedelta(days=1)

        return sorted(trade_days)
    except Exception as e:
        print(f'[{get_current_time_str()}] ❌ 日期范围解析异常 ({e})')
        return []


def input_with_timeout(prompt, timeout=10):
    import sys
    import threading

    print(f'[{get_current_time_str()}] {prompt}', end='', flush=True)
    user_input = [None]

    def target():
        try:
            user_input[0] = sys.stdin.readline().strip()
        except Exception:
            pass

    thread = threading.Thread(target=target)
    thread.daemon = True
    thread.start()

    for i in range(timeout, 0, -1):
        if not thread.is_alive():
            break
        print(
            f'\r[{get_current_time_str()}] {prompt} (剩余 {i} 秒): ',
            end='',
            flush=True,
        )
        time.sleep(1)

    print()
    if thread.is_alive():
        print(
            f'[{get_current_time_str()}] ⌛ 输入超时，将自动使用默认长期范围 (2024.1.1 - 2026.10.31)！'
        )
        return None

    return user_input[0] if user_input[0] else None


def parse_net_position(val_str):
    if not val_str or not isinstance(val_str, str):
        return 0

    clean_str = re.sub(r'[^\d.-]', '', val_str.replace(',', ''))
    try:
        num = float(clean_str)
    except ValueError:
        num = 0.0

    if '空' in val_str:
        return -abs(num)
    elif '多' in val_str:
        return abs(num)
    return num


def parse_change_val(val_str):
    if not val_str or not isinstance(val_str, str):
        return 0.0
    clean_str = re.sub(r'[^\d.-]', '', val_str.replace(',', ''))
    try:
        return float(clean_str)
    except ValueError:
        return 0.0


def check_company_match(seat_name, target_companies):
    for company in target_companies:
        if company in seat_name:
            return True
    return False


def safe_goto(page, url, timeout=60000, wait_until='domcontentloaded'):
    for attempt in range(1, 4):
        try:
            page.goto(url, wait_until=wait_until, timeout=timeout)
            page.wait_for_timeout(2000)
            return True
        except Exception as e:
            print(
                f'[{get_current_time_str()}] ⚠️ 打开页面超时或异常 ({e})，正在重试...'
            )
            page.wait_for_timeout(2000)
    return False


def auto_login_and_close_ad(page, context, username, password):
    # 打印当前加载的 Cookie 数量排查问题
    try:
        cookies = context.cookies()
        print(f'[{get_current_time_str()}] 🍪 当前上下文已加载 Cookie 数量: {len(cookies)}')
    except Exception as e:
        print(f'[{get_current_time_str()}] ⚠️ 获取 Cookie 列表异常: {e}')

    if os.path.exists(COOKIE_FILE):
        print(
            f'[{get_current_time_str()}] 检测到本地 Cookie 文件，正在尝试免密直接登录...'
        )
        try:
            with open(COOKIE_FILE, 'r', encoding='utf-8') as f:
                cookies = json.load(f)
            context.add_cookies(cookies)

            safe_goto(
                page,
                'https://x.qhkch.com/variety/position?variety=%E8%9E%BA%E7%BA%B9%E9%92%A2&code=all&date=2024-01-02',
            )
            page.wait_for_timeout(4000)

            # 关键检查：如果跳转到了登录页，说明 Cookie 失效
            if 'login' not in page.url and 'passport' not in page.url:
                print(
                    f'[{get_current_time_str()}] 🎉 Cookie 有效！成功实现免登录直接进入系统！'
                )
                return page
            else:
                print(
                    f'[{get_current_time_str()}] ⚠️ Cookie 已失效，当前页面 URL: {page.url}，将转入账号密码自动登录...'
                )
        except Exception as e:
            print(f'[{get_current_time_str()}] ⚠️ Cookie 加载异常: {e}')

    print(f'[{get_current_time_str()}] 正在打开登录页面...')
    login_url = 'https://passport.qhkch.com/login?continuation_id=qhkch_authz_2HgWV1yGs3Ryw3-l_swF2Vj3RWleWNm99eSWKZupzlc'
    if not safe_goto(page, login_url):
        return None

    ocr = ddddocr.DdddOcr(show_ad=False)
    for attempt in range(1, 11):
        try:
            print(
                f'\n[{get_current_time_str()}] --- 第 {attempt} 次尝试登录 ---'
            )
            page.wait_for_selector(
                'input:not([type="hidden"]):not([type="password"])',
                timeout=15000,
            )

            user_locator = page.locator(
                'input:not([type="hidden"]):not([type="password"])'
            ).first
            pwd_locator = page.locator('input[type="password"]').first

            user_locator.click()
            user_locator.fill('')
            user_locator.type(username, delay=30)
            user_locator.dispatch_event('input')

            pwd_locator.click()
            pwd_locator.fill('')
            pwd_locator.type(password, delay=30)
            pwd_locator.dispatch_event('input')

            captcha_locator = page.locator(
                "img[src*='captcha'], img[class*='code'], .el-input + img, canvas,"
                ' form img'
            ).first
            captcha_bytes = captcha_locator.screenshot()
            captcha_code = ocr.classification(captcha_bytes)
            print(
                f'[{get_current_time_str()}] -> 【ddddocr】识别验证码结果: {captcha_code}'
            )

            code_locator = page.locator(
                "input[placeholder*='验证码'], input:not([type='hidden']):not([type='password'])"
            ).nth(1)
            if code_locator.count() == 0:
                code_locator = page.locator('input:not([type="hidden"])').nth(
                    2
                )

            code_locator.click()
            code_locator.fill('')
            code_locator.type(captcha_code, delay=30)
            code_locator.dispatch_event('input')
            page.wait_for_timeout(500)

            login_btn = page.locator(
                "button:has-text('登录'), .el-button--primary, button"
            ).first
            if login_btn.count() > 0:
                login_btn.click(force=True)
            else:
                page.locator('button').last.click(force=True)

            page.wait_for_timeout(4000)

            if (
                'Passport' in page.title()
                or page.locator("text='奇货可查'").count() > 0
            ):
                try:
                    with context.expect_page(timeout=5000) as new_page_info:
                        page.locator("text='奇货可查'").first.click(force=True)
                    page = new_page_info.value
                    page.wait_for_load_state()
                except Exception:
                    pass

            if 'login' not in page.url and 'passport' not in page.url:
                print(f'[{get_current_time_str()}] 🎉 恭喜！成功登录系统！')
                try:
                    cookies = context.cookies()
                    with open(COOKIE_FILE, 'w', encoding='utf-8') as f:
                        json.dump(cookies, f, ensure_ascii=False, indent=4)
                except Exception:
                    pass
                return page
            else:
                captcha_locator.click()
                page.wait_for_timeout(1000)
        except Exception as e:
            print(f'[{get_current_time_str()}] ❌ 登录异常: {e}')
            page.wait_for_timeout(1000)

    return None


def get_configured_varieties():
    varieties_list = [
        '1000中证/IM',
        '20号胶/NR',
        '300沪深/IF',
        '500中证/IC',
        '50上证/IH',
        'PP/PP',
        'PTA/TA',
        'PVC/V',
        '三十债/TL',
        '不锈钢/SS',
        '丙烯/PL',
        '乙二醇/EG',
        '二债/TS',
        '五债/TF',
        '十债/T',
        '双胶纸/OP',
        '合成橡胶/BR',
        '国际铜/BC',
        '塑料/L',
        '多晶硅/PS',
        '天然橡胶/RU',
        '对二甲苯/PX',
        '尿素/UR',
        '工业硅/SI',
        '棉纱/CY',
        '棉花/CF',
        '棕榈油/P',
        '氧化铝/AO',
        '沥青/BU',
        '沪金/AU',
        '沪铅/PB',
        '沪铜/CU',
        '沪铝/AL',
        '沪银/AG',
        '沪锌/ZN',
        '沪锡/SN',
        '沪镍/NI',
        '液化气/PG',
        '淀粉/CS',
        '烧碱/SH',
        '热卷/HC',
        '焦炭/J',
        '焦煤/JM',
        '燃料油/FU',
        '玉米/C',
        '玻璃/FG',
        '瓶片/PR',
        '生猪/LH',
        '甲醇/MA',
        '白糖/SR',
        '短纤/PF',
        '硅铁/SF',
        '碳酸锂/LC',
        '红枣/CJ',
        '纯碱/SA',
        '纯苯/BZ',
        '纸浆/SP',
        '花生/PK',
        '苯乙烯/EB',
        '苹果/AP',
        '菜油/OI',
        '菜粕/RM',
        '螺纹钢/RB',
        '豆一/A',
        '豆二/B',
        '豆油/Y',
        '豆粕/M',
        '铁矿石/I',
        '铝合金/AD',
        '锰硅/SM',
        '鸡蛋/JD',
    ]
    return sorted(list(set(varieties_list)))


def load_existing_data_for_date(output_folder, date_str):
    excel_path = os.path.join(
        output_folder, f'全品种期货持仓龙虎榜汇总_{date_str}.xlsx'
    )
    parsed_variety_data = {}
    if not os.path.exists(excel_path):
        return parsed_variety_data

    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True)
        sheets = wb.sheetnames
        for s_name in sheets:
            if s_name == '统计汇总':
                continue
            df_full = pd.read_excel(excel_path, sheet_name=s_name)
            main_contract = '全合约'
            try:
                wb_full = openpyxl.load_workbook(excel_path, data_only=True)
                ws = wb_full[s_name]
                title_cell = ws.cell(row=1, column=1).value
                if title_cell and '合约:' in title_cell:
                    main_contract = title_cell.split('合约:')[1].strip()
            except Exception:
                pass

            cols = df_full.columns.tolist()
            if len(cols) >= 4:
                df_long = df_full.iloc[:, :4].copy()
                df_long.columns = [
                    '席位/期货公司',
                    '持仓量',
                    '增减量',
                    '净持仓',
                ]
                df_long = df_long.dropna(subset=['席位/期货公司'])
                df_long['num_val'] = df_long['净持仓'].apply(
                    lambda x: parse_net_position(str(x))
                )
            else:
                df_long = pd.DataFrame()

            if len(cols) >= 9:
                df_short = df_full.iloc[:, 5:9].copy()
                df_short.columns = [
                    '席位/期货公司',
                    '持仓量',
                    '增减量',
                    '净持仓',
                ]
                df_short = df_short.dropna(subset=['席位/期货公司'])
                df_short['num_val'] = df_short['净持仓'].apply(
                    lambda x: parse_net_position(str(x))
                )
            else:
                df_short = pd.DataFrame()

            parsed_variety_data[s_name] = {
                'long': df_long,
                'short': df_short,
                'main_contract': main_contract,
            }
    except Exception as e:
        print(f'[{get_current_time_str()}] ⚠️ 恢复历史缓存数据异常: {e}')

    return parsed_variety_data


def export_excel_for_date(parsed_variety_data, date_str, output_folder):
    output_file = os.path.join(
        output_folder, f'全品种期货持仓龙虎榜汇总_{date_str}.xlsx'
    )

    group1_targets = ['永安期货', '国泰君安', '东证期货', '高盛期货']
    group2_targets = ['方正中期', '方正', '东方财富', '平安期货', '徽商期货']

    summary_records = []

    for variety_name, v_data in parsed_variety_data.items():
        df_l = v_data['long']
        df_s = v_data['short']
        main_contract = v_data['main_contract']

        if df_l.empty and df_s.empty:
            continue

        df_l_top10 = df_l.head(10) if not df_l.empty else pd.DataFrame()
        df_s_top10 = df_s.head(10) if not df_s.empty else pd.DataFrame()

        g1_l_count = sum(
            1
            for _, r in df_l_top10.iterrows()
            if check_company_match(r['席位/期货公司'], group1_targets)
        )
        g1_l_inc = sum(
            1
            for _, r in df_l_top10.iterrows()
            if check_company_match(r['席位/期货公司'], group1_targets)
            and parse_change_val(r.get('增减量', '')) > 0
        )

        g1_s_count = sum(
            1
            for _, r in df_s_top10.iterrows()
            if check_company_match(r['席位/期货公司'], group1_targets)
        )
        g1_s_inc = sum(
            1
            for _, r in df_s_top10.iterrows()
            if check_company_match(r['席位/期货公司'], group1_targets)
            and parse_change_val(r.get('增减量', '')) > 0
        )

        g2_l_count = sum(
            1
            for _, r in df_l_top10.iterrows()
            if check_company_match(r['席位/期货公司'], group2_targets)
        )
        g2_l_inc = sum(
            1
            for _, r in df_l_top10.iterrows()
            if check_company_match(r['席位/期货公司'], group2_targets)
            and parse_change_val(r.get('增减量', '')) > 0
        )

        g2_s_count = sum(
            1
            for _, r in df_s_top10.iterrows()
            if check_company_match(r['席位/期货公司'], group2_targets)
        )
        g2_s_inc = sum(
            1
            for _, r in df_s_top10.iterrows()
            if check_company_match(r['席位/期货公司'], group2_targets)
            and parse_change_val(r.get('增减量', '')) > 0
        )

        summary_records.append({
            '日期': date_str,
            '品种': variety_name,
            '主力合约': main_contract,
            '组1(永安/国泰/东证/高盛) 多头 Top10 出现数': g1_l_count,
            '组1多头增仓': g1_l_inc,
            '组1(永安/国泰/东证/高盛) 空头 Top10 出现数': g1_s_count,
            '组1空头增仓': g1_s_inc,
            '组2(方正中期/东财/平安/徽商) 多头 Top10 出现数': g2_l_count,
            '组2多头增仓': g2_l_inc,
            '组2(方正中期/东财/平安/徽商) 空头 Top10 出现数': g2_s_count,
            '组2空头增仓': g2_s_inc,
        })

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        if summary_records:
            df_summary = pd.DataFrame(summary_records)
            df_summary['__sort_key1'] = (
                df_summary['组1(永安/国泰/东证/高盛) 多头 Top10 出现数']
                + df_summary['组1(永安/国泰/东证/高盛) 空头 Top10 出现数']
            )
            df_summary = df_summary.sort_values(
                by=[
                    '__sort_key1',
                    '组1(永安/国泰/东证/高盛) 多头 Top10 出现数',
                ],
                ascending=False,
            ).reset_index(drop=True)
            df_summary = df_summary.drop(columns=['__sort_key1'])
            df_summary.to_excel(writer, sheet_name='统计汇总', index=False)
        else:
            pd.DataFrame([{'提示': '该日期无可用持仓数据'}]).to_excel(
                writer, sheet_name='统计汇总', index=False
            )

        for variety_name, v_data in parsed_variety_data.items():
            df_l = v_data['long']
            df_s = v_data['short']
            main_contract = v_data['main_contract']

            if df_l.empty and df_s.empty:
                continue

            clean_l = (
                df_l.drop(columns=['num_val']) if 'num_val' in df_l.columns else df_l
            )
            clean_s = (
                df_s.drop(columns=['num_val']) if 'num_val' in df_s.columns else df_s
            )

            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', variety_name)[:31]

            if not clean_l.empty:
                clean_l.to_excel(
                    writer,
                    sheet_name=safe_sheet_name,
                    startcol=0,
                    startrow=1,
                    index=False,
                )
            if not clean_s.empty:
                clean_s.to_excel(
                    writer,
                    sheet_name=safe_sheet_name,
                    startcol=5,
                    startrow=1,
                    index=False,
                )

            ws = writer.sheets[safe_sheet_name]
            ws.cell(
                row=1,
                column=1,
                value=f'【多头龙虎榜 (按净持仓从多到少)】 日期: {date_str} 合约: {main_contract}',
            )
            ws.cell(
                row=1,
                column=6,
                value=f'【空头龙虎榜 (按净持仓从少到多)】 日期: {date_str} 合约: {main_contract}',
            )

    wb = openpyxl.load_workbook(output_file)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value is not None:
                    length = sum(
                        2 if ord(char) > 127 else 1 for char in str(cell.value)
                    )
                    if length > max_len:
                        max_len = length
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output_file)
    print(
        f'[{get_current_time_str()}] ✅ 交易日【{date_str}】的所有品种已完整获取，汇总文件已保存: {output_file}'
    )


def main():
    print('=' * 60)
    user_input_raw = input_with_timeout(
        '请输入想要抓取的时间范围 (默认回车使用 20240101-20261031): ',
        timeout=10,
    )
    print('=' * 60)

    if not user_input_raw or '-' not in user_input_raw:
        start_str = '20240101'
        end_str = '20261031'
    else:
        parts = user_input_raw.split('-')
        start_str = parts[0].strip()
        end_str = parts[1].strip()

    target_dates = get_trade_days_in_range(start_str, end_str)
    print(
        f'[{get_current_time_str()}] 📌 目标时间段 ({start_str} 至 {end_str}) 内共有交易日: {len(target_dates)} 天'
    )

    username = os.environ.get('QHKCH_USER', '18392272125')
    password = os.environ.get('QHKCH_PWD', 'lmn123')

    output_folder = OUTPUT_ROOT
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    is_ci = os.environ.get('CI') == 'true'

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=is_ci,
            slow_mo=300 if not is_ci else 0,
            args=[
                '--ignore-certificate-errors',
                '--ignore-ssl-errors',
                '--allow-insecure-localhost',
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-blink-features=AutomationControlled',
            ],
        )
        context = browser.new_context(
            ignore_https_errors=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        page.set_default_timeout(60000)
        page.set_default_navigation_timeout(60000)

        page = auto_login_and_close_ad(page, context, username, password)
        if not page:
            print(
                f'[{get_current_time_str()}] ❌ 登录状态获取失败，退出程序！'
            )
            # 登录失败时截个图保存，方便在 GitHub Artifacts 里查看
            try:
                page.screenshot(path='login_fail.png')
            except Exception:
                pass
            browser.close()
            return

        raw_varieties = get_configured_varieties()

        for d_idx, date_str in enumerate(target_dates, 1):
            parsed_variety_data = load_existing_data_for_date(
                output_folder, date_str
            )

            if len(parsed_variety_data) >= len(raw_varieties):
                print(
                    f'[{get_current_time_str()}] ⏩ [交易日进度 {d_idx}/{len(target_dates)}] 日期 {date_str} 已完整抓取过 ({len(parsed_variety_data)}/{len(raw_varieties)}品种)，自动跳过！'
                )
                continue

            print(
                f"\n{'=' * 20} [{get_current_time_str()}] [交易日进度 {d_idx}/{len(target_dates)}] 开始处理日期: {date_str} (已缓存 {len(parsed_variety_data)} 个品种) {'=' * 20}"
            )

            for index, raw_variety in enumerate(raw_varieties, 1):
                variety_search_kw = (
                    raw_variety.split('/')[0] if '/' in raw_variety else raw_variety
                )
                variety_code = (
                    raw_variety.split('/')[1] if '/' in raw_variety else ''
                )
                display_variety_name = raw_variety

                if display_variety_name in parsed_variety_data:
                    continue

                long_rows = []
                short_rows = []
                main_contract_str = '全合约'
                start_t = time.time()

                try:
                    encoded_variety = urllib.parse.quote(variety_search_kw)
                    target_url = f'https://x.qhkch.com/variety/position?variety={encoded_variety}&code=all&date={date_str}'
                    safe_goto(page, target_url)

                    start_time = time.time()
                    has_data = False
                    while time.time() - start_time < 8:
                        page.wait_for_timeout(400)
                        page_status = page.evaluate("""
                            () => {
                                const bodyText = document.body.innerText || '';
                                const tables = document.querySelectorAll('table');
                                let hasTableWithData = false;
                                for (let t of tables) {
                                    if (t && t.querySelectorAll && t.querySelectorAll('tr').length > 1) {
                                        hasTableWithData = true;
                                        break;
                                    }
                                }
                                const noData = bodyText.includes('暂无数据') || bodyText.includes('没有数据');
                                const isLogin = bodyText.includes('登录') && bodyText.includes('密码');
                                return { hasData: hasTableWithData, noData: noData, isLogin: isLogin };
                            }
                        """)
                        # 如果不小心被踢回登录页，打印提示
                        if page_status['isLogin']:
                            print(f'[{get_current_time_str()}] ⚠️ 警告：检测到页面跳转回了登录页，可能登录态已失效！')
                        if page_status['noData']:
                            break
                        if page_status['hasData']:
                            has_data = True
                            break

                    if has_data:
                        page_info = page.evaluate(
                            """
                            (varietyCode) => {
                                let contractStr = "";
                                const contractInput = document.querySelector('input[placeholder*="合约"], .el-select input');
                                if (contractInput && contractInput.value) {
                                    let val = contractInput.value.trim();
                                    if (/^[a-zA-Z]+\\d{3,4}$/.test(val)) {
                                        contractStr = val;
                                    }
                                }
                                if (!contractStr) {
                                    const activeTags = document.querySelectorAll('.el-tag, .is-active, .active, button');
                                    for (let tag of activeTags) {
                                        let text = tag ? tag.innerText.trim() : "";
                                        if (/^[a-zA-Z]{1,3}\\d{3,4}$/.test(text)) {
                                            contractStr = text;
                                            break;
                                        }
                                    }
                                }
                                if (!contractStr) {
                                    const bodyText = document.body.innerText || '';
                                    const match = bodyText.match(/([a-zA-Z]{1,3}\\d{3,4})/);
                                    if (match) {
                                        contractStr = match[1];
                                    }
                                }
                                return { main_contract: contractStr };
                            }
                        """,
                            variety_code,
                        )

                        raw_contract = page_info.get('main_contract', '')
                        if raw_contract and variety_code:
                            clean_prefix = re.sub(
                                r'\d', '', raw_contract
                            ).upper()
                            if (
                                variety_code.upper() not in clean_prefix
                                and clean_prefix not in variety_code.upper()
                            ):
                                main_contract_str = '全合约'
                            else:
                                main_contract_str = raw_contract
                        else:
                            main_contract_str = (
                                raw_contract if raw_contract else '全合约'
                            )

                        table_data = page.evaluate("""
                            () => {
                                let longData = [];
                                let shortData = [];
                                const tables = document.querySelectorAll('table');
                                if (!tables || tables.length === 0) return { long_table: [], short_table: [] };

                                tables.forEach((table) => {
                                    if (!table || !table.parentElement) return;
                                    let parent = table.parentElement;
                                    let cardText = "";
                                    for (let i = 0; i < 3; i++) {
                                        if (parent && typeof parent.innerText === 'string') {
                                            cardText += " " + parent.innerText;
                                            parent = parent.parentElement;
                                        } else {
                                            break;
                                        }
                                    }

                                    const rows = Array.from(table.querySelectorAll('tr')).map(tr => {
                                        return Array.from(tr.querySelectorAll('th, td')).map(td => td ? td.innerText.trim() : "");
                                    }).filter(cells => cells.length > 0);

                                    if ((cardText.includes('多头') || cardText.includes('多头持仓')) && longData.length === 0 && rows.length > 1) {
                                        longData = rows;
                                    } 
                                    else if ((cardText.includes('空头') || cardText.includes('空头持仓')) && shortData.length === 0 && rows.length > 1) {
                                        shortData = rows;
                                    }
                                });

                                if (longData.length === 0 && tables.length > 0 && tables[0]) {
                                    longData = Array.from(tables[0].querySelectorAll('tr')).map(tr => 
                                        Array.from(tr.querySelectorAll('th, td')).map(td => td ? td.innerText.trim() : "")
                                    ).filter(cells => cells.length > 0);
                                }
                                if (shortData.length === 0 && tables.length > 1 && tables[1]) {
                                    shortData = Array.from(tables[1].querySelectorAll('tr')).map(tr => 
                                        Array.from(tr.querySelectorAll('th, td')).map(td => td ? td.innerText.trim() : "")
                                    ).filter(cells => cells.length > 0);
                                }

                                return { long_table: longData, short_table: shortData };
                            }
                        """)

                        long_rows = table_data.get('long_table', [])
                        short_rows = table_data.get('short_table', [])

                except Exception as e:
                    print(
                        f'[{get_current_time_str()}] ⚠️ 品种 {display_variety_name} 异常: {e}'
                    )

                long_list = []
                for row in long_rows:
                    if len(row) >= 2 and row[0] not in ['席位', '持仓', '合计', '']:
                        net_val = row[3] if len(row) > 3 else ''
                        long_list.append({
                            '席位/期货公司': row[0],
                            '持仓量': row[1] if len(row) > 1 else '',
                            '增减量': row[2] if len(row) > 2 else '',
                            '净持仓': net_val,
                            'num_val': parse_net_position(net_val),
                        })

                short_list = []
                for row in short_rows:
                    if len(row) >= 2 and row[0] not in ['席位', '持仓', '合计', '']:
                        net_val = row[3] if len(row) > 3 else ''
                        short_list.append({
                            '席位/期货公司': row[0],
                            '持仓量': row[1] if len(row) > 1 else '',
                            '增减量': row[2] if len(row) > 2 else '',
                            '净持仓': net_val,
                            'num_val': parse_net_position(net_val),
                        })

                df_long = pd.DataFrame(long_list)
                if not df_long.empty:
                    df_long = df_long.sort_values(
                        by='num_val', ascending=False
                    ).reset_index(drop=True)

                df_short = pd.DataFrame(short_list)
                if not df_short.empty:
                    df_short = df_short.sort_values(
                        by='num_val', ascending=True
                    ).reset_index(drop=True)

                parsed_variety_data[display_variety_name] = {
                    'long': df_long,
                    'short': df_short,
                    'main_contract': main_contract_str,
                }

                elapsed = time.time() - start_t
                print(
                    f'[{get_current_time_str()}] [{index}/{len(raw_varieties)}] 日期: {date_str} | 品种: {display_variety_name:<10} | 主力合约: {main_contract_str:<8} | 多头: {len(df_long):2d}条 | 空头: {len(df_short):2d}条 | 耗时: {elapsed:.1f}s'
                )

                time.sleep(random.uniform(0.3, 0.6))

            export_excel_for_date(parsed_variety_data, date_str, output_folder)

        browser.close()

    print(
        f"\n{'=' * 20} [{get_current_time_str()}] 全部历史抓取任务完成 {'=' * 20}"
    )
    print(
        f'[{get_current_time_str()}] 📁 所有 Excel 文件已存至文件夹: {OUTPUT_ROOT}'
    )


if __name__ == '__main__':
    main()
