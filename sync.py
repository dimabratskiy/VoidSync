
import os
import time
import re
import locale
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Устанавливаем локаль для парсинга чисел и дат
locale.setlocale(locale.LC_NUMERIC, 'ru_RU.UTF-8')
locale.setlocale(locale.LC_TIME, 'ru_RU.UTF-8')

# === Конфигурация проекта ===
SHEET_NAME       = "test"               # имя Google-таблицы
WORKSHEET        = "Лист1"             # имя листа в таблице
CREDENTIALS_FILE = "credentials.json"   # путь к JSON для сервис-аккаунта
DEBUG_PAGE_LIMIT = 10                    # лимит страниц при отладке

# Авторизация в Google Sheets
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
gc    = gspread.authorize(creds)
sh    = gc.open(SHEET_NAME)
ws    = sh.worksheet(WORKSHEET)

# Проверяем, есть ли уже данные на листе (ключевой ID во второй строке)
records = ws.get_all_records()
key_id  = records[0].get('ID заявки') if records else None
print("Ключевой ID (для инкрементации):", key_id)

# Настраиваем Chrome с помощью webdriver-manager и портативного Chrome
chrome_binary = os.path.join(os.getcwd(), 'chrome-portable', 'Chrome.exe')
options = Options()
if os.path.exists(chrome_binary):
    options.binary_location = chrome_binary
# сохраняем профиль рядом
options.add_argument(f"--user-data-dir={os.path.join(os.getcwd(), 'profile')}")
#options.add_argument('--headless')  # раскомментировать для безголового режима

driver_path = ChromeDriverManager(path=os.path.join(os.getcwd(), 'drivers')).install()
driver = webdriver.Chrome(executable_path=driver_path, options=options)

def inject_clipboard_override():
    """Подменяем navigator.clipboard.writeText внутри страницы для сохранения ID."""
    script = '''
        if (navigator.clipboard) {
            navigator.clipboard.writeText = function(text) {
                window._lastCopiedId = text;
                return Promise.resolve();
            };
        }
    '''
    driver.execute_script(script)

# Функция разбора текущей страницы

def parse_current_page():
    rows = []
    trs = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
    for tr in trs:
        tds = tr.find_elements(By.TAG_NAME, 'td')
        # Дата и время
        dt = tds[1].text.strip().split()
        date_raw = ' '.join(dt[:-1])  # "28 апр. 2025"
        time_raw = dt[-1]             # "21:40"
        # конвертируем в нужный формат DD.MM.YYYY
        try:
            d_obj = datetime.strptime(date_raw, '%d %b. %Y')
        except ValueError:
            d_obj = datetime.strptime(date_raw, '%d %b %Y')
        date_fmt = d_obj.strftime('%d.%m.%Y')
        time_fmt = time_raw  # уже HH:MM

        # Сумма ₽
        m_rub = re.search(r'[\d\s,]+₽', tds[2].text)
        rub   = m_rub.group().replace('₽', '').replace(' ', '') if m_rub else '0'
        rub_val = locale.atof(rub)
        rub_fmt = f"{rub_val:.2f}".replace('.', ',')

        # Сумма $ (USDT)
        usd_txt = tds[2].find_element(By.CSS_SELECTOR, 'div.font-semibold').text
        usd     = usd_txt.replace('USDT', '').replace(' ', '')
        usd_val = locale.atof(usd)
        usd_fmt = f"{usd_val:.2f}".replace('.', ',')

        # Курс ₽/USDT
        rate_val = rub_val / usd_val if usd_val else 0
        rate     = f"{rate_val:.2f}".replace('.', ',')

        # Доход $ (USDT)
        prof_txt = tds[3].find_element(By.CSS_SELECTOR, 'div.font-semibold').text
        prof     = prof_txt.replace('USDT', '').replace(' ', '')
        prof_val = locale.atof(prof)
        prof_fmt = f"{prof_val:.2f}".replace('.', ',')

        # Процент прибыли
        profit_pct = '6%'

        # Телефон без пробелов и '+'
        phone = ''
        try:
            grp = tds[5].find_element(By.CSS_SELECTOR, 'div.group')
            ph  = grp.find_elements(By.XPATH, './div[2]/div')[0].text.strip()
            phone = ph.replace(' ', '').lstrip('+')
        except Exception:
            pass

        # ID заявки через clipboard override
        try:
            btn = tds[-1].find_element(By.XPATH, ".//div[contains(@class,'cursor-pointer')]")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.1)
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(0.1)
            req_id = driver.execute_script('return window._lastCopiedId || "";')
        except Exception:
            req_id = tds[-1].find_element(By.CSS_SELECTOR, 'div.overflow-hidden').get_attribute('textContent').strip()

        # Статус
        status = 'Оплаченная' if 'Оплачено' in tds[4].text else 'Отменённая'

        rows.append([
            date_fmt,
            time_fmt,
            req_id,
            phone,
            rub_fmt,
            rate,
            usd_fmt,
            profit_pct,
            prof_fmt,
            status
        ])
    return rows

# Главная функция синхронизации
def main():
    driver.get('https://cabinet.voidpay.biz/')
    time.sleep(5)
    inject_clipboard_override()

    # Шапка таблицы
    if not key_id:
        ws.clear()
        ws.append_row([
            'Дата', 'Время', 'ID заявки', 'Телефон',
            'Сумма (₽)', 'Курс', 'Сумма ($)', 'Доход %', 'Доход ($)', 'Статус'
        ], value_input_option='USER_ENTERED')
        # Полная синхронизация
        page = 0
        while True:
            page += 1
            print(f'Полная синхронизация, страница {page}')
            rows = parse_current_page()
            if rows:
                ws.append_rows(rows, value_input_option='USER_ENTERED')
            try:
                btn = driver.find_element(By.XPATH, "//button[./div[contains(@style,'pagination-right')]]")
                if btn.get_attribute('disabled'):
                    break
                btn.click()
                time.sleep(2)
            except NoSuchElementException:
                break
        print('Полная синхронизация завершена')

    else:
        # Инкрементальная синхронизация
        print(f"Инкрементальная синхронизация от ID {key_id}")
        new_rows = []
        found = False
        page  = 0
        while not found and page < DEBUG_PAGE_LIMIT:
            page += 1
            print(f'Страница {page}')
            rows = parse_current_page()
            for r in rows:
                if r[2] == key_id:
                    found = True
                    break
                new_rows.append(r)
            if not found:
                try:
                    btn = driver.find_element(By.XPATH, "//button[./div[contains(@style,'pagination-right')]]")
                    if btn.get_attribute('disabled'):
                        break
                    btn.click()
                    time.sleep(2)
                except NoSuchElementException:
                    break
        if found:
            # Удаляем дубли и вставляем новые сверху
            ids = ws.col_values(3)
            idx = ids.index(key_id) + 1
            if idx > 2:
                ws.delete_rows(2, idx - 1)
            for row in reversed(new_rows):
                ws.insert_row(row, index=2, value_input_option='USER_ENTERED')
            print(f'Добавлено {len(new_rows)} новых строк')
        else:
            print('Ключевой ID не найден, выполняю полную синхронизацию')
            ws.clear()
            main()

if __name__ == '__main__':
    try:
        main()
    finally:
        driver.quit()
