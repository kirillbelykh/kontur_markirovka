# main.py
import os
import time
import logging
import pandas as pd
from concurrent.futures import ProcessPoolExecutor, as_completed
from multiprocessing import cpu_count
from dataclasses import dataclass, asdict
from typing import List, Dict

# selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException

# -----------------------------
# ========== CONFIG ===========
# -----------------------------
YANDEX_DRIVER_PATH = r"driver\yandexdriver.exe"  # проверь путь
YANDEX_BROWSER_PATH = r"C:\Users\sklad\AppData\Local\Yandex\YandexBrowser\Application\browser.exe"
NOMENCLATURE_XLSX = "data/nomenclature.xlsx"
LOG_FILE = "kontur_log.log"

# Запускать в фоне (headless). Если нужен профиль/авторизация - ставь False.
HEADLESS = False

# Кол-во параллельных процессов (по умолчанию cpu_count())
MAX_WORKERS = max(1, cpu_count() - 1)

# Тайминги (настрой, если нужно)
SHORT_SLEEP = 0.2
MEDIUM_SLEEP = 1.0
LONG_SLEEP = 2.5

# -----------------------------
# logging (минимальные сообщения в терминал, подробности в файл)
# -----------------------------
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# helper prints only for prompts / summary
def ui_print(msg: str):
    print(msg)

# -----------------------------
# Data container
# -----------------------------
@dataclass
class OrderItem:
    order_name: str         # Заявка № или текст для "Заказ кодов №"
    simpl_name: str         # Упрощенно
    size: str               # Размер
    units_per_pack: str     # Количество единиц в упаковке (строка, для поиска)
    codes_count: int        # Количество кодов для заказа
    gtin: str = ""          # найдём перед запуском воркеров
    full_name: str = ""     # опционально: полное наименование из справочника

# -----------------------------
# Lookup GTIN in nomenclature.xlsx
# -----------------------------
def lookup_gtin(df: pd.DataFrame, simpl_name: str, size: str, units_per_pack: str,
                color: str = None, venchik: str = None):
    """
    Поиск GTIN и полного наименования по заданным полям.
    Для венчика используется точное совпадение.
    """
    try:
        simpl = simpl_name.strip().lower()
        size_l = str(size).strip().lower()
        units_str = str(units_per_pack).strip()
        color_l = color.strip().lower() if color else None
        venchik_l = venchik.strip().lower() if venchik else None

        # Создаём недостающие колонки, чтобы не ломалось
        for col in ['GTIN', 'Наименование', 'Упрощенно', 'Размер',
                    'Количество единиц употребления в потребительской упаковке', 'Цвет', 'венчик']:
            if col not in df.columns:
                df[col] = ""

        # Основной фильтр
        cond = (
            df['Упрощенно'].astype(str).str.strip().str.lower() == simpl
        ) & (
            df['Размер'].astype(str).str.strip().str.lower().str.contains(size_l, na=False)
        ) & (
            df['Количество единиц употребления в потребительской упаковке'].astype(str).str.strip() == units_str
        )

        if venchik_l:
            cond &= df['венчик'].astype(str).str.strip().str.lower() == venchik_l
        if color_l:
            cond &= df['Цвет'].astype(str).str.strip().str.lower() == color_l

        matches = df[cond]

        if not matches.empty:
            row = matches.iloc[0]
            return str(row['GTIN']).strip(), str(row.get('Наименование', '')).strip()

        # Частичный поиск по Упрощенно и размеру
        cond2 = (
            df['Упрощенно'].astype(str).str.strip().str.lower().str.contains(simpl, na=False)
        ) & (
            df['Размер'].astype(str).str.strip().str.lower().str.contains(size_l, na=False)
        )
        if venchik_l:
            cond2 &= df['Венчик'].astype(str).str.strip().str.lower() == venchik_l
        if color_l:
            cond2 &= df['Цвет'].astype(str).str.strip().str.lower() == color_l

        matches2 = df[cond2]
        if not matches2.empty:
            row = matches2.iloc[0]
            return str(row['GTIN']).strip(), str(row.get('Наименование', '')).strip()

    except Exception as e:
        logging.exception("Ошибка в lookup_gtin")
    return None, None


# -----------------------------
# Worker: выполняет заказ для одной позиции
# -----------------------------
browser_not_found = []
not_found_list = []
def perform_order_item(item: Dict):
    """
    Запускается в отдельном процессе. Получает словарь item (OrderItem -> asdict).
    Делает браузерную автоматизацию для создания заявки.
    Возвращает (True/False, message)
    """
    # В процессе логируем в файл
    logging.info(f"Worker start for order: {item.get('order_name')} - {item.get('simpl_name')}")

    # Преобразуем принимаемые данные
    order_name = item['order_name']
    gtin = item['gtin']
    codes_count = item['codes_count']
    # Additional fields if needed
    simpl_name = item['simpl_name']
    size = item['size']
    units_per_pack = item['units_per_pack']

    # Selenium setup (each process создает свой драйвер)
    try:
        options = Options()
        options.binary_location = YANDEX_BROWSER_PATH

        # headless (background) — используй современный режим, если поддерживается
        if HEADLESS:
            # New headless mode
            options.add_argument("--headless=new")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument(r"--user-data-dir=C:\Users\sklad\AppData\Local\Yandex\YandexBrowser\User Data\Default")
            options.add_argument(r'--profile-directory=Vinsent O`neal')
            options.add_argument("--disable-blink-features=AutomationControlled")
        else:
            # If you need to use profile, configure here (be careful with concurrent profile usage)
            # options.add_argument(r"--user-data-dir=...")  # uncomment if necessary
            options.add_argument(r"--user-data-dir=C:\Users\sklad\AppData\Local\Yandex\YandexBrowser\User Data\Default")
            options.add_argument(r'--profile-directory=Vinsent O`neal')

        # Common options
        options.add_argument("--disable-features=VizDisplayCompositor")
        options.add_argument("--disable-popup-blocking")
        # prevent Selenium from stealing focus (but some behaviors on Windows still bring window forward)
        options.add_argument("--disable-backgrounding-occluded-windows")

        service = Service(YANDEX_DRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 20)

        # --- Begin navigation & form filling ---
        # NOTE: we rely on the XPATHs/selectors you provided earlier. Adjust if the page changes.
        driver.get("https://mk.kontur.ru/organizations/5cda50fa-523f-4bb5-85b6-66d7241b23cd/warehouses")
        time.sleep(3)

        # profile select (if visible) - best-effort, ignore if not
        try:
            profile_card = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="root"]/div/div/div[1]/div[2]/div/div/div/div/div[2]/div/div/div/div/div/div/div[1]/div/div/div/div[1]/div/div')
            ))
            profile_card.click()
            print("Выбрали профиль")
            time.sleep(1)
        except Exception:
            logging.info("Profile card not found/clickable or already selected")

        # warehouse select (best-effort)
        try:
            warehouse_card = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="root"]/div/div/div[2]/div/div/div[1]/div[3]/ul/li/div[2]')
            ))
            warehouse_card.click()
            print("Тыкнули профиль")
            time.sleep(1)
        except Exception:
            logging.info("Warehouse card not found/clickable or already selected")

        # Step 1: 'Заказать коды'
        try:
            order_codes_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="root"]/div/div/div[2]/div/div[1]/div/div/div[2]/div/div[2]/div/div/div/span[1]/span/button/div[2]/span[2]')
            ))
            order_codes_btn.click()
            print("Нажали Заказать Коды")
            time.sleep(2)
        except Exception:
            logging.error("Не удалось найти/кликнуть 'Заказать коды'")
            # try continue — some pages might already be in ordering flow
            pass

        # Step 2: 'Производство РФ'
        try:
            rf_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[5]/div/div[2]/div/div/div/div/div[2]/div[2]/div/div[1]/span/span/div/div[2]/div/label/div')
            ))
            rf_btn.click()
            print("Нажали производство РФ")
            time.sleep(2)
        except Exception:
            logging.info("Производство РФ выбор: не найден/не понадобился")

        # Step 3: Далее
        try:
            next_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[5]/div/div[2]/div/div/div/div/div[2]/div[3]/div/div/div/div[2]/div/div/span[1]/span/button/div[2]/span')
            ))
            next_btn.click()
            time.sleep(1)
        except Exception:
            logging.info("Кнопка 'Далее' не обнаружена — продолжаем")

        # Step: "Наполнить из справочника"
        try:
            fill_from_catalog_checkbox = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="root"]/div/div/div[2]/div/span/div/div[2]/div/div/span/div/div[2]/div')
            ))
            driver.execute_script("arguments[0].click();", fill_from_catalog_checkbox)
            time.sleep(1)
        except Exception:
            logging.info("Галочка 'Наполнить из справочника' не найдена/не нужна")

        # Step: Fill "Заказ кодов №" — insert order_name from input
        try:
            order_number_input = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="root"]/div/div/div[2]/div/span/div/div[1]/div/div[1]/div[1]/div[1]/div/span/label/span[2]/input')
            ))
            order_number_input.clear()
            # slow typed input to trigger React listeners
            order_number_input.send_keys(str(order_name))
            time.sleep(1)
            
        except Exception:
            logging.warning("Поле 'Заказ кодов №' не найдено/не удалось заполнить")


        # Step: Далее к заполнению реквизитов
        try:
            # Ждем и кликаем кнопку
            next_req_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Далее к заполнению реквизитов')]]"))
            )
            
            # Прокручиваем и кликаем через JS
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_req_button)
            time.sleep(1)  # Увеличиваем задержку
            
            # Проверяем, что кнопка видима и кликабельна
            if next_req_button.is_displayed() and next_req_button.is_enabled():
                driver.execute_script("arguments[0].click();", next_req_button)
                print("Кнопка нажата через JS")
            else:
                print("Кнопка не кликабельна, пробуем другой метод")
                ActionChains(driver).move_to_element(next_req_button).click().perform()
            
            # Ждем загрузки следующей страницы - проверяем появление селектора cisTypeField
            WebDriverWait(driver, 20).until(  # Увеличили время ожидания
                EC.presence_of_element_located((By.CSS_SELECTOR, "[data-test-id='cisTypeField']"))
            )
            time.sleep(2)
            print("Успешно перешли к заполнению реквизитов")
        except Exception:
            logging.info("Кнопка 'Далее к заполнению реквизитов' не найдена/пропускаем")

        # Step: Ensure 'Единица товара' selected - try selecting if not
        try:
            # Находим лейбл выбранного значения
            label = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "[data-test-id='cisTypeField'] [data-tid='Select__label']")
            ))
            selected_text = label.text.strip()
            if selected_text == "Единица товара":
                print("✅ Уже выбрано 'Единица товара', пропускаем выбор")
            else:
                print("Выбираем 'Единица товара' явно")
        
                # Находим кнопку селекта
                select_button = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "[data-test-id='cisTypeField'] button[data-tid='Button__root']")
                ))
                
                # Получаем ID меню
                menu_id = select_button.get_attribute("aria-controls")
                print(f"ID меню: {menu_id}")
                
                # Кликаем чтобы открыть дропдаун
                driver.execute_script("arguments[0].click();", select_button)
                
                # Ждем появления меню
                menu = wait.until(EC.visibility_of_element_located((By.ID, menu_id)))
                
                # Находим опцию по тексту
                option_xpath = f"//*[@id='{menu_id}']//*[normalize-space(text())='Единица товара']"
                option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
                
                # Кликаем на опцию
                driver.execute_script("arguments[0].click();", option)
                
                # Ждем закрытия меню
                wait.until(EC.invisibility_of_element_located((By.ID, menu_id)))
                
                # Подтверждаем выбор
                selected_text = label.text.strip()
                if selected_text == "Единица товара":
                    print("✅ 'Единица товара' выбрано успешно")
                else:
                    raise ValueError(f"Не удалось выбрать, текущее значение: {selected_text}")
        except Exception:
            logging.info("Не удалось установить 'Единица товара' (возможно уже выбрано)")

        # Step: Далее к загрузке товаров
        try:
            # Ждем кнопку
            next_upload_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Далее к загрузке товаров')]]"))
            )

            # Прокручиваем к кнопке и кликаем через JS
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_upload_btn)
            time.sleep(0.5)
            next_upload_btn.click()
            logging.info("Попытка клика через обычный click() выполнена")

            # Проверка, что кнопка пропала (следствие перехода на следующую страницу)
            try:
                WebDriverWait(driver, 5).until(EC.staleness_of(next_upload_btn))
                logging.info("✅ Кнопка 'Далее к загрузке товаров' успешно нажата")
            except TimeoutException:
                logging.warning("Кнопка не исчезла после клика, пробуем JS-клик")
                driver.execute_script("arguments[0].click();", next_upload_btn)
                time.sleep(1)
                logging.info("Попытка клика через JS выполнена")

        except Exception as e:
            logging.error(f"Ошибка при клике на кнопку 'Далее к загрузке товаров': {e}")
            driver.save_screenshot("error_next_upload.png")
            logging.info("Сделан скриншот error_next_upload.png")

        # Step: Ввод GTIN и количество (работаем строго с выпадающим элементом, ожидаем option, кликаем по тому, что содержит GTIN)
        try:
            # Вводим GTIN
            gtin_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-test-id="productCatalogSearchInput"] input'))
            )
            gtin_input.clear()
            gtin_input.send_keys(str(gtin))
            logging.info(f"Введен GTIN: {gtin}")
            time.sleep(2)  # время на появление списка

            # После ввода GTIN ждём появления выпадающего списка
            options = driver.find_elements(By.CSS_SELECTOR, '[data-test-id="productCatalogSearchInput"] ul li')
            if not options:
                logging.warning(f"❌ GTIN {gtin} не найден в справочнике браузера")
                browser_not_found.append(gtin)
                return

            
            # Жмем стрелку вниз + Enter (выбираем первый вариант)
            gtin_input.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.3)
            gtin_input.send_keys(Keys.ENTER)
            logging.info("✅ GTIN выбран через клавиатуру (↓ + Enter)")
            time.sleep(2)

            # После выбора GTIN DOM может обновиться, поэтому нужно заново найти элементы
            # Ввод количества кодов - находим поле заново после обновления DOM
            qty_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-test-id="codesQuantityInput"] input'))
            )
            
            # Прокручиваем к полю и кликаем на него
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", qty_input)
            time.sleep(0.5)
            
            # Очищаем поле (несколько способов)
            qty_input.click()
            qty_input.send_keys(Keys.CONTROL + "a")  # Выделяем весь текст
            qty_input.send_keys(Keys.DELETE)         # Удаляем выделенный текст
            time.sleep(0.5)
            
            # Вводим значение медленно, посимвольно
            for char in str(codes_count):
                qty_input.send_keys(char)
                time.sleep(0.1)
            
            # Убеждаемся, что значение установилось
            time.sleep(0.5)
            
            # Имитируем потерю фокуса (TAB) для активации валидации
            qty_input.send_keys(Keys.TAB)
            time.sleep(1)
            
            # Проверяем, что значение установилось правильно
            current_qty = qty_input.get_attribute("value")
            if current_qty != str(codes_count):
                logging.warning(f"⚠ Количество не совпадает: ожидалось {codes_count}, получено {current_qty}")
                # Пробуем установить значение через JavaScript
                driver.execute_script("""
                    arguments[0].value = arguments[1];
                    var event = new Event('input', { bubbles: true });
                    arguments[0].dispatchEvent(event);
                    var changeEvent = new Event('change', { bubbles: true });
                    arguments[0].dispatchEvent(changeEvent);
                """, qty_input, str(codes_count))
                time.sleep(1)
            else:
                logging.info(f"✅ Количество кодов подтверждено: {codes_count}")

        except Exception as e:
            logging.error(f"Ошибка при вводе GTIN или количества: {e}")
            driver.save_screenshot("error_gtin_qty.png")

        # Step: Нажать "Отправить в ГИС МТ"
        try:
            send_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-test-id="codesOrderSendToGISMT"] button'))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", send_button)
            time.sleep(0.5)
            send_button.click()

            # Проверка, что кнопка действительно нажата
            time.sleep(2)
            if send_button.is_enabled():
                logging.info("✅ Кнопка 'Отправить в ГИС МТ' нажата")
            else:
                logging.warning("⚠️ Кнопка 'Отправить в ГИС МТ' возможно не сработала")

        except Exception as e:
            logging.error(f"Ошибка при нажатии кнопки 'Отправить в ГИС МТ': {e}")
            driver.save_screenshot("error_send_to_gis.png")

        # Step: Подписать сертификатом
        logging.info("Нажимаем ПОДПИСАТЬ СЕРТИФИКАТОМ")
        try:
            # Ждём кнопку "Подписать сертификатом"
            sign_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-test-id="codesOrderSignCert"] button'))
            )
            logging.info("Кнопка 'Подписать сертификатом' найдена")

            # Кликаем через JS (надёжнее для React)
            driver.execute_script("arguments[0].click();", sign_button)
            logging.info("✅ Нажата кнопка 'Подписать сертификатом'")
            time.sleep(2)  # даём время на обработку

        except Exception as e:
            logging.error(f"Ошибка при нажатии кнопки 'Подписать сертификатом': {e}")
            driver.save_screenshot("error_sign_cert.png")

        # Step: Подписать и отправить в ГИС МТ
            
        try:
            #Ждём кнопку
            sign_send_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-test-id="signAndSendToGISMT"] button'))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", sign_send_button)
            time.sleep(0.5)

            #Кликаем через JS для надежности
            driver.execute_script("arguments[0].click();", sign_send_button)
            logging.info("✅ Кнопка 'Подписать и отправить в ГИС МТ' нажата")

            #Небольшая задержка на обработку
            time.sleep(3)

        except Exception as e:
            logging.error(f"Ошибка при нажатии кнопки 'Подписать и отправить в ГИС МТ': {e}")
            driver.save_screenshot("error_sign_and_send.png")
        
        # success
        driver.quit()
        return True, f"OK: {simpl_name} ({order_name})"

    except Exception as exc:
        logging.exception("Unhandled exception in worker")
        try:
            driver.quit()
        except Exception:
            pass
        return False, str(exc)

# -----------------------------
# Main interactive collection + execution
# -----------------------------
def main():
    # Load nomenclature
    if not os.path.exists(NOMENCLATURE_XLSX):
        ui_print(f"ERROR: файл {NOMENCLATURE_XLSX} не найден.")
        return

    df = pd.read_excel(NOMENCLATURE_XLSX)
    ui_print("=== Kontur Automation — ввод позиций ===")
    collected: List[OrderItem] = []

    while True:
        ui_print("\nВведите данные по позиции:")
        order_name = input("Заявка (текст, будет вставлен в 'Заказ кодов №'): ").strip()
        simpl = input("Упрощенно: ").strip().lower()
        size = input("Размер: ").strip().lower()
        units_per_pack = input("Количество единиц в упаковке: ").strip()
        codes_count_str = input("Количество кодов (целое): ").strip()

        try:
            codes_count = int(codes_count_str)
        except:
            ui_print("Неверно введено количество кодов. Попробуй ещё раз.")
            continue

        # Найдём GTIN заранее
        gtin, full_name = lookup_gtin(df, simpl, size, units_per_pack)
        if not gtin:
            ui_print(f"GTIN не найден для ({simpl}, {size}, {units_per_pack}) — позиция не добавлена. Проверь nomenclature.xlsx.")
        else:
            it = OrderItem(
                order_name=order_name,
                simpl_name=simpl,
                size=size,
                units_per_pack=units_per_pack,
                codes_count=codes_count,
                gtin=gtin,
                full_name=full_name or ""
            )
            collected.append(it)
            ui_print(f"Добавлено: {simpl} ({size}) — GTIN {gtin} — {codes_count} кодов — заявка '{order_name}'")

        # Вариант продолжения
        ui_print("\n1 - Ввести ещё позицию\n2 - Выполнить все накопленные позиции")
        choice = input("Выбор (1/2): ").strip()
        if choice == "1":
            continue
        elif choice == "2":
            break
        else:
            ui_print("Неверный выбор — продолжаем ввод.")
            continue

    if not collected:
        ui_print("Нет накопленных позиций — выходим.")
        return

    ui_print(f"\nБудет выполнено {len(collected)} задач(и) ПОСЛЕДОВАТЕЛЬНО.")
    ui_print("Запуск...")

    results = []
    for it in collected:
        try:
            ok, msg = perform_order_item(asdict(it))
            results.append((ok, msg, it))
            ui_print(f"[{'OK' if ok else 'ERR'}] {it.simpl_name} — {msg}")
        except Exception as e:
            logging.exception("Ошибка при выполнении задачи")
            results.append((False, str(e), it))
            ui_print(f"[ERR] {it.simpl_name} — exception: {e}")

    ui_print("\n=== Выполнение завершено ===")
    success = sum(1 for r in results if r[0])
    ui_print(f"Успешно: {success}, Ошибок: {len(results)-success}. Подробности в {LOG_FILE}.")

    if not_found_list or browser_not_found:
        print("\n=== Итоговый отчёт ===")

        if not_found_list:
            print("\n❌ Не найдены в Excel:")
            for item in not_found_list:
                print(" - ", item)

        if browser_not_found:
            print("\n❌ Не найдены в браузере:")
            for gtin in browser_not_found:
                print(" - ", gtin)

        print("\n✅ Остальные позиции успешно обработаны")


if __name__ == "__main__":
    main()