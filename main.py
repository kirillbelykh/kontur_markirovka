import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import time
import logging


# Настройка логгирования
logging.basicConfig(
    filename="kontur_log.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# -----------------------------
# Настройка YandexDriver и браузера
# -----------------------------
yandexdriver_path = r"driver\yandexdriver.exe"
yandex_browser_path = r"C:\Users\sklad\AppData\Local\Yandex\YandexBrowser\Application\browser.exe"

options = Options()
options.binary_location = yandex_browser_path

options.add_argument(r"--user-data-dir=C:\Users\sklad\AppData\Local\Yandex\YandexBrowser\User Data\Default")
options.add_argument(r'--profile-directory=Vinsent O`neal')

service = Service(yandexdriver_path)
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)

# -----------------------------
# Ввод данных с терминала
# -----------------------------
name = input("Упрощенно: ").strip().lower()
size = input("Размер: ").strip().lower()
units_per_pack = input("Количество единиц в упаковке: ").strip()
quantity_value = input("Количество кодов: ").strip()

# -----------------------------
# Чтение Excel
# -----------------------------
df = pd.read_excel("data/nomenclature.xlsx")
print(f"Всего строк в Excel: {len(df)}")

# Фильтруем по вводным
row = df[
    (df['Упрощенно'].str.lower() == name) &
    (df['Размер'].str.lower().str.contains(size)) &
    (df['Количество единиц употребления в потребительской упаковке'].astype(str).str.contains(units_per_pack))
]

if row.empty:
    print("Ошибка: GTIN для этой номенклатуры не найден")
    driver.quit()
    exit()

gtin = row['GTIN'].values[0]
print(f"✅ Найден GTIN: {gtin}")

# -----------------------------
# Переход на страницу маркировки
# -----------------------------
driver.get("https://mk.kontur.ru/organizations/5cda50fa-523f-4bb5-85b6-66d7241b23cd/warehouses")
time.sleep(3)

# Шаг 0.2: Выбор профиля ООО 'Грундлаге'
print("Шаг 0.2: Выбираем профиль ООО 'Грундлаге'")
profile_card = wait.until(EC.element_to_be_clickable(
    (By.XPATH, '//*[@id="root"]/div/div/div[1]/div[2]/div/div/div/div/div[2]/div/div/div/div/div/div/div[1]/div/div/div/div[1]/div/div')
))
profile_card.click()
time.sleep(3)

# Шаг 0.3: Выбор склада "Лахта"
print("Шаг 0.3: Выбираем склад 'Лахта'")
warehouse_card = wait.until(EC.element_to_be_clickable(
    (By.XPATH, '//*[@id="root"]/div/div/div[2]/div/div/div[1]/div[3]/ul/li/div[2]')
))
warehouse_card.click()
time.sleep(3)

# -----------------------------
# Шаг 1: Заказать коды
# -----------------------------
print("Шаг 1: Нажимаем 'Заказать коды'")
order_codes_btn = wait.until(EC.element_to_be_clickable(
    (By.XPATH, '//*[@id="root"]/div/div/div[2]/div/div[1]/div/div/div[2]/div/div[2]/div/div/div/span[1]/span/button/div[2]/span[2]')
))
order_codes_btn.click()
time.sleep(2)

# -----------------------------
# Шаг 2: Производство РФ
# -----------------------------
print("Шаг 2: Нажимаем 'Производство РФ'")
rf_production_btn = wait.until(EC.element_to_be_clickable(
    (By.XPATH, '/html/body/div[5]/div/div[2]/div/div/div/div/div[2]/div[2]/div/div[1]/span/span/div/div[2]/div/label/div')
))
rf_production_btn.click()
time.sleep(2)

# -----------------------------
# Шаг 3: Далее
# -----------------------------
print("Шаг 3: Нажимаем 'Далее'")
next_btn = wait.until(EC.element_to_be_clickable(
    (By.XPATH, '/html/body/div[5]/div/div[2]/div/div/div/div/div[2]/div[3]/div/div/div/div[2]/div/div/span[1]/span/button/div[2]/span')
))
next_btn.click()
time.sleep(2)

# -----------------------------
# Шаг 4: Галочка "Наполнить из справочника"
# -----------------------------
print("Шаг 4: Ставим галочку 'Наполнить из справочника'")
fill_from_catalog_checkbox = wait.until(EC.element_to_be_clickable(
    (By.XPATH, '//*[@id="root"]/div/div/div[2]/div/span/div/div[2]/div/div/span/div/div[2]/div')
))
driver.execute_script("arguments[0].click();", fill_from_catalog_checkbox)
time.sleep(2)  # Увеличиваем задержку

# -----------------------------
# Шаг 5: Заполняем поле 'Заказ кодов №'
print("Шаг 5: Заполняем поле 'Заказ кодов №'")
# -----------------------------
# Находим поле
order_number_input = wait.until(EC.element_to_be_clickable(
    (By.XPATH, '//*[@id="root"]/div/div/div[2]/div/span/div/div[1]/div/div[1]/div[1]/div[1]/div/span/label/span[2]/input')
))

# Очищаем поле и вводим значение медленно, посимвольно
order_number_input.clear()
time.sleep(0.5)
for char in "ТЕСТ":
    order_number_input.send_keys(char)
    time.sleep(0.1)  # Небольшая задержка между символами
time.sleep(1)

# Проверяем, что поле заполнено
if order_number_input.get_attribute("value") != "ТЕСТ":
    print("Поле не заполнено, заполняем снова")
    order_number_input.clear()
    order_number_input.send_keys("ТЕСТ")
    time.sleep(1)

# -----------------------------
# Шаг 6: Кликаем 'Далее к заполнению реквизитов'
print("Шаг 6: Кликаем 'Далее к заполнению реквизитов'")

# Добавляем дополнительную проверку, что поле действительно заполнено
WebDriverWait(driver, 5).until(
    lambda driver: order_number_input.get_attribute("value") == "ТЕСТ"
)

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
    
except Exception as e:
    print(f"Ошибка при клике на кнопку: {e}")
    # Делаем скриншот для диагностики
    driver.save_screenshot("error_next_button.png")
    print("Сделан скриншот error_next_button.png")
    # Вместо выхода, продолжаем, если страница загрузилась вручную
    print("Пытаемся продолжить несмотря на ошибку в ожидании")

# -----------------------------
# Шаг 7: Проверяем и выбираем "Единица товара" если не выбрано
print("Шаг 7: Проверяем и выбираем 'Единица товара' если не выбрано")

try:
    # Находим лейбл выбранного значения
    label = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "[data-test-id='cisTypeField'] [data-tid='Select__label']")
    ))
    selected_text = label.text.strip()
    print(f"Текущее значение: '{selected_text}'")
    
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
    
except Exception as e:
    print(f"Ошибка при выборе 'Единица товара': {e}")
    driver.save_screenshot("error_select_unit.png")
    print("Сделан скриншот error_select_unit.png")
    print("Продолжаем несмотря на ошибку в выборе")

# Просто ждем немного чтобы страница стабилизировалась
time.sleep(2)

# -----------------------------
# Шаг 8: Кликаем 'Далее к загрузке товаров'
print("Шаг 8: Кликаем 'Далее к загрузке товаров'")

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

# -----------------------------
# -----------------------------
# Шаг 9: Вводим GTIN и количество
print("Шаг 9: Вводим GTIN и количество кодов")

try:
    # Вводим GTIN
    gtin_input = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-test-id="productCatalogSearchInput"] input'))
    )
    gtin_input.clear()
    gtin_input.send_keys(str(gtin))
    logging.info(f"Введен GTIN: {gtin}")
    time.sleep(2)  # время на появление списка

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
    for char in str(quantity_value):
        qty_input.send_keys(char)
        time.sleep(0.1)
    
    # Убеждаемся, что значение установилось
    time.sleep(0.5)
    
    # Имитируем потерю фокуса (TAB) для активации валидации
    qty_input.send_keys(Keys.TAB)
    time.sleep(1)
    
    # Проверяем, что значение установилось правильно
    current_qty = qty_input.get_attribute("value")
    if current_qty != str(quantity_value):
        logging.warning(f"⚠ Количество не совпадает: ожидалось {quantity_value}, получено {current_qty}")
        # Пробуем установить значение через JavaScript
        driver.execute_script("""
            arguments[0].value = arguments[1];
            var event = new Event('input', { bubbles: true });
            arguments[0].dispatchEvent(event);
            var changeEvent = new Event('change', { bubbles: true });
            arguments[0].dispatchEvent(changeEvent);
        """, qty_input, str(quantity_value))
        time.sleep(1)
    else:
        logging.info(f"✅ Количество кодов подтверждено: {quantity_value}")

except Exception as e:
    logging.error(f"Ошибка при вводе GTIN или количества: {e}")
    driver.save_screenshot("error_gtin_qty.png")



# -----------------------------
# Шаг 10: Отправка в ГИС МТ
logging.info("Шаг 10: Нажимаем 'Отправить в ГИС МТ'")

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

# -----------------------------
# Шаг 11: ПОДПИСАТЬ СЕРТИФИКАТОМ
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


# -----------------------------
# Шаг 12: Подписать и отправить в ГИС МТ
logging.info("Шаг 12: Нажимаем 'Подписать и отправить в ГИС МТ'")

try:
    # Ждём кнопку
    sign_send_button = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-test-id="signAndSendToGISMT"] button'))
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", sign_send_button)
    time.sleep(0.5)

    # Кликаем через JS для надежности
    driver.execute_script("arguments[0].click();", sign_send_button)
    logging.info("✅ Кнопка 'Подписать и отправить в ГИС МТ' нажата")

    # Небольшая задержка на обработку
    time.sleep(3)

except Exception as e:
    logging.error(f"Ошибка при нажатии кнопки 'Подписать и отправить в ГИС МТ': {e}")
    driver.save_screenshot("error_sign_and_send.png")

# -----------------------------
# Завершаем работу
time.sleep(10)
driver.quit()
logging.info("✅ Работа скрипта завершена")