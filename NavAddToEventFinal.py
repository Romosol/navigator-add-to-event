import os
import json
import time
import datetime
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# ==================== КЛАСС ДЛЯ ЛОГИРОВАНИЯ КОНСОЛИ ====================
class ConsoleLogger(object):
    def __init__(self, filename):
        self.terminal = sys.stdout
        self.log = open(filename, "a", encoding="utf-8")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
        self.log.flush()  # Мгновенная запись в файл, чтобы ничего не терялось при сбое

    def flush(self):
        self.terminal.flush()
        self.log.flush()

# Генерируем текущую дату и время для названий файлов
current_time_str = datetime.datetime.now().strftime("%d.%m.%Y_%H-%M-%S")
steps_filename = f"шаги_{current_time_str}.txt"
log_filename = f"логи_{current_time_str}.txt"

# Включаем перехватчик: теперь всё из print() пойдет и в консоль, и в файл "шаги_..."
sys.stdout = ConsoleLogger(steps_filename)
# =======================================================================

# ==================== РАБОТА С ФАЙЛОМ КОНФИГУРАЦИИ ====================
CONFIG_FILE = "config.json"
config_data = {}

# Проверяем, существует ли уже локальный файл конфигурации
if os.path.exists(CONFIG_FILE):
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config_data = json.load(f)
        print("\n[КОНФИГ] Найдена сохраненная конфигурация:")
        print(f"  • Логин: {config_data.get('login')}")
        print(f"  • Мероприятие: {config_data.get('event_text')}")
        print(f"  • Дата: {config_data.get('event_date')}")
        print(f"  • Время: {config_data.get('event_time')}")
        
        choice = input("\nИспользовать эти данные? [Нажмите Enter, чтобы продолжить. Введитель n, чтобы ввести данные заного.]: ").strip().lower()
        if choice in ['н', 'n', 'нет', 'no']:
            config_data = {} # Сбрасываем, чтобы запросить новые данные
    except Exception as e:
        print(f"[ВНИМАНИЕ] Не удалось прочитать файл конфигурации ({e}), создаем новый.")
        config_data = {}

# Если файла нет или пользователь решил обновить данные
if not config_data:
    print("\n=== НАСТРОЙКА ПАРАМЕТРОВ (Данные сохранятся локально) ===")
    config_data['login'] = input("1. Введите E-mail (логин): ").strip()
    config_data['password'] = input("2. Введите пароль: ").strip()
    config_data['event_text'] = input("3. Введите ПОЛНОЕ название мероприятия: ").strip()
    
    default_date = datetime.datetime.now().strftime("%d.%m.%Y")
    user_date = input(f"4. Введите дату мероприятия (например, 22.05.2026) [Enter для {default_date}]: ").strip()
    config_data['event_date'] = user_date if user_date else default_date
    
    user_time = input("5. Введите время мероприятия (например, 11:00) [Enter для 11:00]: ").strip()
    config_data['event_time'] = user_time if user_time else "11:00"
    
    # Сохраняем введенные данные в файл config.json
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False, indent=4)
    print(f"[УСПЕШНО] Настройки сохранены в файл '{CONFIG_FILE}'")

# Распределяем переменные из конфигурации
LOGIN = config_data['login']
PASSWORD = config_data['password']
EVENT_TEXT = config_data['event_text']
EVENT_DATE = config_data['event_date']
EVENT_TIME = config_data['event_time']
# ======================================================================

# 1. ЗАГРУЖАЕМ СПИСОК ДЕТЕЙ
df = pd.read_excel('list.xlsx')

# 3. ЗАПУСКАЕМ БРАУЗЕР
driver = webdriver.Chrome()
driver.get('https://админка02.навигатор.дети/admin/#activity_order') 

# ================= АВТОМАТИЧЕСКАЯ АВТОРИЗАЦИЯ =================
print("\n[ИНФО] Начинаем автоматическую авторизацию...")
try:
    login_input = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="textfield-1598-inputEl"] | //input[@placeholder="E-mail"]'))
    )
    login_input.clear()
    login_input.send_keys(LOGIN)
    print("[УСПЕШНО] Шаг А1: Поле ввода логина заполнено")

    password_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="textfield-1599-inputEl"] | //input[@placeholder="Пароль"]'))
    )
    password_input.clear()
    password_input.send_keys(PASSWORD)
    print("[УСПЕШНО] Шаг А2: Поле ввода пароля заполнено")

    login_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="button-1605"] | //button[contains(., "ВОЙТИ")] | //*[text()="ВОЙТИ"]'))
    )
    login_btn.click()
    print("[УСПЕШНО] Шаг А3: Нажата кнопка 'ВОЙТИ'")
    
    print("\n[ИНФО] Ожидаем загрузки основного рабочего кабинета...")
    time.sleep(5) 
    
    print("--> Пожалуйста, убедитесь, что вы находитесь в разделе 'Заявки на мероприятия'.")
    print("--> Пожалуйста, разместите окно браузера и консоли, чтобы вам были видны все данные.")
    print("--> НЕ МЕНЯЙТЕ РАЗМЕР ОКНА БРАЗУЕРА!!! Иначе программа может не увидеть некоторые элементы.")
    input("--> Как только будете готовы запустить обработку списка, нажмите ENTER здесь, в консоли...")

except Exception as auth_err:
    print(f"\n[КРИТИЧЕСКАЯ ОШИБКА АВТОРИЗАЦИИ] Не удалось войти автоматически: {auth_err}")
    print("[ИНФО] Пожалуйста, введите данные на сайте вручную.")
    input("Как зайдете в нужный раздел заявок, нажмите ENTER здесь, в консоли для продолжения...")
# ==============================================================


# 4. ЦИКЛ АВТОМАТИЗАЦИИ СПИСКА ДЕТЕЙ
for index, row in df.iterrows():
    
    # === ПРОВЕРКА НА ПУСТОЕ ЗНАЧЕНИЕ ФИО ===
    first_cell = row.iloc[0]
    if pd.isna(first_cell) or str(first_cell).strip() == "" or str(first_cell).strip().lower() == "nan":
        log_status = "в списке пустое значение"
        print(f"\n==================================================")
        print(f" Строка №{index + 1}: ПРОПУЩЕНА (Обнаружено пустое значение)")
        print(f"==================================================")
        
        # Записываем в лог информацию о пустой строке
        with open(log_filename, "a", encoding="utf-8") as log_file:
            log_file.write(f"Строка №{index + 1} - {log_status}\n")
        continue  # Сразу переходим к следующей строке Excel
    # =======================================

    fio_child = str(first_cell).strip()
    log_status = "пропущен" # Статус по умолчанию для текущей итерации
    
    # Считываем дату рождения из соседнего (второго) столбца
    birth_date_val = row.iloc[1]
    if pd.isna(birth_date_val):
        birth_date = ""
    else:
        try:
            birth_date = pd.to_datetime(birth_date_val, dayfirst=True).strftime("%d.%m.%Y")
        except:
            birth_date = str(birth_date_val).strip()
    
    print(f"\n==================================================")
    print(f" Начинаем обработку ребенка №{index + 1}: {fio_child}")
    if birth_date:
        print(f" Дата рождения по Excel: {birth_date}")
    print(f"==================================================")
    
    try:
        # Шаг 1: Нажимаем кнопку "Создать"
        create_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@data-qtip="Создать"]'))
        )
        create_btn.click()
        print("[УСПЕШНО] Шаг 1.1: Нажата кнопка 'Создать'")
        
        # Ожидаем появления формы создания заявки
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@id, "activity_order-create-form")]'))
        )
        time.sleep(1) 
        print("[УСПЕШНО] Шаг 1.2: Форма создания успешно открылась")

        # Шаг 2: Заполнение полей СТРОГО внутри формы создания
        
        # 2.1. Мероприятие
        event_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@id, "activity_order-create-form")]//input[contains(@id, "activitycombo")]'))
        )
        
        driver.execute_script("arguments[0].value = arguments[1];", event_input, EVENT_TEXT)
        driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", event_input)
        time.sleep(0.5)
        
        driver.execute_script("arguments[0].focus(); arguments[0].click();", event_input)
        actions = ActionChains(driver)
        actions.send_keys(Keys.SPACE).send_keys(Keys.BACKSPACE).perform()
        
        time.sleep(1.5)  
        actions.send_keys(Keys.ENTER).perform()
        print("[УСПЕШНО] Шаг 2.1: Название мероприятия зафиксировано")

        # 2.2. Дата
        date_input = driver.find_element(By.XPATH, '//div[contains(@id, "activity_order-create-form")]//input[contains(@id, "datefield")]')
        driver.execute_script("arguments[0].focus(); arguments[0].click();", date_input)
        time.sleep(0.3)
        
        actions = ActionChains(driver)
        actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
        actions.send_keys(Keys.DELETE).perform()
        actions.send_keys(EVENT_DATE).perform()
        time.sleep(0.3)
        actions.send_keys(Keys.ENTER).perform()
        print(f"[УСПЕШНО] Шаг 2.2: Дата ({EVENT_DATE}) заполнена")

        # 2.3. Время
        time_input = driver.find_element(By.XPATH, '//div[contains(@id, "activity_order-create-form")]//input[contains(@id, "timefield")]')
        driver.execute_script("arguments[0].focus(); arguments[0].click();", time_input)
        time.sleep(0.3)
        
        actions = ActionChains(driver)
        actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
        actions.send_keys(Keys.DELETE).perform()
        actions.send_keys(EVENT_TIME).perform()
        actions.send_keys(Keys.ENTER).perform()
        print(f"[УСПЕШНО] Шаг 2.3: Время '{EVENT_TIME}' заполнено")

        # Шаг 3: Поиск ФИО ребенка
        search_input = driver.find_element(By.XPATH, '//div[contains(@id, "activity_order-create-form")]//input[contains(@id, "kidcombo")]')
        
        driver.execute_script("arguments[0].value = arguments[1];", search_input, fio_child)
        driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", search_input)
        time.sleep(0.5)
        
        driver.execute_script("arguments[0].focus(); arguments[0].click();", search_input)
        actions = ActionChains(driver)
        actions.send_keys(Keys.SPACE).send_keys(Keys.BACKSPACE).perform()
        print(f"[ИНФО] Шаг 3.1: ФИО '{fio_child}' отправлено в поиск. Анализируем выпадающий список...")
        
        # Даем сайту время подгрузить результаты фильтрации
        time.sleep(1.5)

        # ================= ШАГ 4: ПОИСК С УЧЕТОМ ПОЛНЫХ ТЁЗОК (ФИО + ДАТА РОЖДЕНИЯ) =================
        try:
            picker_items_xpath = "//*[contains(@id, 'kidcombo') and contains(@id, 'picker')]//li[contains(@class, 'x-boundlist-item')]"
            items = driver.find_elements(By.XPATH, picker_items_xpath)
            
            target_element = None
            clean_excel_fio = fio_child.lower().replace('\xa0', ' ').strip()
            clean_excel_dob = birth_date.strip()
            
            # Подсчитываем количество совпадений по ФИО в списке для определения тёзок
            fio_matches_count = 0
            for item in items:
                item_text = item.text.lower().replace('\xa0', ' ').strip()
                if clean_excel_fio in item_text or item_text in clean_excel_fio:
                    fio_matches_count += 1
            
            for item in items:
                item_text = item.text.lower().replace('\xa0', ' ').strip()
                
                # Критерий 1: В строке должно содержаться ФИО
                fio_matches = (clean_excel_fio in item_text or item_text in clean_excel_fio)
                
                # Критерий 2: Если дата рождения есть в Excel, она ДОЛЖНА быть в этой строке сайта
                dob_matches = True
                if clean_excel_dob:
                    dob_matches = (clean_excel_dob in item.text)
                
                # Если оба условия выполнены — мы нашли нужного ребенка (даже среди тёзок)
                if fio_matches and dob_matches:
                    target_element = item
                    break
            
            if target_element:
                driver.execute_script("arguments[0].scrollIntoView(false);", target_element)
                time.sleep(0.2)
                try:
                    target_element.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", target_element)
                
                # Присваиваем статус в зависимости от наличия тёзок в списке
                if fio_matches_count > 1:
                    log_status = "успешно добавлен среди тёзок"
                else:
                    log_status = "успешно добавлен"
                    
                print(f"[УСПЕШНО] Шаг 4: Ребенок '{fio_child}' ({birth_date}) успешно выбран автоматически.")
                time.sleep(0.5)
            else:
                found_texts = [i.text.strip() for i in items if i.text.strip()]
                print(f"[ИНФО] Робот увидел на экране варианты: {found_texts}")
                raise Exception("Не найдено точное совпадение по связке ФИО + Дата рождения")
                
        except Exception:
            # Блок ручного выбора или пропуска, если автоматика не смогла сопоставить пару ФИО+Дата
            print(f"\n[ВНИМАНИЕ !!!] Не удалось автоматически выбрать ребенка '{fio_child}' с датой {birth_date}!")
            print("--> Пожалуйста, выберите действие:")
            print("    1. Выбрать ВРУЧНУЮ: Кликните по нужному ребенку (тёзке) на сайте сами и нажмите ENTER в консоли.")
            print("    2. ПРОПУСТИТЬ: Введите букву 'н' (или 'n') в консоли и нажмите ENTER, чтобы перейти к следующему.")
            
            user_choice = input("--> Ваш выбор: ").strip().lower()
            
            if user_choice in ['н', 'n', 'т', 'y']:
                print(f"[ИНФО] Пропускаем ребенка {fio_child} по команде пользователя.")
                log_status = "пропущен"
                # Записываем в лог перед переходом к следующей итерации
                with open(log_filename, "a", encoding="utf-8") as log_file:
                    log_file.write(f"{fio_child} - {log_status}\n")
                try:
                    cancel_btn = driver.find_element(By.XPATH, "//div[contains(@id, 'activity_order-create-form')]//*[text()='Закрыть']")
                    cancel_btn.click()
                    print("[ИНФО] Текущая форма создания успешно закрыта.")
                    time.sleep(1)
                except Exception:
                    pass
                continue  # Переходим к следующему ребенку в Excel
            else:
                log_status = "выбран вручную"
                
                # === ЗАЩИТА: ПРОВЕРКА АКТИВНОГО ОКНА ФОРМЫ ===
                time.sleep(0.5) 
                active_form = driver.find_elements(By.XPATH, '//div[contains(@id, "activity_order-create-form")]')
                
                if not active_form:
                    print(f"[УСПЕШНО] Обнаружено, что форма уже была сохранена и закрыта вами вручную на сайте.")
                    with open(log_filename, "a", encoding="utf-8") as log_file:
                        log_file.write(f"{fio_child} - {log_status}\n")
                    time.sleep(1.0)
                    continue 
        # ====================================================================================

        # Шаг 5: Нажимаем кнопку "Сохранить" (выполняется, только если форма все еще открыта)
        save_btn = driver.find_element(By.XPATH, "//div[contains(@id, 'activity_order-create-form')]//*[text()='Сохранить']")
        save_btn.click()
        print("[УСПЕШНО] Шаг 5: Кнопка 'Сохранить' нажата")

        print(f"\n[ГОТОВО] Ребенок {fio_child} успешно обработан.")
        
        # Записываем успешный статус работы в файл логов
        with open(log_filename, "a", encoding="utf-8") as log_file:
            log_file.write(f"{fio_child} - {log_status}\n")
            
        time.sleep(1.5) 

    except Exception as e:
        print(f"\n[КРИТИЧЕСКАЯ ОШИБКА] На каком-то этапе произошел сбой!")
        print(f"Причина ошибки: {e}")
        print(f"Пропускаем ребенка {fio_child} и закрываем текущее окно.")
        
        log_status = "пропущен"
        with open(log_filename, "a", encoding="utf-8") as log_file:
            log_file.write(f"{fio_child} - {log_status}\n")
        
        try:
            cancel_btn = driver.find_element(By.XPATH, "//div[contains(@id, 'activity_order-create-form')]//*[text()='Закрыть']")
            cancel_btn.click()
            print("[ИНФО] Окно формы успешно принудительно закрыто после ошибки")
            time.sleep(1)
        except Exception as cancel_error:
            print(f"[ВНИМАНИЕ] Не удалось закрыть окно кнопкой 'Закрыть': {cancel_error}")
            pass

print(f"\nУра! Весь список успешно обработан.")
print(f"Результаты проверки детей сохранены в: {log_filename}")
print(f"Полный технический протокол консоли сохранен в: {steps_filename}")
driver.quit()