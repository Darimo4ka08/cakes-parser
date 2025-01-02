import time
import re
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment
from selenium import webdriver
from openpyxl import Workbook

# Настройка Selenium
def get_page_source_with_selenium(url):
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # Безголовый режим
    options.add_argument("--disable-blink-features=AutomationControlled")  # Убирает метки автоматизации
    options.add_argument(
        "user-agent=Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Mobile Safari/537.36"
    )
    driver = webdriver.Chrome(options=options)
    driver.get(url)

    # Задержка для загрузки страницы
    time.sleep(3)  # Увеличьте время, если нужно

    # Сохранение HTML-кода страницы
    page_source = driver.page_source
    driver.quit()
    return page_source

# Шаг 1. Получение списка ссылок на товары
def get_product_links(page_source):
    soup = BeautifulSoup(page_source, "lxml")

    # Найти карточки товаров и извлечь ссылки
    product_links = []
    cards = soup.find_all("article", class_="tab-content-products-item")
    for card in cards:
        link_tag = card.find("a", href=True)  # Найти ссылку
        if link_tag:
            product_links.append(link_tag["href"])  # Добавить ссылку в список

    return product_links

# Шаг 2. Извлечение данных из каждой страницы товара
def get_product_details(product_url):
    # Используем Selenium для загрузки страницы
    page_source = get_page_source_with_selenium(product_url)
    soup = BeautifulSoup(page_source, "lxml")

    # Добавить задержку
    time.sleep(3)

    # Извлечение данных:
    # Извлечение категории товара:
    category_tag = soup.find("a", class_="breadcrumb-item-title")  # Поиск хлебных крошек
    category = category_tag.get_text(strip=True) if category_tag else "Не указана"

    # Извлечение названия товара:
    name = soup.find("h1").text.strip() if soup.find("h1") else None
    # Извлечение цены товара:
    price = (
        re.sub(r'[^\d]', '', soup.find("span", class_="footer-price").text).strip()
        if soup.find("span", class_="footer-price")
        else None
    )
    # Извлечение ссылки на первое фото:
    image_tag = soup.find("img", class_="thumb")  # Найти изображение
    image_link = image_tag["src"] if image_tag else None

    # Извлечение веса товара:
    weight_cakes = soup.find("p", class_="pre-line property-text").text.strip() if soup.find("p", class_="pre-line property-text") else None

    # Извлечение категории "Добавили в подборки":
    info_text = soup.find_all("div", class_="info-text")
    # Проверить, найден ли второй элемент
    added_to_collections = re.sub(r"[^\d]", '', info_text[2].text.strip() if len(info_text) > 2 else None)

    # Извлечение названия магазина:
    shop = soup.find("p", class_="shop-name").text.strip() if soup.find("p", class_="shop-name") else None

    # Извлечение ссылки на магазин:
    shop_tag = soup.find("div", class_="shop-info-content")
    shop_url = shop_tag.find("a", href=True)["href"] if shop_tag and shop_tag.find("a", href=True) else None

    # Извлечение оценки магазина:
    shop_stats_item = soup.find("div", class_="shop-stats-item")  # Поиск тега <div> с классом shop-stats-item
    shop_rating_tag = shop_stats_item.find("p", class_="shop-rating") if shop_stats_item else None  # Поиск <p> внутри найденного <div>
    shop_rating = (
        shop_rating_tag.find("span").text.strip()
        if shop_rating_tag and shop_rating_tag.find("span")
        else "Нет данных"

)
    # Извлечение количества оценок и покупок:
    rating = None
    purchases = None

    shop_score_text = (
        soup.find("p", class_="shop-score").text.strip()
        if soup.find("p", class_="shop-score")
        else None
    )

    if shop_score_text:
        # Ищем все числа в тексте
        numbers = re.findall(r"\d+\.?\d*", shop_score_text)

        if len(numbers) == 2:
            rating = int(numbers[0])  # Первое число - количество оценок
            purchases = int(numbers[1])  # Второе число - количество покупок
        else:
            print("Не удалось извлечь рейтинг и количество покупок.")
    else:
        print("Информация о рейтинге и покупках отсутствует.")

    return {
        "Категория": category,
        "Ссылка на товар": product_url,
        "Ссылка на первое фото": image_link,
        "Название": name,
        "Цена": price,
        "Вес товара": weight_cakes,
        "Добавили в подборки": added_to_collections,
        "Магазин": shop,
        "Ссылка на магазин": shop_url,
        "Оценка магазина": shop_rating,
        "Количество оценок": rating,
        "Количество покупок": purchases
    }

# Сохранение в XLSX
def save_to_xlsx(data, filename):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Товары"

    # Заголовки
    headers = [
        "Категория",
        "Ссылка на товар",
        "Ссылка на первое фото",
        "Название",
        "Цена",
        "Вес товара",
        "Добавили в подборки",
        "Магазин",
        "Ссылка на магазин",
        "Оценка магазина",
        "Количество оценок",
        "Количество покупок"
    ]
    sheet.append(headers)

    # Данные
    for item in data:
        sheet.append([
            item.get("Категория"),
            item.get("Ссылка на товар"),
            item.get("Ссылка на первое фото"),
            item.get("Название"),
            item.get("Цена"),
            item.get("Вес товара"),
            item.get("Добавили в подборки"),
            item.get("Магазин"),
            item.get("Ссылка на магазин"),
            item.get("Оценка магазина"),
            item.get("Количество оценок"),
            item.get("Количество покупок"),
        ])
        # Настройка ширины столбцов и выравнивания
        for col in sheet.columns:
            max_length = 0
            col_letter = col[0].column_letter  # Получение буквы столбца
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = max_length + 2
            sheet.column_dimensions[col_letter].width = adjusted_width

            # Установка выравнивания
            for cell in col:
                cell.alignment = Alignment(horizontal="left", vertical="top")

    # Сохранение в файл
    workbook.save(filename)
    print(f"Данные сохранены в {filename}")

# Основная функция
def main():
    main_url = "https://flowwow.com/moscow/cakes/?cake_sort=1"

    # Шаг 1: Получить HTML главной страницы через Selenium
    main_page_source = get_page_source_with_selenium(main_url)

    # Шаг 2: Получить ссылки на товары
    product_links = get_product_links(main_page_source)
    print(f"Найдено {len(product_links)} товаров.")

    # Ограничить обработку первыми 10 товарами
    product_links = product_links[:10]
    print(f"Обрабатываем первые {len(product_links)} товаров.")

    # Шаг 3: Извлечь данные из каждой ссылки
    all_product_details = []
    for link in product_links:
        full_url = f"https://flowwow.com{link}"
        product_details = get_product_details(full_url)
        all_product_details.append(product_details)
        print(f"Извлечено: {product_details}")

        # Добавить задержку
        time.sleep(2)

    # Сохранение данных в XLSX
    save_to_xlsx(all_product_details, "products.xlsx")

if __name__ == "__main__":
    main()