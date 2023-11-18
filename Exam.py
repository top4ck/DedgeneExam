import requests
from bs4 import BeautifulSoup
import openpyxl

user_agent = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
headers = {"User-agent": user_agent}
session = requests.Session()
book = openpyxl.Workbook()
book.save("catalog.xlsx")
sheet = book.active
sheet["A1"] = "title"
sheet["B1"] = "reviews"
sheet["C1"] = "price"
count = 2

for j in range(1, 25):
    print(f"Page {j}")
    with open('catalog.txt', "a", encoding="utf-8") as file:
        url = f"https://allo.ua/ua/televizory/p-{j}/"
        response = session.get(url, headers=headers)

        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "lxml")
            all_products = soup.find_all("div", class_="products-layout__item")

            for i in all_products:
                if i.find('div', class_="product-card__content"):
                    title = i.find('a', class_='product-card__title').text
                    print(title)

                    try:
                        reviews_element = i.find('span', class_="review-button__text review-button__text--count")

                        if reviews_element and int(reviews_element.text) > 0:
                            reviews = reviews_element.text
                            print("К-сть вiдгукiв:", reviews)
                        else:
                            print("Вiдгукiв нема.")
                    except AttributeError:
                        print("-")

                    price_element = i.find('div', class_="v-pb__cur")

                    if price_element is not None:
                        price = price_element.text
                        print(price)
                    else:
                        print("Нема в наявностi")

                    file.write(f"{title} {reviews} {price} \n")
                    sheet[f"A{count}"] = title
                    sheet[f"B{count}"] = reviews
                    sheet[f"C{count}"] = price
                    count += 1

book.save("catalog.xlsx")
book.close()