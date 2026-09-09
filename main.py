import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Lists to collect data from every page, plus the final list of book dictionaries
title_list = []
price_list = []
books = []

# The site has 50 pages in total
num_of_pages = range(1, 51)

# Loop through every page: request it, parse it, and collect its titles and prices
for i in num_of_pages:
    web = requests.get(f"https://books.toscrape.com/catalogue/page-{i}.html")
    web.encoding = "utf-8"
    html = BeautifulSoup(web.text, "html.parser")

    # Prices are inside <p class="price_color"> tags
    prices = html.findAll("p", class_="price_color")

    # Titles are inside <a title="..."> tags (using the title attribute avoids truncated titles)
    titles = html.findAll("a", title=True)

    for title in titles:
        title_list.append(title["title"])

    for price in prices:
        price_list.append(price.text)

# Once every page has been collected, combine titles and prices into a list of dictionaries
for title, price in zip(title_list, price_list):
    book = {"title": title, "price": price}
    books.append(book)

# Create the Excel file and write the header row
wb = Workbook()
sheet = wb.active
sheet.append(["Title", "Price"])

# Write one row per book
for book in books:
    sheet.append([book["title"], book["price"]])

wb.save("Books.xlsx")