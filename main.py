import requests
import time
import random
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Lists to collect data from every page, plus the final list of book dictionaries
titles_list = []
prices_list = []
books = []

# The site has 50 pages in total
num_pages = range(1, 51)

# Create the Excel file in memory
wb = Workbook()
sheet = wb.active

# Loop through every page: request it, check it loaded correctly, then parse and collect its data
for i in num_pages:
    try:
        web = requests.get(f"https://books.toscrape.com/catalogue/page-{i}.html")

        # If the page doesn't exist or failed to load, raise an error on purpose
        # so the except block can catch it and skip this page without stopping the whole loop
        if web.status_code == 404:
            raise ValueError(f"Error {web.status_code}")

        html = BeautifulSoup(web.text, "html.parser")

        # Titles are inside <a title="..."> tags (using the title attribute avoids truncated titles)
        titles = html.find_all("a", title=True)
        # Prices are inside <p class="price_color"> tags
        prices = html.find_all("p", class_="price_color")

        for title in titles:
            titles_list.append(title["title"])

        for price in prices:
            prices_list.append(price.text)

    except ValueError as error:
        # If this page failed, just report it and move on to the next one
        print(f"Error {error}")

    time.sleep(random.uniform(1, 3))

# Once every page has been collected, combine titles and prices into a list of dictionaries
for title, price in zip(titles_list, prices_list):
    book = {"titles": title, "prices": price}
    books.append(book)

# Write the header row and then one row per book
sheet.append(["TITLES", "PRICES"])

for book in books:
    sheet.append([book["titles"], book["prices"]])

wb.save("Books.xlsx")