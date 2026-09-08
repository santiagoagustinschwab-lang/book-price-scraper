# Imports the request library, which is used to request information from websites, and BeautifulSoup from bs4, which provides the raw HTML of the site
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# The request to the books to scrape site is stored in the web variable
web = requests.get("https://books.toscrape.com/?")
web.encoding = "utf-8"

# Then the "transformation" of the site's request into something readable is stored in the html variable. The .text attribute returns the data without the tags, and "html.parser" is an HTML "translator"
html = BeautifulSoup(web.text, "html.parser")

# Then the findAll function is used on html and stored in the prices variable, where it searches for all "p" tags that have the "price_color" class
prices = html.findAll("p", class_="price_color")
titles = html.findAll("a", title=True)

price_list = []
title_list = []
books = []

# Here "price" is iterated through prices so that all the elements in prices are printed using the .text attribute, removing the tags and giving a cleaner result
for price in prices:
    print(price.text)
    price_list.append(price.text)

for title in titles:
    print(title["title"])
    title_list.append(title["title"])

# Then price and title are packaged into the two lists using zip(). After that, the book variable is defined as a dictionary headed by "title": title, which stores the titles, and "price": price, which stores the prices. Then the dictionary is added to the books list
for price, title in zip(price_list, title_list):
    book = {"title": title, "price": price}
    books.append(book)

# Finally, the wb variable stores Workbook(), which creates an empty Excel file in memory
wb = Workbook()

# The sheet variable stores the active "sheet" or Excel worksheet, which is the first one by default
sheet = wb.active

# Adds the headers Titles | Prices to the worksheet
sheet.append(["Titles", "Prices"])

# Iterates through book in books (which is already a list of dictionaries containing the titles and prices) and adds book["title"] to sheet, which stores the titles, and book["price"], which stores the prices
for book in books:
    sheet.append([book["title"], book["price"]])

# Finally, the xlsx (Excel) file is saved in the folder where the .py file is located
wb.save("Books.xlsx")