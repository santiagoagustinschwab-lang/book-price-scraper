# Imports the request library, which is used to request information from websites, and BeautifulSoup from bs4, which provides the raw HTML of the site
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# The request to the books to scrape site is stored in the response variable
respuesta = requests.get("https://books.toscrape.com/?")
respuesta.encoding = "utf-8"
# Then the "transformation" of the site's request into something readable is stored in the sopa variable. The .text attributes return the data without the tags, and "html.parser" is an HTML "translator"
sopa = BeautifulSoup(respuesta.text, "html.parser")

# Then the findAll function is used on sopa and stored in the precios variable, where it searches for all "p" tags that have the "price_color" class
precios = sopa.findAll("p", class_="price_color")

lista_de_precios = []
lista_de_titulos = []
libros = []

# Here "precio" is iterated through precios so that all the elements in precios are printed using the .text attribute, removing the tags and giving a cleaner result
for precio in precios:
    print(precio.text)
    lista_de_precios.append(precio.text)

# Here the same thing is repeated but with titles
titulos = sopa.findAll("a", title=True)

for titulo in titulos:
    print(titulo["title"])
    lista_de_titulos.append(titulo["title"])

# Then price and title are packaged into the two lists using zip(). After that, the libro variable is defined as a dictionary headed by "titulo":titulo, which stores the titles, and "precio":precio, which stores the prices. Then the dictionary is added to the libros list
for precio, titulo in zip(lista_de_precios, lista_de_titulos):
    libro = {"titulo": titulo, "precio": precio}
    libros.append(libro)

# Finally, the wb variable stores Workbook(), which creates an empty Excel file in memory
wb = Workbook()

# The hoja variable stores the active "sheet" or Excel worksheet, which is the first one by default
hoja = wb.active

# Adds the headers Titulos | Precios to the worksheet
hoja.append(["Titulos", "Precios"])

# Iterates through libro in libros (which is already a list of dictionaries containing the titles and prices) and adds libro["titulos"] to hoja, which stores the titles, and precio, which stores the prices
for libro in libros:
    hoja.append([libro["titulo"], libro["precio"]])

# Finally, the xlsx (Excel) file is saved in the folder where the .py file is located
wb.save("Libros.xlsx")