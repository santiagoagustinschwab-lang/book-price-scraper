# Importa las librerias request, que es para pedir informacion a citios web, y de bs4 Beautifulsoup, que da el html crudo del citio
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Se guarda en la variable respuesta el pedido de request del citio books to scrape
respuesta = requests.get("https://books.toscrape.com/?")
respuesta.encoding = "utf-8"
# Luego se guarda en la variable sopa la "transformacion" del request del citio a algo legible, los atrivutos .text returnan los datos sin las etiquetas, y "html.parser" es un "traductor de html"
sopa = BeautifulSoup(respuesta.text, "html.parser")

# Luego se guarda en la variable precios la funcion finAll a sopa donde busca todas las etiquetas "p" que tengan la clase "price_color"
precios = sopa.findAll("p", class_="price_color")

lista_de_precios = []
lista_de_titulos = []
libros = []

# Aqui se recorre "precio" en precios para que imprima todos los elementos en precios con el atributo .text asi eliminan las etiquetas y da un resulktado mas limpio
for precio in precios:
    print(precio.text)
    lista_de_precios.append(precio.text)

# Aqui se repite lo mismo pero con titulos
titulos = sopa.findAll("a", title=True)

for titulo in titulos:
    print(titulo["title"])
    lista_de_titulos.append(titulo["title"])

# Luego se empaqueta precio y titulo en las dos listas con zip(), despues se define la variable libro como un diccionario encabezado por "titulo":titulo que guarda los titulos y "precio":precio que guarda los precios, luego agrega el diccionario a la lista libros
for precio, titulo in zip(lista_de_precios, lista_de_titulos):
    libro = {"titulo": titulo, "precio": precio}
    libros.append(libro)

# Finalizando se guarda en la variable wb Workbook() la cual crea un archibo de exel vacio en memoria
wb = Workbook()

# Se guarda en la vaariable hoja la "sheet" o hoja de exel activa, por defecto la primera
hoja = wb.active

# Agrega al encabezado de la hoja titulos | precios
hoja.append(["Titulos", "Precios"])

# Recorre libro en libros (que ya es una lista dew diccionarios de los titulos y precios)y agrega a hoja libro["titulos"] que almacena los titulos, y precio que almacena los precios
for libro in libros:
    hoja.append([libro["titulo"], libro["precio"]])

# Finalmente se guarda en un archibo xlsx (exel) en la carpeta donde se ubica el archibo .py
wb.save("Libros.xlsx")