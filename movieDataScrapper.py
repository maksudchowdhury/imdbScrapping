# --------------------------------------------------------------------------------------------
# first prepare your environment by installing the following packages: (better if a virtual environment is used)
# pip install requests
# pip install bs4
# pip install openpyxl
# --------------------------------------------------------------------------------------------
import requests, openpyxl
from bs4 import BeautifulSoup as bs

try:
    source = requests.get("https://www.imdb.com/chart/top/")
    source.raise_for_status()
    excelFile = openpyxl.Workbook()
    sheet = excelFile.active
    sheet.title = "IMDB top movies Scrapped Data"
    sheet.append(['Serial No','Movie Name','Release Year','Rating'])
    soup = bs(source.text,'html.parser')
    movieTable = soup.find('tbody', class_='lister-list')
    tableRows = movieTable.find_all('tr')
    for count, i in enumerate(tableRows):
        serialNo = count+1
        movieName=(i.find('td',class_='titleColumn').find('a').text)
        releaseYear=(i.find('td',class_='titleColumn').find('span').text.strip('()'))
        rating=(i.find('td',class_='ratingColumn imdbRating').find('strong').text)
        sheet.append([serialNo,movieName,releaseYear,rating])
        print(serialNo,movieName,releaseYear,rating)
except Exception as e:
    print(e)

excelFile.save("IMDB top movies Scrapped Data.xlsx")
