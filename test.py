from bs4 import BeautifulSoup
import requests , openpyxl

excel=openpyxl.Workbook()
sheet=excel.active
sheet.title="Movies list"
sheet.append(['Rank','Movie Name','Year of Release','IMDB Ratings'])


try:
    response=requests.get("https://www.imdb.com/chart/top/")
    soup=BeautifulSoup(response.text,'html.parser')
    movies=soup.find('tbody',class_='lister-list').find_all("tr")
    for movie in movies:
        # print(movie)
        rank=movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
        movie_name=movie.find('td',class_='titleColumn').a.text
        # print(movie_name)
        rate=movie.find('td',class_='ratingColumn').strong.text
        year=movie.find('td',class_='titleColumn').span.text.replace("("," ")
        year=year.replace(")"," ")
        # print(rank,movie_name,rate,year)
        sheet.append([rank,movie_name,year,rate])

    # print(soup)
except Exception as e:
    print(e)

excel.save('Movies_demo.xlsx')