from bs4 import BeautifulSoup
import requests , openpyxl
import re

excel=openpyxl.Workbook()
sheet=excel.active
sheet.title="Movies list"
sheet.append(['S_No','Movie Name','Year of Release','IMDB Ratings','Story','Director','Gross'])


try:
    response=requests.get("https://www.imdb.com/search/title/?genres=adventure&sort=user_rating,desc&title_type=feature&num_votes=25000,&pf_rd_m=A2FGELUUNOQJNL&pf_rd_p=94365f40-17a1-4450-9ea8-01159990ef7f&pf_rd_r=C4V8MT0FK71M0QPPAMJH&pf_rd_s=right-6&pf_rd_t=15506&pf_rd_i=top&ref_=chttp_gnr_2")
    soup=BeautifulSoup(response.text,'html.parser')
    movies=soup.find('div',class_='lister-list').find_all("div",class_='lister-item')
    for movie in movies:
        print(movie)
        index=movie.find('h3').find('span',class_='lister-item-index').get_text(strip=True).split('.')[0]
        name=movie.find('h3').a.text
        year=movie.find('h3').find('span',class_='lister-item-year').text
        year=re.sub("\D"," ",year)
        rate=movie.find('div',class_="ratings-imdb-rating").strong.text
        story=movie.find("p").findNext('p').get_text(strip=True)
        director=movie.find("p").findNext('p').findNext("p").a.text
        gross=movie.find("p",class_="sort-num_votes-visible").find_all('span')[-1].get_text()
        # print(index,name,year,rate,story,director,gross)
        sheet.append([index,name,year,rate,story,director,gross])

except Exception as e:
    print(e)

excel.save('Action_demo.xlsx')