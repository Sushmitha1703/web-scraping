from bs4 import BeautifulSoup
import requests,openpyxl 

excel=openpyxl.Workbook()
print(excel.sheetnames)
sheet=excel.active
sheet.title="Top Rated shows"
print(excel.sheetnames)
sheet.append([' Rank','Name','Year of Release','IMDB rating'])

try:
  source= requests.get('https://www.imdb.com/chart/toptv/')
  source.raise_for_status()
  soup=BeautifulSoup(source.text,'html.parser')
  shows=soup.find('tbody',class_='lister-list').find_all('tr')
  for show in shows:
     name=show.find('td',class_='titleColumn').a.text
     year=show.find('td',class_='titleColumn').span.text.strip('()')
     rating=show.find('td',class_='ratingColumn imdbRating').strong.text
     rank=show.find('td',class_='titleColumn').get_text(strip=True).split('.')[0]
     print(rank,name,year,rating)
     sheet.append([rank,name,year,rating])
except Exception as e:
   print(e)

excel.save('IMDB show ratings.xlsx')