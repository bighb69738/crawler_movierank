import requests
import pandas as pd
from bs4 import BeautifulSoup
import requests

import xlwt


contents1 = []
contents2 = []
dict1 = {}
for i in range(10):
	year = 2009+i
	param = {"search_year": str(year)} # 將年份自動轉換 從2009-2019 
	url = 'https://movies.yahoo.com.tw/chart.html?cate=rating'
	resp = requests.get(url, params=param)
	resp.encoding = 'utf-8' 
	soup = BeautifulSoup(resp.text, 'lxml')
	rows = soup.find_all('div', class_='tr')
	colname = list(rows.pop(0).stripped_strings) 
	colname.remove('預告片')
	colname.remove('上映日期')
	colname.remove('排名')


	for row in rows:
	    rank = row.find_next('div',attrs={'class':'td'})
	    updown = rank.find_next('div')
	    lastweek_rank = updown.find_next('div')
	
	    if rank.string == str(1):
	        movie_title = rank.find_next('h2')
	    else:
	        movie_title = rank.find_next('div',attrs={'class':'rank_txt'})

	    stars = row.find('h6',attrs={'class':'count'})
	    movie_name = [movie_title.string]
	    movie_star = [stars.string]



	    if (float(stars.string)>3.9):
	        contents1.append(movie_name)
	        contents2.append(movie_star)
	    else:
	        pass


for x in range(len(contents1)):
	dict1[contents1[x][0]]=contents2[x][0]

dict1 = sorted(dict1.items(), key=lambda d: d[1], reverse=True) #星星數由大至小排列

#以下由 Xlwt 寫至 Excel  

workbook = xlwt.Workbook(encoding='utf-8')
booksheet = workbook.add_sheet('Sheet 1', cell_overwrite_ok=True)

booksheet.write(0,0,'MOVIE')
booksheet.write(0,1,'STARTS')

for count in range(len(dict1)):
    booksheet.write(1+count,0,dict1[count][0])
    booksheet.write(1+count,1,dict1[count][1])


workbook.save('/home/vic/Downloads/MOVIE_rank.xls')
