from selenium import webdriver
from time import sleep as sl
from selenium.webdriver.support.ui import Select
import re
import datetime
import pymongo
import csv
import pymongo 
import pandas as pd
import matplotlib.pyplot as mp 
import seaborn 
import numpy as np
import matplotlib.dates as md
import warnings

client = pymongo.MongoClient("localhost",port = 27017)
db = client.txDay
collection = db.txDataNew2

today = str(input("請輸入今天日期YYYY/MM/DD: "))
SearchDay = str(input("請輸入查詢日期YYYY/MM/DD: "))

dateString1 = today
dateFormatter = "%Y/%m/%d"
a =datetime.datetime.strptime(dateString1, dateFormatter)

dateString2 = SearchDay
dateFormatter = "%Y/%m/%d"
b = datetime.datetime.strptime(dateString2, dateFormatter)

driver = webdriver.Chrome()
driver.get("https://www.taifex.com.tw/cht/3/futDailyMarketReport")
sl(0.5)
driver.find_element_by_name("queryDate").clear()
dataElement=driver.find_element_by_name("queryDate")
dataElement.send_keys(SearchDay)
sl(0.5)
submit = driver.find_element_by_id("button")
submit.click()

while True:		
	if a >= b :
		logObject = open("txL1.csv","a+")
		with logObject:
			logWriter = csv.writer(logObject)
		try:
			selectOp = Select(driver.find_element_by_name("MarketCode"))
			selectOp.select_by_value("0")
			r_list = driver.find_elements_by_xpath('//*[@id="printhere"]/table/tbody/tr[2]/td/table[2]/tbody/tr')
			sl(1)
			data = driver.find_element_by_xpath('//*[@id="printhere"]/table/tbody/tr[2]/td/h3')
			dataString = re.findall('\d+[/]\d+[/]\d+',data.text)

			xList = []
			for i in r_list:
				xList.append(i.text.split())
			titleList = xList[0]
			titleList[1] = "".join(titleList[1:4])
			del titleList[2:4]
			titleList[5] = "".join(titleList[5:7])
			del titleList[6]
			titleList.append("日期")
			
			del xList[-1]
			xList = [i+dataString for i in xList]
			print(xList)

			for i in xList[1:]:
				dataDict = {}
				for j,k in zip(titleList,i):
					dataDict[j] = k
				collection.insert(dataDict)
			print(dataDict)
				# with open("txData.txt","a+") as t:
				# 	t.write(str(dataDict)+"\n")
			logObject = open("txL1.csv","a+")
			with logObject:
				logWriter = csv.writer(logObject)
				logWriter.writerow([datetime.datetime.now(),"success"])
			sl(0.5)
		except Exception as e:
			print(e)
			logObject = open("txL1.csv","a+")
			with logObject:
				logWriter = csv.writer(logObject)
				logWriter.writerow([datetime.datetime.now(),"error:"+str(e)])			
		finally:
			submit = driver.find_element_by_id("button4")
			submit.click()			
			b += datetime.timedelta(days=1)	
			print(b)	
	else:
		break

driver.quit()


result = collection.find()
df = pd.DataFrame([i for i in result]).drop(columns = "_id")

df["漲跌價"] = np.array([re.findall("[-]\\d+",i) if "-" in i else re.findall("\\d+",i) for i in df["漲跌價"]])

df["漲跌%"] = [re.findall("[-]\\d+[.]\\d+",i)[0] if "-" in i else re.findall("\\d+[.]\\d+",i)[0] for i in df["漲跌%"]]

df["結算價"] = [i if i != "-" else int(pd.Series([int(i) for i in df["結算價"] if i != "-"]).mean()) for i in df["結算價"]]

df["日期"] = pd.to_datetime(df["日期"]) 

xlist = ["契約","到期月份(週別)","日期"]
for i in df.columns:
	if i in xlist:
		continue
	elif i == "漲跌%":
		df[i] = df[i].astype("float")
	else :
		df[i] = df[i].astype("int")

mp.rcParams['font.sans-serif'] = ['Microsoft YaHei']
mp.rcParams['axes.unicode_minus']=False

df2 = df.loc[df["到期月份(週別)"]=="202004",["日期","最高價","最低價"]].sort_values(by="日期")
mp.figure(u'2019/12 到期月份202001 折線圖', facecolor='lightgray', figsize=(18,10))
mp.title('2019/12 到期月份202001 折線圖', fontsize=30,c='red')
mp.xlabel('日期', fontsize=14)
mp.ylabel('金額', fontsize=14)
mp.plot(df2["日期"], df2["最低價"], linestyle='--',marker='o',markersize=8,c='limegreen',linewidth=3, label='low')
mp.plot(df2["日期"], df2["最高價"], linestyle='-.',marker='o',markersize=8,c='orange',linewidth=3, label='height')

mp.xlim(df2.iloc[0,0]-pd.Timedelta(hours=10) , df2.iloc[df2.shape[0]-1,0]+pd.Timedelta(hours=10) )
ax = mp.gca()
ax.xaxis.set_major_locator(md.DayLocator())
ax.xaxis.set_major_formatter(md.DateFormatter('%d %b %Y'))
for x,y in zip(df2["日期"], df2["最低價"]):
    mp.annotate("%s" %y, xy=(x,y), xytext=(0, 0), textcoords='offset points')
for x,y in zip(df2["日期"], df2["最高價"]):
    mp.text(x, y, "%s" %y, size = 8,\
         family = "fantasy", color = "r", style = "italic", weight = "light",bbox = dict(facecolor = "r", alpha = 0.2))
mp.tick_params(labelsize=10)
mp.grid(linestyle=':')
mp.legend()
mp.gcf().autofmt_xdate()
mp.show()