'''
@FileName:	WeatherSpider.py
@Author:	Corning
@Corp:		燕京啤酒
@Date:		2018.01.25
@Note:		抓取网页上关于呼和浩特天气数据，保存到excel表格中

'''



import urllib.request
from xlwt import Workbook
import os
import time
import xlrd 
from xlutils.copy import copy  

ErrorFileName 	= 'ExcelError.xls'
FileName 		= 'Excel.xls'
MINUTE			= 0

def SavePrevExcel():
	try:
		old_excel = xlrd.open_workbook(FileName, formatting_info=True)  
	except Exception as e:
		print("Failed!")
		return False
	
	else:
		# 将操作文件对象拷贝，变成可写的workbook对象  
		new_excel = copy(old_excel)  
	
		# 获得第一个sheet的对象  
		sheet1 = new_excel.get_sheet(0)  
		sheet2 = new_excel.get_sheet(1)
		sheet1.write(1,6,"数据读取错误")
		#os.popen("del "+FileName)
		new_excel.save(FileName)  #保存excel文件 
		print("Succeed!")
		return True
	
def FetchData():
	book = Workbook(encoding='gbk')    #如果采集数据有中文，需要添加这个 
	sheet1 = book.add_sheet('sheet1') #表格缓存 
	sheet2 = book.add_sheet('sheet2') #表格缓存 
	'''
	themin = time.localtime(time.time()).tm_min
	thehour = time.localtime(time.time()).tm_hour
	themon = time.localtime(time.time()).tm_mon
	theday = time.localtime(time.time()).tm_mday
	theyear = time.localtime(time.time()).tm_year
	thedate = (theyear * 10000 + themon * 100 + theday) * 100 + thehour + 1
	thedate_str = str(thedate)
	'''
	sheet1.write(0,0,"时间")  #写表格 
	sheet1.write(0,1,"气温")  #写表格 
	sheet1.write(0,2,"天气")  #写表格 
	sheet1.write(0,3,"风力")  #写表格 
	sheet1.write(0,4,"风向")  #写表格 
	sheet1.write(0,5,"湿度")  #写表格 
	sheet1.write(0,6,"错误报告")  #写表格 
	
	
	sheet2.write(0,0,"时间")  #写表格 
	sheet2.write(0,1,"气温")  #写表格 
	sheet2.write(0,2,"天气")  #写表格 
	sheet2.write(0,3,"风力")  #写表格 
	sheet2.write(0,4,"风向")  #写表格 
	
	
	check_url = r'http://forecast.weather.com.cn/town/weather1dn/101080104005.shtml' #网页地址
	#check_url = r'https://www.baidu.com/' #网页地址
	
	try:
		checkfile = urllib.request.urlopen(check_url).read()  #网页保存为文本文件		
	except Exception as e:
		print(e)
		sheet1.write(1,6,"网络数据读取错误")
		book.save(ErrorFileName)  #保存excel文件
		SavePrevExcel()
		return False;
	else:
		sheet1.write(1,6,"网络数据读取正确")
		checkfile = checkfile.decode('UTF-8')
		print(checkfile)

	
	#实时时间 realTime
	realTime_start = checkfile.find('实况</span>', 0) 
	if (realTime_start != -1):
		realTime = checkfile[realTime_start-5:realTime_start]
		sheet1.write(1,0,realTime)
	else:
		sheet1.write(1,0,"Error:网页改版,实时时间")
		book.save(ErrorFileName)  #保存excel文件 
		SavePrevExcel()
		return False;
	
	
	#实时温度 realTimeTemp
	realTimeTemp_start = checkfile.find('class=\"temp\"', 0)
	realTimeTemp_end = checkfile.find('</span>', realTimeTemp_start)
	try:
		realTimeTemp = int(checkfile[realTimeTemp_start+13:realTimeTemp_end])
	except Exception as e:
		sheet1.write(1,1,"Error:网页改版,实时温度")
		book.save(ErrorFileName)  #保存excel文件 
		SavePrevExcel()
		return False;
	else:
		sheet1.write(1,1,realTimeTemp)
	
	#实时天气 realTimeWthr
	realTimeWthr_start = checkfile.find('class=\"weather dis\"', 0)
	realTimeWthr_end = checkfile.find('</div>', realTimeWthr_start)
	if (realTimeWthr_start != -1)and(realTimeWthr_end != -1):
		realTimeWthr = checkfile[realTimeWthr_start+20:realTimeWthr_end]
		sheet1.write(1,2,realTimeWthr)
	else:
		sheet1.write(1,2,"Error:网页改版,实时天气")
		book.save(ErrorFileName)  #保存excel文件 
		SavePrevExcel()
		return False;
	
	
	#实时风向 风力 realTimeWindL  realTimeWindD
	realTimeWind_end = checkfile.find('级</span>', 0)
	realTimeWind_start = checkfile.find('<span>', realTimeWind_end-20)
	realTimeWind = checkfile[realTimeWind_start+6:realTimeWind_end+1]
	realTimeSpace = realTimeWind.find(' ', 0)
	if (realTimeWind_end != -1)and(realTimeWind_start != -1)and(realTimeSpace != -1):
		realTimeWindD = realTimeWind[0:realTimeSpace]
		realTimeWindL = realTimeWind[realTimeSpace+1:]
		sheet1.write(1,3,realTimeWindL)
		sheet1.write(1,4,realTimeWindD)
	else:
		sheet1.write(1,3,"Error:网页改版,实时风力")
		sheet1.write(1,4,"Error:网页改版,实时风向")
		book.save(ErrorFileName)  #保存excel文件 
		SavePrevExcel()
		return False;
	
	
	
	#实时相对湿度 realTimeRh
	realTimeRh_start = checkfile.find('<span>相对湿度', 0)
	realTimeRh_end = checkfile.find('</span>', realTimeRh_start)
	if (realTimeRh_start != -1)and(realTimeRh_end != -1):
		realTimeRh = checkfile[realTimeRh_start+11:realTimeRh_end]
		sheet1.write(1,5,realTimeRh)
	else:
		sheet1.write(1,5,"Error:网页改版,实时相对湿度")
		book.save(ErrorFileName)  #保存excel文件 
		SavePrevExcel()
		return False;
	
	hour_start = 0
	weather_end = 0
	temp_end = 0
	windL_end = 0
	windD_end = 0
	for i in range(24):
		#预报时间 hour
		hour_start = checkfile.find('\"time\"', hour_start+1)
		if (hour_start != -1):
			hour = checkfile[hour_start+8:hour_start+10] + '时'
			sheet2.write(i+1,0,hour)
		else:
			sheet2.write(i+1,0,"Error:网页改版,预报时间")
			book.save(ErrorFileName)  #保存excel文件 
			SavePrevExcel()
			return False;
		
		#预报天气 weather
		weather_start = checkfile.find('\"weather\"', weather_end)
		weather_end = checkfile.find('\"', weather_start+11)
		if (weather_start != -1)and(weather_end != -1):
			weather = checkfile[weather_start+11:weather_end]
			sheet2.write(i+1,2,weather)
		else:
			sheet2.write(i+1,2,"Error:网页改版,预报天气")
			book.save(ErrorFileName)  #保存excel文件 
			SavePrevExcel()
			return False;
		
		#预报温度 temperature
		try:
			temp_start = checkfile.find('\"temp\"', temp_end)
			temp_end = checkfile.find(',', temp_start+7)
			temperature = int(checkfile[temp_start+7:temp_end])
		except Exception as e:
			sheet2.write(i+1,1,"Error:网页改版,预报温度")
			book.save(ErrorFileName)  #保存excel文件 
			SavePrevExcel()
			return False;
		else:
			sheet2.write(i+1,1,temperature)
		
		#预报风力 windLevel
		windL_start = checkfile.find('\"windL\"', windL_end)
		windL_end = checkfile.find('\"', windL_start+9)
		if (windL_start != -1)and(windL_end != -1):
			windLevel = checkfile[windL_start+9:windL_end]
			sheet2.write(i+1,3,windLevel)
		else:
			sheet2.write(i+1,3,"Error:网页改版,预报风力")
			book.save(ErrorFileName)  #保存excel文件 
			SavePrevExcel()
			return False;
		
		#预报风向 windDirection
		windD_start = checkfile.find('\"windD\"', windD_end)
		windD_end = checkfile.find('\"', windD_start+9)
		if (windD_start != -1)and(windD_end != -1):
			windDirection = checkfile[windD_start+9:windD_end]
			sheet2.write(i+1,4,windDirection)
		else:
			sheet2.write(i+1,4,"Error:网页改版,预报风向")
			book.save(ErrorFileName)  #保存excel文件 
			SavePrevExcel()
			return False;
	
	#try:
		#os.popen("del "+FileName)	
	#finally:	
	book.save(FileName)  #保存excel文件 
	print('finish!')
	return True;


def main():
	FetchData()
	#os.popen("start .\\"+FileName)  #父亲打开
	while 1:
		themin = time.localtime(time.time()).tm_min
		if themin == MINUTE:
			webState = FetchData()
			if webState == False:
				time.sleep(610)
				#continue
				return False
			else:
				time.sleep(61)
				#time.sleep(3)
			
'''

def main():
	FetchData()
'''

		
if __name__ == '__main__':
    main()
