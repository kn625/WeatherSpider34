# WeatherSpider34
This is a simple Weather-fetch program with python.
# Grab Weather Data on web.(网页上采集天气数据)


## Introducion(介绍)

每次在整点时，抓取网页(http://forecast.weather.com.cn/town/weather1dn/101080104005.shtml )上关于呼和浩特盛乐经济开发区实时天气数据，
保存到Excel.xls excel表格中。

## Excel format(excel 数据格式)

气温数据为整数格式，其余均为字符串格式

## Files(文件说明)

1. python34					--Dependency for .py files.(Python3解释器文件夹)未显示
2. README.txt				--An introduction about the Project.(该工程的介绍文件)
3. WeatherSpider.py			--Source Code for this project.(该工程源码)
4. 天气数据采集.bat			--An .bat files, Run the script 'WeatherSpider.py' with python3.
							(批处理文件，以运行WeatherSpider.py脚本)
5. Excel.xls				--A file saving the Data.(数据存储文件)  

## Usage(用法)

Double click 天气数据采集.bat, and the Program is running in the background. When there's error happening,
the program stops running, and the error logs are printed in the excel file.
(双击 天气数据采集.bat 文件，程序在后台运行。当异常发生时，程序结束运行，错误信息打印在excel文件中)
