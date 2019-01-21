#!/usr/bin/python3
import json
import requests
import time
import xlwings as xw
import os
#当前时间
t=1534401238
#请求头
typeHeader_dict = {
"Accept": "application/json, text/plain, */*",
"Accept-Encoding": "gzip, deflate",
"Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
"Connection": "keep-alive",
"Content-Length": "81",
"Content-Type": "application/x-www-form-urlencoded",
"Host": "7ddapi.7dingdong.com",
"Origin": "http://www.7dingdong.com",
"Referer": "http://www.7dingdong.com/",
"User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36"
}
#请求地址:获取商品的种类和名称
typeUrl = "http://7ddapi.7dingdong.com/shoppingmall/getallcate"
#请求参数
typeFormdata = {
"token":"FE9D7B55FE084D334EF4B8DF9B542E3A",
"t":t,
"api":"shoppingmall/getallcate"
}
#请求一下商品种类
typeResponse=requests.post(url=typeUrl,data=typeFormdata,headers=typeHeader_dict)
typeJsonResult=typeResponse.json()
#获得结果
typeJsonResult=typeJsonResult["data"]
if typeJsonResult==None or len(typeJsonResult)<=0:
	print("获取商品种类失败",typeUrl)
	os._exit(0)
#请求地址:获取商品的地址
goodsUrl = "http://7ddapi.7dingdong.com/shoppingmall/goods_list"
#请求头,通用
goodsHeader_dict = {
"Accept":"application/json, text/plain, */*",
"Accept-Encoding":"gzip, deflate",
"Accept-Language":"zh-CN,zh;q=0.9,en;q=0.8",
"Connection":"keep-alive",
"Content-Length":"189",
"Content-Type":"application/x-www-form-urlencoded",
"Host":"7ddapi.7dingdong.com",
"Origin":"http://www.7dingdong.com",
"Referer":"http://www.7dingdong.com/",
"User-Agent":"Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36"
}
#请求参数
goodsFormdata = {
"api":"shoppingmall/goods_list",
"token":"264FE08D8302DFECFFE059FBEB7F39CE",
"tokenId":"3F5C9BB8F563A18FF44472F931F1FFF4",
"t":t,
"limit":"10000",
"order":"Desc",
"cate_id_1":"0",
"cate_id":"0",
"keywords":"",
"page":"1",
"sort":"0",
"zhe":"0",
"c":"0"
}
#打开表格
app = xw.App(visible=False,add_book=False)
book = app.books.add()
sheets = book.sheets
#图片的大小
size=10.58*10
#创建图片文件夹
imgFilePath = "img"
if not os.path.exists(imgFilePath):
	os.makedirs(imgFilePath)
	pass
#往表格里写输入
for item in typeJsonResult:
	goodsFormdata["cate_id_1"]=item["cate_id"]
	sheet = sheets.add(name=item["cate_name"])

	#发起请求
	response=requests.post(url=goodsUrl,data=goodsFormdata,headers=goodsHeader_dict)
	jsonResult=response.json()
	#获得结果
	jsonResult=jsonResult["data"]#解析json
	if jsonResult==None or len(jsonResult)<=0:
		print("加载错误",goodsUrl)
		os._exit(0)
	#写入表头
	sheet.range((1,1)).value="商品编号"
	sheet.range((1,1)).value="商品名称"
	sheet.range((1,2)).value="商品零售价"
	sheet.range((1,3)).value="商品标价"
	sheet.range((1,4)).value="商品尊享价"
	sheet.range((1,5)).value="详情"
	sheet.range((1,6)).value="商品图片"
	#写入数据
	for x in range(len(jsonResult)):
		print("写入"+jsonResult[x]['goods_id']+"号数据")
		sheet.range((x+2,1)).value=jsonResult[x]['goods_id']
		sheet.range((x+2,2)).value=jsonResult[x]['goods_name']
		sheet.range((x+2,3)).value=jsonResult[x]['retail_price']
		sheet.range((x+2,4)).value=jsonResult[x]['b2b_price']
		sheet.range((x+2,5)).value=jsonResult[x]['enjoy_price']
		sheet.range((x+2,6)).value=jsonResult[x]['url']
		sheet.range((x+2,7)).column_width=size/6
		sheet.range((x+2,7)).row_height=size
		imgName = os.path.basename(jsonResult[x]['default_image'])
		with open(imgFilePath+"/"+imgName,"wb") as img:
			img.write(requests.get(jsonResult[x]['default_image']).content)
		sheet.pictures.add(os.path.abspath(imgFilePath+"/"+imgName),width=size,height=size,left=sheet.range("A1:F1").width,top=sheet.range("G1:G"+str(x+1)).height)
	print(item["cate_name"]+"创建完成")
	sheet.autofit("c")
print("所有数据获取完成")
#保存表格
book.save('企叮咚.xls')
book.close()
