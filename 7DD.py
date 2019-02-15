#!/usr/bin/python3
import json
import requests
import time
import xlwings as xw
import os
#当前时间
t=1534401238
#---------------------------------------------------- 获取商品种类配置 ----------------------------------------------------
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
typeFormdata = {
"token":"FE9D7B55FE084D334EF4B8DF9B542E3A",
"t":t,
"api":"shoppingmall/getallcate"
}
typeUrl = "http://7ddapi.7dingdong.com/shoppingmall/getallcate"
#---------------------------------------------------- 获取商品列表配置 ----------------------------------------------------
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
goodsFormdata = {
"api":"shoppingmall/goods_list",
"token":"264FE08D8302DFECFFE059FBEB7F39CE",
"tokenId":"9CA13F4759018DD48F46C3AE86EAA92C",
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
goodsUrl = "http://7ddapi.7dingdong.com/shoppingmall/goods_list"
#---------------------------------------------------- 工作簿配置 ----------------------------------------------------
pictureSize=10.58*10

#---------------------------------------------------- 分类采集 ----------------------------------------------------
def classify():
	app = xw.App(visible=False,add_book=False)
	book = app.books.add()
	sheets = book.sheets
	imgFilePath = "img"
	if not os.path.exists(imgFilePath):
		os.makedirs(imgFilePath)
		pass
	#1.获取商品的种类
	typeList=requests.post(url=typeUrl,data=typeFormdata,headers=typeHeader_dict).json()["data"]
	if typeList==None or len(typeList)<=0:
		print("获取商品种类失败",typeUrl)
		os._exit(0)
	#2.按商品种类创建不同的表,写入数据
	for goodsType in typeList:
		sheet = sheets.add(name=goodsType["cate_name"])
		sheet.range((1,1)).value="商品图片"
		sheet.range((1,2)).value="商品编号"
		sheet.range((1,3)).value="商品名称"
		sheet.range((1,4)).value="商品零售价"
		sheet.range((1,5)).value="商品标价"
		sheet.range((1,6)).value="商品尊享价"
		sheet.range((1,7)).value="详情"
		goodsFormdata["cate_id_1"]=goodsType["cate_id"]
		# 获取商品列表
		goodList=requests.post(url=goodsUrl,data=goodsFormdata,headers=goodsHeader_dict).json()["data"]
		if goodList==None or len(goodList)<=0:
			print("加载错误",goodsUrl)
			os._exit(0)
		#写入数据
		for x in range(len(goodList)):
			sheet.range((x+2,1)).row_height=pictureSize
			imgName = os.path.basename(goodList[x]['default_image'])
			with open(imgFilePath+"/"+imgName,"wb") as img:
				img.write(requests.get(goodList[x]['default_image']).content)
			sheet.pictures.add(os.path.abspath(imgFilePath+"/"+imgName),width=pictureSize,height=pictureSize,left=0,top=sheet.range("G1:G"+str(x+1)).height)
			sheet.range((x+2,2)).value=goodList[x]['goods_id']
			sheet.range((x+2,3)).value=goodList[x]['goods_name']
			sheet.range((x+2,4)).value=goodList[x]['retail_price']
			sheet.range((x+2,5)).value=goodList[x]['b2b_price']
			sheet.range((x+2,6)).value=goodList[x]['enjoy_price']
			sheet.range((x+2,7)).value=goodList[x]['url']
			print("写入"+goodList[x]['goods_id']+"号数据完成")
		print(goodsType["cate_name"]+"创建完成")
		sheet.autofit("c")
		sheet.range("A1").column_width=pictureSize/6
	print("所有数据获取完成")
	book.save('企叮咚.xls')
	book.close()
	pass
	return;
#---------------------------------------------------- 综合采集 ----------------------------------------------------
def composite():
	print("开始采集商品数据")
	app = xw.App(visible=False,add_book=False)
	book = app.books.add()
	sheets = book.sheets
	sheet = sheets.add("综合")
	sheet.range((1,1)).value="商品图片"
	sheet.range((1,2)).value="商品编号"
	sheet.range((1,3)).value="商品名称"
	sheet.range((1,4)).value="商品零售价"
	sheet.range((1,5)).value="商品标价"
	sheet.range((1,6)).value="商品尊享价"
	sheet.range((1,7)).value="详情"
	imgFilePath = "img"
	if not os.path.exists(imgFilePath):
		os.makedirs(imgFilePath)
		pass
	# 获取全部商品列表
	goodList=requests.post(url=goodsUrl,data=goodsFormdata,headers=goodsHeader_dict).json()["data"]
	if goodList==None or len(goodList)<=0:
		print("加载错误",goodsUrl)
		os._exit(0)
	#写入数据
	for x in range(len(goodList)):
		row = x + 2
		sheet.range((row,1)).row_height=pictureSize
		imgName = os.path.basename(goodList[x]['default_image'])
		with open(imgFilePath+"/"+imgName,"wb") as img:
			img.write(requests.get(goodList[x]['default_image']).content)
		sheet.pictures.add(os.path.abspath(imgFilePath+"/"+imgName),width=pictureSize,height=pictureSize,left=0,top=sheet.range("G1:G"+str(x+1)).height)
		sheet.range((row,2)).value=goodList[x]['goods_id']
		sheet.range((row,3)).value=goodList[x]['goods_name']
		sheet.range((row,4)).value=goodList[x]['retail_price']
		sheet.range((row,5)).value=goodList[x]['b2b_price']
		sheet.range((row,6)).value=goodList[x]['enjoy_price']
		sheet.range((row,7)).value=goodList[x]['url']
		print("写入"+goodList[x]['goods_id']+"号数据完成")
	sheet.autofit("c")
	sheet.range("A1").column_width=pictureSize/6
	print("所有商品采集完成")
	pass
	book.save('企叮咚.xls')
	book.close()
	return;
#---------------------------------------------------- 运行 ----------------------------------------------------
# 分类	
# classify()
# 综合
composite()
