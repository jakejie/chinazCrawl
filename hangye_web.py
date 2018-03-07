# coding:utf-8
import requests
from lxml import etree
from xlwt import Workbook

book = Workbook()
sheet1 = book.add_sheet("第一张表")
result = []
for i in range(1, 1882):
    if i == 1:
        # url = "http://top.chinaz.com/hangye/index_shopping.html"
        url = "http://top.chinaz.com/all/index.html"
    else:
        # url = "http://top.chinaz.com/hangye/index_shopping_{}.html".format(i)
        url = "http://top.chinaz.com/all/index_{}.html".format(i)
    headers = {
        "Host": "top.chinaz.com",
        "Pragma": "no-cache",
        "Referer": url,
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36",
    }
    try:
        response = requests.get(url, headers=headers)
        response.encoding = 'utf-8'
        tree = etree.HTML(response.text)
        # info_list = tree.xpath('//*[@id="content"]/div[3]/div[3]/div[2]/ul/li')
        info_list = tree.xpath('//*[@id="content"]/div[3]/div[3]/div/ul/li')
        for index, info in enumerate(info_list):
            name = "".join(info.xpath('div[2]/h3/a/text()'))
            link = "".join(info.xpath('div[2]/h3/span/text()'))
            Alexa = "".join(info.xpath('div[2]/div[1]/p[1]/a/text()'))
            baidu = "".join(info.xpath('div[2]/div[1]/p[2]/a/img/@src'))[-5:-4]
            PR = "".join(info.xpath('div[2]/div[1]/p[3]/a/img/@src'))[-5:-4]
            Fan = "".join(info.xpath('div[2]/div[1]/p[4]/a/text()'))
            intro = "".join(info.xpath('div[2]/p/text()'))
            print(name, link, Alexa, baidu, PR, Fan, intro)
            result.append([name, link, Alexa, baidu, PR, Fan, intro])
    except Exception as e:
        print(e)
# 写表头
for k, j in enumerate(["序号", "网站名称", "网站链接", "Alexa排名", "百度权重", "PR", "反链数", "简介"]):
    sheet1.write(0, k, j)
for index, info in enumerate(result):
    # print(index, info)
    for k, j in enumerate(info):
        # 添加序号 在第一列
        if k == 0:
            sheet1.write(index + 1, k, index + 1)
        sheet1.write(index + 1, k + 1, j)
book.save('网站排行.xls')
