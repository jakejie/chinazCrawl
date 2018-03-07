# coding:utf-8
import requests
from lxml import etree
from xlwt import Workbook

book = Workbook()
sheet1 = book.add_sheet("第一张表")

first_url = "https://guanwangdaquan.com/"


def crawl_response(url):
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate, sdch, br",
        "Accept-Language": "zh-CN,zh;q=0.8",
        "Cache-Control": "max-age=0",
        "Connection": "keep-alive",
        # "Host": "guanwangdaquan.com",, headers=headers
        "If-Modified-Since": "Mon, 26 Feb 2018 08:59:03 GMT",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) \
                Chrome/55.0.2883.87 Safari/537.36",
    }
    response = requests.get(url)
    return response.text


def crawl_type(url):
    info = []
    response = crawl_response(url)
    tree = etree.HTML(response)
    type_list = tree.xpath('//*[@id="categorylist1"]/li')
    for types in type_list:
        category = "".join(types.xpath('a/text()'))
        # link = types.xpath('a/@href')
        two_type = types.xpath('ul/li')
        for ty in two_type:
            name = "".join(ty.xpath('a/text()'))
            link = "".join(ty.xpath('a/@href'))
            info.append([category, name, link])
    return info


def crawl_info():
    intro = []
    for info in crawl_type(first_url):
        response = crawl_response(info[2])
        tree = etree.HTML(response)
        product_list = tree.xpath('//*[@id="artilepaging"]/div')
        for product in product_list[:-1]:
            product_name = "".join(product.xpath('div[1]/a/@title'))
            product_link = "".join(product.xpath('div[1]/a/@href'))
            product_intro = "".join(product.xpath('div[2]/p/text()'))
            intro.append([info[0], info[1], info[2], product_name, product_link, product_intro])
    return intro


for index, intro in enumerate(crawl_info()):
    response = crawl_response(intro[4])
    tree = etree.HTML(response)
    introduce = "".join(tree.xpath('//*[@id="current-content"]/div[2]/div[1]/blockquote/p/text()'))
    guanwang = tree.xpath('//*[@id="post-22491"]/ul/li')
    urls = []
    for url in guanwang:
        url_link = "".join(url.xpath('a/@href'))
        url_name = "".join(url.xpath('/text()'))
        urls.append("-".join([url_link, url_name]))
    print(intro[0], intro[1], intro[2], intro[3], intro[4], intro[5], introduce, "/".join(urls))
    result = [intro[0], intro[1], intro[2], intro[3], intro[4], intro[5], introduce, "/".join(urls)]
    for i, j in enumerate(result):
        sheet1.write(index, i, j)
    # print('\n\n\n')
book.save('demo.xls')
