# 爬虫思路汇总：
#   ①，https://book.douban.com/tag/  总书签首页
#       抓取豆瓣图书书签上所有的书签名字，并保存为一个数组
#       当输入一个标签时，根据标签去生成对应的网址。如果标签不存在数组中，提示帮助，然后显示这个数组
#   ②，多线程爬取豆瓣图书信息
#           1，爬取图书名字跟作者
#           2，爬取图书对应的链接
#           3，爬取图书的简介信息
#           4，爬取图书的豆瓣评分
#   ③，将数据储存在xls表格中，按标签分类命名xls文件。
#       按一定条件排序：评分或者默认排序
# 项目目的：
#           ①，熟悉BeautifulSoup与正则    ②，熟悉threading多线程
#           ③，熟悉表格输出操作openpyxl模块操作
#           ④，熟悉文件存取操作

import urllib.request, urllib.error
import re
from bs4 import BeautifulSoup
from lxml import etree
from user_agent.base import generate_user_agent


def url_open(url):
    head = {"User-Agent" : generate_user_agent()}
    req = urllib.request.Request(url, headers=head)
    response = urllib.request.urlopen(req).read()
    return response
    
# 获取首页书签，并存为表格
def get_bookmark(url):
    try:
        response = url_open(url).decode('utf-8')
        soup = BeautifulSoup(response,"html.parser")

        # 获取标签总分类的 tag 列表，可遍历.string获得文字
        categories = soup.find_all('h2', style="padding-top:10px",)

        # 获取《文学》下的标签分类
            # 匹配出文学分类下包含的所有标签内容
        culture_string = re.compile('(<a name="文学" class="tag-title-wrapper">\s(\s|.)*?\s</div>)').findall(response)[0][0]
        # 实例化对象为可xpath操作的对象，xpath返回列表
        html = etree.HTML(culture_string, parser=None, )
        culture = html.xpath('//td/a/text()')

        # 获取《流行》下的标签分类
        popular_string = re.compile('(<a name="文学" class="tag-title-wrapper">\s(\s|.)*?\s</div>)').findall(response)[0][0]
        # 实例化对象为可xpath操作的对象，xpath返回列表
        html = etree.HTML(culture_string, parser=None, )
        popular_string = html.xpath('//td/a/text()')

        '''
        for title in categories:
            yield (title.string.strip(' ·'))
        '''

    except urllib.error.URLError as reason:
        if hasattr(reason, 'code'):
            print(reason.code)
        if hasattr(reason, 'reason'):
            print(reason.reason) 

def fun(url):
    try:
        response = url_open(url).decode('utf-8')
        soup = BeautifulSoup(response,"html.parser")
        title = soup.find_all('h2')
        print(title)
    except urllib.error.URLError as reason:
        if hasattr(reason, 'code'):
            print(reason.code)
        if hasattr(reason, 'reason'):
            print(reason.reason)


if __name__ == "__main__":
    url = "https://book.douban.com/tag/"
    titles = get_bookmark(url)
    print(titles)
    #for i in titles:
     #   print(i)
        
