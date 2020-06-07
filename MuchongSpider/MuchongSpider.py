'''
@Author: dong.zhili
@Date: 1970-01-01 08:00:00
@LastEditors: dong.zhili
@LastEditTime: 2020-06-05 11:38:42
@Description: 
'''

import requests
from lxml import etree
from retrying import retry
from xlsxwriter import Workbook

class MuchongSpider():
    def __init__(self, url, total = 0, path = "./调剂信息.xlsx"):
        # 临时的url地址，一般是待爬取页面集合的首页
        self.url_temp = url
        # 待爬取的页数
        self.total = total
        # 调剂信息字典的列表
        self.info_list = []
        # 最终保存文件的路径
        self.path = path
        # 过滤字符串列表
        self.white_list = ["计算机", "软件", "通信", "电信", "自动控制", "自控", "控制科学",]
    
    # 取html字符串的底层方法
    @retry(stop_max_attempt_number = 3)
    def _get_html(self, url):
        # 设置请求头
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36 Edg/83.0.478.44"}
        # 发起get请求，设置超时时间5s
        response = requests.get(url, headers = headers, timeout = 5)
        # 返回解码后的html字符串
        return response.content.decode("gbk")
    
    # 取html字符串的方法
    def get_html(self, url):
        try:
            html_str = self._get_html(url)
        except requests.exceptions.HTTPError:
            print("HTTPError")
            html_str = None
        return html_str

    # 取总页数的方法
    def get_total(self):
        # 获取html字符串
        html_str = self.get_html(self.url_temp)
        # 定位显示总页数的节点
        element = etree.HTML(html_str)
        td_list = element.xpath("//div[@class='xmc_fr xmc_Pages xmc_tm10 solid']//td[@class='header']/text()")
        # 取td_list中后面一项字符串再用/切割取后者
        td = td_list[1].split('/')
        return int(td[1])
    
    # 白名单过滤
    def white_list_pass(self, text:str, white_list:list):
        for white in white_list:
            if -1 is not text.find(white):
                return True
        return False

    # 读取某一页的调剂信息
    def get_page(self, url):
        html_str = self.get_html(url)
        element = etree.HTML(html_str)
        # 读取该页存放调剂数据的列表
        item_list = element.xpath("//tbody[@class='forum_body_manage']/tr")
        # 遍历列表
        for i, item in enumerate(item_list, start = 1):
            print("  第{}项".format(i))
            # 读取标题
            title = item.xpath("./td/a/text()")[0] if item.xpath("./td/a/text()") else "暂无数据"
            # 读取学校
            school = item.xpath("./td[2]/text()")[0] if item.xpath("./td[2]/text()") else "暂无数据"
            # 读取门类
            category = item.xpath("./td[3]/text()")[0] if item.xpath("./td[3]/text()") else "暂无数据"

            # 白名单过滤，因为读取详细内容比较慢，放在此位置可以快速过滤不需要的数据
            if False is self.white_list_pass(title, self.white_list) and False is self.white_list_pass(category, self.white_list):
                continue
            
            # 读取招生人数
            quota = item.xpath("./td[4]/text()")[0] if item.xpath("./td[4]/text()") else "暂无数据"
            # 读取发布时间
            time = item.xpath("./td[5]/text()")[0] if item.xpath("./td[5]/text()") else "暂无数据"
            # 读取详细内容
            if item.xpath("./td/a/@href"):
                a_url = item.xpath("./td/a/@href")[0]
                a_elm = etree.HTML(self.get_html(a_url))
                addition = a_elm.xpath("string(//tbody[@id='pid1']//div[@class='t_fsz']//td[@valign='top'])")
            else:
                a_url = "暂无数据"
                addition = "暂无数据"
            # 将表项内容以字典格式存与infos
            self.info_list.append({"标题": title, "学校": school, "专业":category, "招生人数":quota, "时间":time, "原文链接": a_url, "发布内容": addition,})
    
    # 将infos列表中的数据存入xlsx
    def save_infos(self, path):
        wb = Workbook(path)
        ws = wb.add_worksheet()
        # 定义行列
        row = 0
        col = 0
        # 第一行写入列标题
        col_header_list = ["标题", "学校", "专业", "招生人数", "时间", "原文链接", "发布内容",]
        for col_header in col_header_list:
            col = col_header_list.index(col_header)
            ws.write(row, col, col_header)
        # 接下来写入infos中值的内容
        row = 1
        for info in self.info_list:
            for _key, _value in info.items():
                col = col_header_list.index(_key)
                ws.write(row, col, _value)
            row += 1
        # 关闭文件
        wb.close()
    
    def run(self):
        # 如果未设置待爬取的页数，则自动爬取最大页数
        if 0 == self.total:
            self.total = self.get_total()
        # 构造url_list
        url_list = [self.url_temp.format(i) for i in range(1, self.total+1)]
        
        # self.get_page(self.url_temp)
        
        # 遍历url_list读取数据
        for i, url in enumerate(url_list, start = 1):
            print("第{}页".format(i))
            self.get_page(url)
        self.save_infos(self.path)


if __name__ == '__main__':
    spider = MuchongSpider("http://muchong.com/bbs/kaoyan.php?action=adjust&year=2020&type=1&r1%5B%5D=08&page={}")
    spider.run()
    
    