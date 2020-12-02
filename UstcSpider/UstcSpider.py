'''
@Author: dong.zhili
@Date: 1970-01-01 08:00:00
LastEditors: Please set LastEditors
LastEditTime: 2020-09-28 21:35:02
@Description: 
'''

import requests
import json
import getpass
from lxml import etree
from retrying import retry
from xlsxwriter import Workbook

class UstcSpider():
    def __init__(self, total = 0, path = "./选课信息.xlsx"):
        # 登陆研究生信息平台
        self.url_login = "https://passport.ustc.edu.cn/login?service=http://yjs.ustc.edu.cn/default.asp"
        # 研究生信息平台上有“选课与成绩”页面的html
        self.url_info = "http://yjs.ustc.edu.cn/m_left.asp?area=5&menu=1"
        # 选课系统主页
        self.url_select = "https://jw.ustc.edu.cn/for-std/course-select"
        # 获取turnId的url
        self.url_openturn = "https://jw.ustc.edu.cn/ws/for-std/course-select/open-turns"
        # 读取课程信息
        self.url_lessons = "https://jw.ustc.edu.cn/ws/for-std/course-select/addable-lessons"
        # 读取选课人数
        self.url_counts = "https://jw.ustc.edu.cn/ws/for-std/course-select/std-count"
        # 软院选课信息字典的列表
        self.list_software = []
        # 最终保存文件的路径
        self.path = path
        # 连接选课服务器的session
        self.session = None
        # 设置请求头
        self.headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36 Edg/85.0.564.51"}

    # 取html字符串的底层方法
    @retry(stop_max_attempt_number = 3)
    def _get_html(self, url, data=None, cookies=None, encoding="utf-8"):
        # 发起get请求，设置超时时间5s
        if(self.session is not None):
            response = self.session.get(url, headers = self.headers, data = data, cookies = cookies, timeout = 5) # cookies = cookie_dict
        else:
            response = requests.get(url, headers = self.headers, data = data, cookies = cookies, timeout = 5)
        # 返回解码后的html字符串
        return response.content.decode(encoding) # 默认utf-8解码

    # 发送数据请求的底层方法
    @retry(stop_max_attempt_number = 3)
    def _post_data(self, url, data=None, cookies=None, encoding="utf-8"):
        # 发起post请求，设置超时时间5s
        if(self.session is not None):
            response = self.session.post(url, headers = self.headers, data = data, cookies = cookies, timeout = 5) # cookies = cookie_dict
        else:
            response = requests.post(url, headers = self.headers, data = data, cookies = cookies, timeout = 5)
        # 返回解码后的html字符串
        return response.content.decode(encoding) # 默认utf-8解码

    # 取html字符串的方法
    def get_html(self, url, data=None, cookies=None, encoding="utf-8"):
        try:
            html_str = self._get_html(url, data = data, cookies = cookies, encoding = encoding)
        except requests.exceptions.HTTPError:
            print("HTTPError")
            html_str = None
        return html_str

    # 发送form表单的方法
    def post_data(self, url, data=None, cookies=None, encoding="utf-8"):
        try:
            data_json = self._post_data(url, data = data, cookies = cookies, encoding = encoding)
        except requests.exceptions.HTTPError:
            print("HTTPError")
            data_json = None
        return data_json
    
    # 建立session，并登录研究生信息平台
    def create_session(self, url):
        # 使用session发送post请求，获取cookie
        username = input("\n[*]请输入用户名：")
        password = getpass.getpass("\n[*]请输入密码：")
        post_data = {"username":username, "password":password}
        
        self.session = requests.session()
        self.post_data(self.url_login, data = post_data, encoding="gbk")
        # 请求研究生信息平台主页，如果能正确拿到网页文件，则登录成功
        self.get_html(self.url_login, encoding="gbk")
    
    # 获取session中的cookie信息，dict类型
    def get_cookies(self):
        return self.session.cookies.get_dict()
        
    def run(self):
        self.create_session(self.url_login)
        # self.get_html("http://yjs.ustc.edu.cn/m_top.asp", encoding="gbk")
        # 获取cookie
        cookie_dict = self.get_cookies()

        # 获取研究生信息平台上有“选课与成绩”页面的html
        html_str = self.get_html(self.url_info, encoding="gbk")
        
        # 定位网上选课超链接的节点
        element = etree.HTML(html_str)
        td_list = element.xpath("//*[@id='mm_2']//@href")
        # print(td_list[0])
        
        # 模拟前端js机制找到合适的cookie组成新的url
        for key in cookie_dict.keys():
            if(-1 != key.find("ASPSESSION")):
                url_lesson = td_list[0]+'&'+key+'='+cookie_dict[key]
        # print(url_lesson)
        
        # 进入网上选课
        html_str = self.get_html(url_lesson, encoding="utf-8")

        # 取得bizTypeId和studentId
        response = self.session.get(self.url_select, headers = self.headers)
        
        # 访问重定向之后的url
        html_str = self.get_html(response.url, encoding="utf-8")
        # print(html_str)
        element = etree.HTML(html_str)
        list_js1_str = element.xpath("/html/body/script[1]/text()")
        
        list_bizTypeId = list_js1_str[0].replace(" ", "").replace(",", "").split('\n')[4].split(':')
        list_studentId = list_js1_str[0].replace(" ", "").replace(",", "").split('\n')[5].split(':')

        dict_temp = {}

        dict_temp[list_bizTypeId[0]] = list_bizTypeId[1]
        dict_temp[list_studentId[0]] = list_studentId[1]

        # 提交bizTypeId和studentId取得turnId
        dict_turnId = json.loads(self.post_data(self.url_openturn, data=dict_temp))
        turnId = dict_turnId[0]['id']
        print(turnId)
        
        # 提取软件学院课程列表
        list_ret = json.loads(self.post_data(self.url_lessons, data = {"turnId": str(turnId), "studentId": list_studentId[1]}))
        # 用一个列表list_id记录所有的id，用于后面发送选课人数的请求
        list_id = []
        for dict_temp in list_ret:
            if dict_temp["openDepartment"]["nameZh"] == "软件学院苏州":
                # 新建一个字典记录需要的数据
                dict_new = {}
                dict_new['编号'] = dict_temp['id']
                dict_new['课名'] = dict_temp['course']['nameZh']
                # print(type(dict_temp['teachers'][0]['nameZh']))
                # print(dict_temp['teachers'][0]['nameZh'])
                try:
                    dict_new['授课老师'] = dict_temp['teachers'][0]['nameZh']
                except IndexError:
                    dict_new['授课老师'] = "unkown"
                dict_new['课程类型'] = dict_temp['courseType']['nameZh']
                dict_new['限制人数'] = dict_temp['limitCount']
                dict_new['详情'] = dict_temp['dateTimePlace']['textZh']
                dict_new['授课语言'] = dict_temp['teachLang']['nameZh']
                self.list_software.append(dict_new)
                # 列表list_id收集当前的课程id
                list_id.append(str(dict_temp['id']))
        
        # 提取各课程报名人数，并加入self.list_software中的每个字典
        dict_ret = json.loads(self.post_data(self.url_counts, data = {"lessonIds[]":list_id}, cookies = cookie_dict))
        print(dict_ret)
        for dict_temp in self.list_software:
            dict_temp['选课人数'] = dict_ret[str(dict_temp['编号'])]

        print(self.list_software)
        self.save_infos(self.path)
        
        # html_str = self.post_data(self.url_counts, data = {"lessonIds[]":"130086"}, encoding="utf-8")
        # print(html_str)
        # print(json.loads(html_str))


    # 将infos列表中的数据存入xlsx
    def save_infos(self, path):
        wb = Workbook(path)
        ws = wb.add_worksheet()
        # 定义行列
        row = 0
        col = 0
        # 第一行写入列标题
        col_header_list = ["编号", "课名", "授课老师", "课程类型", "限制人数", "选课人数", "详情", "授课语言",]
        for col_header in col_header_list:
            col = col_header_list.index(col_header)
            ws.write(row, col, col_header)
        # 接下来写入infos中值的内容
        row = 1
        for info in self.list_software:
            for _key, _value in info.items():
                col = col_header_list.index(_key)
                ws.write(row, col, _value)
            row += 1
        # 关闭文件
        wb.close()

        # with open("软件学院苏州选课.txt", "w", encoding="utf-8") as f:
        #     f.write(json.dumps(self.list_software, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    spider = UstcSpider()
    spider.run()
    
    