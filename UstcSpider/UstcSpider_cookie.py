'''
Author: your name
Date: 2021-03-03 21:11:14
LastEditTime: 2021-03-07 03:34:52
LastEditors: Please set LastEditors
Description: In User Settings Edit
FilePath: /Spider/UstcSpider/UstcSpider_cookie.py
'''

import getpass
import json
import smtplib
import time
from email.header import Header
from email.mime.text import MIMEText
from enum import Enum

import requests
from lxml import etree
from retrying import retry
from xlsxwriter import Workbook


class Mode(Enum):
    """
    枚举工作状态
    """
    Unknown = 0
    LoopCatch = 1
    LoopQuery = 2
    SingleQuery = 3


class EmailInform(object):
    """
    EmailInform类，发送邮件通知
    """

    def __init__(self, sender, receivers, message):
        self.sender = 'from@runoob.com'
        self.receivers = ['429240967@qq.com']
        self.message = MIMEText('Python 邮件发送测试...', 'plain', 'utf-8')
        self.message['From'] = Header("菜鸟教程", 'utf-8')   # 发送者
        self.message['To'] = Header("测试", 'utf-8')        # 接收者
        subject = 'Python SMTP 邮件测试'
        self.message['Subject'] = Header(subject, 'utf-8')

    def inform(self):
        try:
            smtpObj = smtplib.SMTP('localhost')
            smtpObj.sendmail(self.sender, self.receivers,
                             self.message.as_string())
            print("邮件发送成功")
        except smtplib.SMTPException:
            print("Error: 无法发送邮件")


class UstcSpider(object):
    """
    UstcSpider类，网络爬虫，模拟浏览器与研究生信息平台网站交互
    """

    def __init__(self, path="./选课信息.xlsx"):
        # 最终保存文件的路径
        self.path = path
        # 连接选课服务器的session
        self.session = None
        # 设置请求头
        self.headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36 Edg/85.0.564.51", \
            "cookie": ""}

        self.workMode = Mode.Unknown  # 默认Unknown模式
        # 待查询课程字典，每个元素键为'id'，值为字典，键为{'limit', 'count', 'prevCount'}
        self.listQueryLessons = {}
        # 待抢课课程字典，每个元素键为'id'，值为字典，键为{'limit', 'count', 'prevCount', 'replaceId'}
        self.listCatchLessons = {}

    @retry(stop_max_attempt_number=3)
    def _getHtml(self, url, data=None, cookies=None, encoding="utf-8"):
        """
        取html字符串的底层方法，发起get请求，设置超时时间5s
        """
        if(self.session is not None):
            response = self.session.get(
                url, headers=self.headers, data=data, cookies=cookies, timeout=5)  # cookies = cookie_dict
        else:
            response = requests.get(
                url, headers=self.headers, data=data, cookies=cookies, timeout=5)
        # 返回解码后的html字符串
        return response.content.decode(encoding)  # 默认utf-8解码

    @retry(stop_max_attempt_number=3)
    def _postData(self, url, data=None, cookies=None, encoding="utf-8"):
        """
        发送数据请求的底层方法，发起post请求，设置超时时间5s
        """
        if(self.session is not None):
            response = self.session.post(
                url, headers=self.headers, data=data, cookies=cookies, timeout=5)  # cookies = cookie_dict
        else:
            response = requests.post(
                url, headers=self.headers, data=data, cookies=cookies, timeout=5)
        # 返回解码后的html字符串
        return response.content.decode(encoding)  # 默认utf-8解码

    def getHtml(self, url, data=None, cookies=None, encoding="utf-8"):
        """
        封装后的取html字符串的方法
        """
        try:
            html_str = self._getHtml(
                url, data=data, cookies=cookies, encoding=encoding)
        except requests.exceptions.HTTPError:
            print("HTTPError")
            html_str = None
        return html_str

    def postData(self, url, data=None, cookies=None, encoding="utf-8"):
        """
        封装好的发送form表单的方法
        """
        try:
            data_json = self._postData(
                url, data=data, cookies=cookies, encoding=encoding)
        except requests.exceptions.HTTPError:
            print("HTTPError")
            data_json = None
        return data_json

    def inputMode(self):
        """
        获取用户输入，得到工作模式
        """
        ret = input("你想进入抢课模式还是查询模式？(lcatch, lquery, squery)：")
        if "lcatch" == ret:
            self.workMode = Mode.LoopCatch
            listCatchIds = input("请输入想抢课程的id：").split(" ")
            for double_id in listCatchIds:
                catchLessonData = {}
                id = double_id.split(":")[0]
                replaceId = double_id.split(":")[1]
                catchLessonData['ZhName'] = None
                catchLessonData['limit'] = None
                catchLessonData['count'] = None
                catchLessonData['prevCount'] = None
                catchLessonData['replaceId'] = replaceId
                self.listCatchLessons[id] = catchLessonData
        elif "lquery" == ret:
            self.workMode = Mode.LoopQuery
            listQueryIds = input("请输入想查课程的id：").split(" ")
            print(listQueryIds)
            if listQueryIds == ['']:
                return
            for id in listQueryIds:
                queryLessonData = {}
                queryLessonData['ZhName'] = None
                queryLessonData['limit'] = None
                queryLessonData['count'] = None
                queryLessonData['prevCount'] = None
                self.listQueryLessons[id] = queryLessonData
        elif "squery" == ret:
            self.workMode = Mode.SingleQuery
            listQueryIds = input("请输入想查课程的id：").split(" ")
            if listQueryIds == ['']:
                return
            for id in listQueryIds:
                self.listQueryLessons[id] = None
        else:
            print("参数错误")
            exit()

    def getCookies(self):
        """
        获取session中的cookie信息，dict类型
        """
        return self.session.cookies.get_dict()

    def getSoftwareLessons(self, turnId, studentId):
        """
        获取软件学院所有课程列表，列表每一项是一个字典结构，注意：无已选人数信息
        """
        # 读取可选课程信息的url，用于获取软件学院课程列表
        urlAddableLessons = "https://jw.ustc.edu.cn/ws/for-std/course-select/addable-lessons"
        # 软院选课信息字典的列表
        listSoftwareLessons = []

        # 提取软件学院课程列表
        list_ret = json.loads(self.postData(urlAddableLessons, data={
            "turnId": str(turnId), "studentId": str(studentId)}))
        for dict_temp in list_ret:
            if dict_temp["openDepartment"]["nameZh"] == "软件学院苏州":
                # 新建一个字典记录需要的数据
                dict_new = {}
                dict_new['编号'] = str(dict_temp['id'])
                dict_new['课名'] = dict_temp['course']['nameZh']
                try:
                    dict_new['授课老师'] = dict_temp['teachers'][0]['nameZh']
                except IndexError:
                    dict_new['授课老师'] = "unkown"
                dict_new['课程类型'] = dict_temp['courseType']['nameZh']
                dict_new['限制人数'] = dict_temp['limitCount']
                dict_new['详情'] = dict_temp['dateTimePlace']['textZh']
                dict_new['授课语言'] = dict_temp['teachLang']['nameZh']
                listSoftwareLessons.append(dict_new)
        return listSoftwareLessons

    def run(self):
        # 读取选课人数的url
        urlStdCount = "https://jw.ustc.edu.cn/ws/for-std/course-select/std-count"
        # 选课请求url
        urlAddRequest = "https://jw.ustc.edu.cn/ws/for-std/course-select/add-request"
        # 退课请求url
        urlDropRequest = "https://jw.ustc.edu.cn/ws/for-std/course-select/drop-request"
        # 选退课响应url
        urlAddDropResponse = "https://jw.ustc.edu.cn/ws/for-std/course-select/add-drop-response"
        
        self.inputMode()
        # self.getHtml("http://yjs.ustc.edu.cn/m_top.asp", encoding="gbk")
        # turnId = input("请输入turnId: ")
        turnId = 
        # studentId = input("请输入studentId: ")
        studentId = 

        listSoftwareLessons = self.getSoftwareLessons(turnId, studentId)
        if self.workMode is Mode.SingleQuery:
            # 装完整数据的列表
            listQueryLessons = []
            # 只包含待查询id的列表
            if len(self.listQueryLessons) != 0:  # 如果有输入待查询课程的id列表则只存储想要的课程信息
                listQueryIds = self.listQueryLessons.keys()
                for softwareLesson in listSoftwareLessons:
                    if softwareLesson['编号'] in listQueryIds:
                        listQueryLessons.append(softwareLesson)
            else:  # 如果未输入待查询课程的id列表则存储所有软院课程信息
                listQueryIds = []
                for softwareLesson in listSoftwareLessons:
                    listQueryIds.append(softwareLesson['编号'])
                listQueryLessons = listSoftwareLessons
            # print(listQueryLessons)
            # 提取各课程报名人数，并加入queryLessons中的每个字典
            dict_ret = json.loads(self.postData(urlStdCount, data={
                "lessonIds[]": listQueryIds}))
            for lesson in listQueryLessons:
                lesson['选课人数'] = dict_ret[lesson['编号']]
            print(listQueryLessons)
            self.saveInfos(listQueryLessons, self.path)

        if self.workMode is Mode.LoopQuery:
            # 只包含待查询id的列表
            listQueryIds = []
            if len(self.listQueryLessons) != 0:  # 如果有输入待查询课程的id列表则只查询想要的课程信息
                listQueryIds = self.listQueryLessons.keys()
            else:
                for softwareLesson in listSoftwareLessons:
                    listQueryIds.append(softwareLesson['编号'])
                    queryLessonData = {}
                    queryLessonData['ZhName'] = None
                    queryLessonData['limit'] = None
                    queryLessonData['count'] = None
                    queryLessonData['prevCount'] = None
                    self.listQueryLessons[softwareLesson['编号']
                                          ] = queryLessonData
            
            # 发起一次查询，得到首次数据
            prevDictRet = json.loads(self.postData(urlStdCount, data={
                "lessonIds[]": listQueryIds}))
            
            # 向self.listQueryLessons填充数据
            for softwareLesson in listSoftwareLessons:
                if softwareLesson['编号'] in listQueryIds:
                    self.listQueryLessons[softwareLesson['编号']
                                          ]['ZhName'] = softwareLesson['课名']
                    self.listQueryLessons[softwareLesson['编号']
                                          ]['limit'] = softwareLesson['限制人数']
                    self.listQueryLessons[softwareLesson['编号']
                                          ]['count'] = prevDictRet[softwareLesson['编号']]
                    # prevCount未赋值

            print(self.listQueryLessons)
            while(True):
                dictRet = json.loads(self.postData(urlStdCount, data={
                    "lessonIds[]": listQueryIds}))
                if dictRet != prevDictRet:
                    # 数据发生了变化
                    for queryId in self.listQueryLessons.keys():
                        self.listQueryLessons[queryId]['prevCount'] = self.listQueryLessons[queryId]['count']
                        self.listQueryLessons[queryId]['count'] = dictRet[queryId]
                    print(self.listQueryLessons)
                prevDictRet = dictRet
                time.sleep(3)

        if self.workMode is Mode.LoopCatch:
            # 抢课模式，不考虑self.listCatchLessons为空的特殊处理，为空直接退出
            # 先对所有目标课程发起一次查询，得到首次数据
            # 发起一次查询，得到首次数据
            listCatchIds = self.listCatchLessons.keys()
            prevDictRet = json.loads(self.postData(urlStdCount, data={
                "lessonIds[]": listCatchIds}))
            # 向self.listCatchLessons填充数据
            for softwareLesson in listSoftwareLessons:
                if softwareLesson['编号'] in listCatchIds:
                    self.listCatchLessons[softwareLesson['编号']
                                          ]['ZhName'] = softwareLesson['课名']
                    self.listCatchLessons[softwareLesson['编号']
                                          ]['limit'] = softwareLesson['限制人数']
                    self.listCatchLessons[softwareLesson['编号']
                                          ]['count'] = prevDictRet[softwareLesson['编号']]
                    self.listCatchLessons[softwareLesson['编号']
                                          ]['prevCount'] = prevDictRet[softwareLesson['编号']]
                    # replaceId在input时赋值
            print(self.listCatchLessons)

            while len(self.listCatchLessons) != 0:
                print("剩余抢课项目：%s" % self.listCatchLessons)
                listCatchIds = self.listCatchLessons.keys()
                dictRet = json.loads(self.postData(urlStdCount, data={
                    "lessonIds[]": listCatchIds}))

                # 这里不对self.listCatchLessons.keys()强制转list会导致选到一次课则程序终止，但是目前需要利用这个bug，后续待优化流程
                for catchId in self.listCatchLessons.keys():
                    if dictRet[catchId] < self.listCatchLessons[catchId]['limit']:
                        # 如果有课程可抢，马上抢课
                        if(0 == int(self.listCatchLessons[catchId]['replaceId'])):
                            # 直接抢课
                            try:
                                html_str = self.postData(urlAddRequest, data={"studentAssoc": studentId, "lessonAssoc": str(
                                    catchId), "courseSelectTurnAssoc": turnId, "scheduleGroupAssoc": "", "virtualCost": "0"})
                                ret = self.postData(urlAddDropResponse, data={
                                                    "studentId": studentId, "requestId": html_str})
                                dictTemp = json.loads(ret)
                                if dictTemp["success"] == True:
                                    print(time.strftime("%Y-%m-%d_%H:%M:%S",time.localtime()))
                                    print("id: %s 抢课成功" % str(catchId))
                                    del self.listCatchLessons[catchId]
                                else:
                                    print("id: %s 抢课失败" % catchId)
                                    print("%s" %
                                            dictTemp["errorMessage"]["textZh"])
                            except TypeError:
                                continue
                            continue
                        else:
                            # 先退课
                            html_str = self.postData(urlDropRequest, data={"studentAssoc": studentId, "lessonAssoc": str(
                                self.listCatchLessons[catchId]['replaceId']), "courseSelectTurnAssoc": turnId, "scheduleGroupAssoc": "", "virtualCost": "0"})
                            ret = self.postData(urlAddDropResponse, data={
                                                "studentId": studentId, "requestId": html_str})
                            dictTemp = json.loads(ret)
                            try:
                                if dictTemp["success"] == True:
                                    print("id: %s 退课成功" % str(
                                        self.listCatchLessons[catchId]['replaceId']))
                                    # 马上开始抢课
                                    html_str = self.postData(urlAddRequest, data={"studentAssoc": studentId, "lessonAssoc": str(
                                        catchId), "courseSelectTurnAssoc": turnId, "scheduleGroupAssoc": "", "virtualCost": "0"})
                                    ret = self.postData(urlAddDropResponse, data={
                                                        "studentId": studentId, "requestId": html_str})
                                    dictTemp = json.loads(ret)
                                    if dictTemp["success"] == True:
                                        print("id: %s 抢课成功" % str(catchId))
                                        del self.listCatchLessons[catchId]
                                    else:
                                        print("id: %s 抢课失败" % catchId)
                                        print("%s" %
                                            dictTemp["errorMessage"]["textZh"])
                                        # 马上把replaceId抢回
                                        html_str = self.postData(urlAddRequest, data={"studentAssoc": studentId, "lessonAssoc": str(
                                            self.listCatchLessons[catchId]['replaceId']), "courseSelectTurnAssoc": turnId, "scheduleGroupAssoc": "", "virtualCost": "0"})
                                        ret = self.postData(urlAddDropResponse, data={
                                                            "studentId": studentId, "requestId": html_str})
                                        dictTemp = json.loads(ret)
                                        if dictTemp["success"] == True:
                                            print("id: %s 还课成功" % str(
                                                self.listCatchLessons[catchId]['replaceId']))
                                        else:
                                            print("id: %s 还课失败" % str(
                                                self.listCatchLessons[catchId]['replaceId']))
                                    continue
                                else:
                                    print("id: %s 退课失败" % str(
                                        self.listCatchLessons[catchId]['replaceId']))
                                    print("%s" %
                                        dictTemp["errorMessage"]["textZh"])
                                    continue
                            except TypeError:
                                continue
                time.sleep(0.1)

    def saveInfos(self, listLessons, path):
        """
        将infos列表中的数据存入xlsx
        """
        wb = Workbook(path)
        ws = wb.add_worksheet()
        # 定义行列
        row = 0
        col = 0
        # 第一行写入列标题
        col_header_list = ["编号", "课名", "授课老师",
                           "课程类型", "限制人数", "选课人数", "详情", "授课语言", ]
        for col_header in col_header_list:
            col = col_header_list.index(col_header)
            ws.write(row, col, col_header)
        # 接下来写入infos中值的内容
        row = 1
        for info in listLessons:
            for _key, _value in info.items():
                col = col_header_list.index(_key)
                ws.write(row, col, _value)
            row += 1
        # 关闭文件
        wb.close()

        # with open("软件学院苏州选课.txt", "w", encoding="utf-8") as f:
        #     f.write(json.dumps(listLessons, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    spider = UstcSpider()
    spider.run()
