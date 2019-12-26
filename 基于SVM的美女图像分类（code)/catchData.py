import requests
import urllib.request as rq
import os
import re
import xlrd
import xlutils.copy
from bs4 import BeautifulSoup


# 获取html文本
def getHtml(url):
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_5) AppleWebKit 537.36 (KHTML, like Gecko) Chrome"
    }
    req = session.get(url, headers=headers)
    html = req.text
    return html


# 图片url列表
def getImg(html):
    Null_img_list = []
    reg = r'img src="([.*\S]*\.jpg)"'  # 正则表达式
    imgre = re.compile(reg)
    imglist1 = re.findall(imgre, html)
    if len(imglist1) > 2:
        return imglist1
    else:
        return Null_img_list


# 获取女生姓名
def getName(html):
    reg = r'<strong>今日女神：(.*?)</strong>'
    textRe = re.compile(reg)
    textList = re.findall(textRe, html)
    return textList


# 获取本期的女神的名字和期数
def getNumb_name(number, html):
    GoddessName = getName(html)
    if GoddessName == []:
        print("女神大会" + number + ":" + "空")
    else:
        print("女神大会" + number + ":" + GoddessName[0])


# 获取期刊标题
def getTitle(html):
    soup = BeautifulSoup(html, "html.parser")
    titleList = soup.find_all('a', {'target': '_self', 'content': ''})
    titleListLength = len(titleList)
    if titleListLength > 1:
        title = str(soup.find_all('a', {'target': '_self', 'content': ''})[1].string)
    elif titleListLength == 1:
        title = str(soup.find_all('a', {'target': '_self', 'content': ''})[0].string)
    else:
        title = "没有获取到该期刊的标题"
    return title


# 获取女神评分
def getGrade(html):
    NullGrade = ["没有获取到评分"]
    reg = r'<span style="color:#ff0000">综合得分(.*?)</span>'
    GradeRe = re.compile(reg)
    GradeList1 = re.findall(GradeRe, html)
    reg = r'<strong>综合得分(.*?)</strong>'
    GradeRe = re.compile(reg)
    GradeList2 = re.findall(GradeRe, html)
    GradeListAll = GradeList1 + GradeList2
    if len(GradeListAll) > 0:
        return  GradeListAll
    else:
        return  NullGrade


# 女神链接列表
def getLinkList(html):
    reg = r'<h3><a href="(.*?)" target="_blank">'
    linkRe = re.compile(reg)
    linkList = re.findall(linkRe, html)
    newLinkList = []
    for i in linkList:
        Goddess_Url = "https://www.dongqiudi.com" + i
        newLinkList.append(Goddess_Url)
    return newLinkList


# 获取上期女神的链接
def getLastGoddessUrl(html):
    soup = BeautifulSoup(html, "html.parser")
    linkList = soup.find_all('a', {'target': '_self', 'content': ''})
    if len(linkList) > 1:
        string = str(soup.find_all('a', {'target': '_self', 'content': ''})[1].get('href'))
        str_link = string[18::]
        Goddess_link = "https://www.dongqiudi.com/archive/" + str_link + ".html"
        LastGoddessUrl = Goddess_link
    else:
        LastGoddessUrl = "获取链接失败"
    return LastGoddessUrl


# 通过遍历主页链接列表，进入子链接，获取女生姓名、图片地址、还有大众评分
def getGoddessInfo_dict(home_Goddess_link):
    length = len(home_Goddess_link)
    dict = {}
    for i in range(1, length - 1):
        url = home_Goddess_link[i]
        html = getHtml(url)
        Goddess_name = getTitle(html)
        Goddess_grade = getGrade(html)[0]
        Goddess_link = getLastGoddessUrl(html)
        if Goddess_link != "获取链接失败":
            Goddess_html = getHtml(Goddess_link)
            # 默认取图片的第三张
            if getImg(Goddess_html) != []:
                Goddess_img_url = getImg(Goddess_html)[3]
            else:
                Goddess_img_url = "无法获取到地址链接"
            # 用字典保存每个女神的信息
            dict[i] = {"Goddess_name":Goddess_name,"imgUrl":Goddess_img_url,"grade":Goddess_grade}
        else:
            dict[i] = {"Goddess_name": Goddess_name, "imgUrl":"无法获取到地址链接", "grade": Goddess_grade}
    return dict

def insert_Excel(filePath, data_dict):
    # filePath:是Excel表的路径，data_dict:你所要插入的数据,数据用字典的形式存放
    book = xlrd.open_workbook(filePath, formatting_info=True)
    wtbook = xlutils.copy.copy(book)
    wtsheet = wtbook.get_sheet(0)
    tableHead = ["序号","女神描述","图片地址","综合评分"]
    for i in range(len(tableHead)):
        wtsheet.write(0,i,tableHead[i])
    # 获取字典的长度
    dictCount = len(data_dict)
    # 获取一个字典中的所有键，用于遍历
    keysList = list(data_dict[1].keys())
    # 字典1中的键长度
    keysLength = len(keysList)
    # 计数
    count = 1
    # 行数
    row = dictCount
    # 列数
    col = keysLength
    # 用来标记从第二行开始插值
    k = 1
    # 开始插值
    for i in range(row):
        wtsheet.write(k, 0, count)
        for j in range(col):
            t = j + 1
            wtsheet.write(k, t, data_dict[k][keysList[j]])
        k = k + 1
        count = count + 1
    wtbook.save(filePath)


if __name__ == "__main__":
    url = "https://www.dongqiudi.com/special/375"
    html = getHtml(url)
    # 获取主页每期女神的杂志链接
    home_Goddess_link = getLinkList(html)
    # 获取女神的信息
    dict = getGoddessInfo_dict(home_Goddess_link)
    # 将数据插入到Excel中
    filePath = 'GoddessList.xls'
    data_dict = dict
    insert_Excel(filePath,data_dict)
    print("插入结束")
