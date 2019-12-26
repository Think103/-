import urllib.request as rq
import xlrd
# 从Excel中获取到图片链接，将所有链接存放在一个列表里
def getImgList(filePath):
    book = xlrd.open_workbook(filePath)
    sheet = book.sheet_by_index(0)
    imgList = sheet.col_values(2)
    return  imgList

# 将评分大于9.0的称为女神级别,大于8.5的为大众级别，小于8.5的为普通级别，获取到她们的图片地址列表
def classify_grade_url(filePath):
    Goddess_url = []
    public_url = []
    commom_url = []
    imgList = getImgList(filePath)
    book = xlrd.open_workbook(filePath)
    sheet = book.sheet_by_index(0)
    gradeList = sheet.col_values(3)
    gradeList_Lenth = len(gradeList)
    for i in range(1,gradeList_Lenth):
        if float(gradeList[i]) >= 9.0:
            Goddess_url.append(imgList[i])
        elif float(gradeList[i]) >= 8.5:
            public_url.append(imgList[i])
        else:
            commom_url.append((imgList[i]))
    return Goddess_url,public_url,commom_url
# 下载图片函数
def download_img(destPath,imgUrl_list):
    imgName = 1
    for i in range(0,len(imgUrl_list)):
        f = open(destPath+"/"+str(imgName)+".jpg", 'wb')
        f.write(rq.urlopen(imgUrl_list[i]).read())
        f.close()
        print("正在下载第%s 张图片"%imgName)
        imgName += 1
    print("该图片下载已经完成")




if __name__ == "__main__":
    filePath = "Modify_GoddessList.xls"
    Goddess_url, public_url, common_url = classify_grade_url(filePath)
    # 下载图片
    destPath1 = "女神级别"
    destPath2 = "大众级别"
    destPath3 = "普通级别"
    download_img(destPath1,Goddess_url)
    download_img(destPath2, public_url)
    download_img(destPath3, common_url)