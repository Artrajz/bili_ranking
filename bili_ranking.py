import xlwt
from bs4 import BeautifulSoup
import re
import urllib.request
import urllib.error
import time

def main():
    baseurl = r'https://www.bilibili.com/v/popular/rank/all'
    datalist = getData(baseurl)
    saveData(datalist)
    print("爬取完毕！")

    # print(getTime() + " " + "bilibili排行榜")
    # for i in datalist:
    #     print("BV：" + i[1], end=" ")
    #     print("标题：" + i[2], end=" ")
    #     print("播放量：" + i[3], end=" ")
    #     print("评论：" + i[4], end="\n")

findLink = re.compile(r'<a class="title" href="//(.*?)"')
findLink_BV = re.compile(r'<a href="//www.bilibili.com/video/(.*?)"')
findTitle = re.compile(r'target="_blank">(.*?)</a>')
findPlay = re.compile(r'<i class="b-icon play"></i>(.*?)</span>',re.S)
findView = re.compile(r'<i class="b-icon view"></i>(.*?)</span>',re.S)

def getTime(select):
    if select == 0:
        return time.strftime("%Y.%m.%d %H:%M:%S", time.localtime())
    else:
        return time.strftime("%Y_%m_%d_%H_%M_%S", time.localtime())
    
def getData(url):
    datalist = []
    html = askUrl(url)
    soup = BeautifulSoup(html, "html.parser")
    i = 1

    for item in soup.find_all('li',class_="rank-item"):
        data = []
        item = str(item)

        link = re.findall(findLink,item)[0]
        data.append(link)

        link_BV = re.findall(findLink_BV,item)[0]
        data.append(link_BV)

        title = re.findall(findTitle,item)[1]  #第二条才是标题信息
        data.append(title)

        play = re.findall(findPlay,item)[0]
        play = re.sub(r"\n?","",play)
        play = play.strip()
        data.append(play)

        view = re.findall(findView,item)[0]
        view = re.sub(r'\n?',"",view)
        view = view.strip()
        data.append(view)

        datalist.append(data)
        # print("title="+title,end=" ")
        # print("play="+play,end=" ")
        # print("view="+view,end=" ")
        # print(link)
    return datalist

def askUrl(url):
    head = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Cache-Control": "max-age=0",
        "Connection": "keep-alive",
        "Host": "www.bilibili.com",
        "Referer": "https://www.bilibili.com/",
        "sec-ch-ua": '" Not;A Brand";v="99", "Google Chrome";v="91", "Chromium";v="91"',
        "sec-ch-ua-mobile": "?0",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.77 Safari/537.36"
    }
    request = urllib.request.Request(url, headers=head)
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8")
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html

def saveData(datalist):
    book = xlwt.Workbook(encoding="uft-8",style_compression=0)#不允许改变表格样式
    sheet = book.add_sheet("bilibili热门视频排行榜",cell_overwrite_ok=True)#允许单元格覆写
    col = ["视频链接","BV号","标题","播放量","评论"]
    sheet.write(0,0,"爬取时间")
    sheet.write(0,1,getTime(0))
    for i in range(0,len(col)):
        sheet.write(1,i,col[i])
    for i in range(0,100):
        data = datalist[i]
        for j in range(0,len(data)):
            sheet.write(i+2,j,data[j])
    bookName = "bili_ranking_"+getTime(1)+".xlsx"
    savepath = "./output/" + bookName
    book.save(savepath)



if __name__ == "__main__":
    main()
