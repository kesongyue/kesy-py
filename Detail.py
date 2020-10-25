import requests
from lxml import etree
import xlsxwriter
import re
import os

# 需修改为自己的agent以及配公司的代理
user_agent2= 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.18362'
user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Safari/537.36'
headers = {
    'User-Agent':user_agent,
    'Connection': 'keep-alive',
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'}
def get_html_content(url):
    try:
        html = requests.get(url,headers = headers)
        html.encoding = "utf-8"
        return html.content
    except:
        print("Error:get html failed")
        return ""

def parse_page(html):
    resultInfo = {}
    htmlInfo = etree.HTML(html)
    # print(etree.tostring(htmlInfo).decode())
    # f = open('html.html','w+')
    # f.write(etree.tostring(htmlInfo).decode())
    # print("".join(htmlInfo.xpath('//span[@id="productTitle"]')))
    resultInfo["productTitle"] = htmlInfo.xpath('//span[@id="productTitle"]')[0].text.strip()
    resultInfo["availability"] = htmlInfo.xpath('//div[@id="availability"]/span')[0].text.strip()
    resultInfo["merchant-info"] = htmlInfo.xpath('//div[@id="merchant-info"]')[0].text.strip()
    sellerProfileTriggerId = htmlInfo.xpath('//div[@id="merchant-info"]/a[@id="sellerProfileTriggerId"]')
    if sellerProfileTriggerId:
        resultInfo["merchant-info"] += " " + sellerProfileTriggerId[0].text.strip()
    return resultInfo

# def getInfo(row):
#     arr = [{"productTitle":"P30","availability":"In stock","merchant-info":"Dispatched from and sold by Amazon."},
#         {"productTitle":"P40","availability":"In stock","merchant-info":"Dispatched from and sold by others."},
#         {"productTitle":"P50","availability":"","merchant-info":""},
#         {"productTitle":"P60","availability":"Only 1 left in stock.","merchant-info":""},
#         {"productTitle":"P70","availability":"Currently unavailable.","merchant-info":""},
#         {"productTitle":"P80","availability":"Temo out of stock","merchant-info":""}]
#     return arr[row-1]

def writeToFile(worksheet,row,col,resultInfo,cell_color):
    worksheet.write(row, col, resultInfo["labelNumber"], cell_color)
    worksheet.write(row, col + 1, resultInfo["productTitle"], cell_color)
    worksheet.write(row, col + 2, resultInfo["availability"], cell_color)
    worksheet.write(row, col + 3, resultInfo["merchant-info"], cell_color)

# html = get_html_content("https://www.amazon.co.uk/HUAWEI-Smartphone-SuperCharge-SIM-Free-Android-Black/dp/B086FCRDXY/ref=sr_1_1?dchild=1&keywords=B086FCRDXY&qid=1603002842&sr=8-1")
# parse_page(html)

url_list = []
# 打开网址存储文件
with open("url.txt","r+") as f:
    lines = f.readlines()
    for line in lines:
        url_list.append(line.strip())

print("Finish reading url.txt---\nStarting analyzing")
workbook = xlsxwriter.Workbook("result.xlsx")
worksheet = workbook.add_worksheet()
cell_red = workbook.add_format({'color':'red'})
cell_blue = workbook.add_format({'color':'blue'})
cell_green = workbook.add_format({'color':'green'})
cell_normal = workbook.add_format()
result_red = []
result_blue = []
result_green = []
result_normal = []
string_normal = "Dispatched from and sold by Amazon"
string_others = "sold by"
string_unavailable = "Currently unavailable"
string_out_of_stock = "out of stock"

worksheet.write(0,0,"标签")
worksheet.write(0,1,"productTitle")
worksheet.write(0,2,"availability")
worksheet.write(0,3,"merchant-info")
row = 1
col = 0

for url in url_list:
    html = get_html_content(url)
    if len(html) == 0:
        print("Error:getting url:" + url + "failed")
        break
    resultInfo = parse_page(html)
    labelNumber = re.findall('dp/[0-9A-Za-z]+[/|?]',url)[0][3:-1]
    print(str(row) + ": Finish Analyzing " + labelNumber + " Start to write to file")
    resultInfo["labelNumber"] = labelNumber
    if re.match(string_normal,resultInfo["merchant-info"]):
        cell_color = cell_normal
        result_normal.append(resultInfo)
    elif re.search(string_others,resultInfo["merchant-info"],flags=re.IGNORECASE):
        cell_color = cell_red
        result_red.append(resultInfo)
    elif re.search(string_unavailable, resultInfo["availability"]) or (len(resultInfo["availability"]) == 0 and len(resultInfo["merchant-info"]) == 0):
        cell_color = cell_blue
        result_blue.append(resultInfo)
    elif re.search(string_out_of_stock, resultInfo["availability"]):
        cell_color = cell_green
        result_green.append(resultInfo)
    else:
        cell_color = cell_normal
        result_normal.append(resultInfo)
    row+=1

row = 1
for result_info in result_red:
    writeToFile(worksheet,row,0,result_info,cell_red)
    row += 1
for result_info in result_blue:
    writeToFile(worksheet,row,0,result_info,cell_blue)
    row += 1
for result_info in result_green:
    writeToFile(worksheet,row,0,result_info,cell_green)
    row += 1
for result_info in result_normal:
    writeToFile(worksheet,row,0,result_info,cell_normal)
    row += 1

workbook.close()
print("Write to result.xlsx successfully")
os.system("pause")