import re
import requests
from lxml import etree

class Tuitui :
    def __init__(self) :
        self.url = "https://beijing.tuitui99.com/grsale/p1.html&sp=1123"
        self.base_url = "https://beijing.tuitui99.com"
        self.headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.139 Safari/537.36"}
        self.success_num = 0
        self.failed_num = 0

    def get_url_lists(self) :
        lists = []
        response = requests.get(self.url, headers = self.headers)
        html = response.content.decode("utf-8")
        tree = etree.HTML(html)
        nodes = tree.xpath("//a[@class='subTitle']/@href")
        for node in nodes:
            if (re.match('/grsaleInfo', node)) :
                lists.append(self.base_url + node)
        print("获取到{}条房源信息".format(len(lists)))
        return lists

    def get_data(self, url):
        data = {}
        try :
            response = requests.get(url, headers = self.headers)
            html = response.content.decode("utf-8")
            tree = etree.HTML(html)
            data["标题"] = tree.xpath("//h2[@class='xq_title']")[0].text
            data["总价"] = tree.xpath("//div[@class='basin_info fl']/ul/li[1]/div[1]/span/b")[0].text + "万"
            data["参考首付"] = tree.xpath("//div[@class='basin_info fl']/ul/li[2]/div[1]/span")[0].text
            data["参考月供"] = tree.xpath("//div[@class='basin_info fl']/ul/li[2]/div[2]/span")[0].text
            data["户型"] = tree.xpath("//div[@class='basin_info fl']/ul/li[3]/div[1]/span")[0].text
            data["建筑面积"] = tree.xpath("//div[@class='basin_info fl']/ul/li[3]/div[2]/span")[0].text
            if len(tree.xpath("//div[@class='basin_info fl']/ul/li")) == 10 :
                data["年代"] = tree.xpath("//div[@class='basin_info fl']/ul/li[4]/div[1]/span")[0].text
                data["住宅类别"] = tree.xpath("//div[@class='basin_info fl']/ul/li[4]/div[2]/span")[0].text
                data["楼层"] = tree.xpath("//div[@class='basin_info fl']/ul/li[5]/div[1]/span")[0].text
                data["朝向"] = tree.xpath("//div[@class='basin_info fl']/ul/li[5]/div[2]/span")[0].text
                data["装修"] = tree.xpath("//div[@class='basin_info fl']/ul/li[6]/div[1]/span")[0].text
                data["小区名称"] = tree.xpath("//div[@class='basin_info fl']/ul/li[7]/div[1]/span")[0].text
                data["配套设施"] = tree.xpath("//div[@class='basin_info fl']/ul/li[8]/div[1]/span")[0].text
                data["联系人"] = tree.xpath("//div[@class='basin_info fl']/ul/li[9]/a/span")[0].text
                data["电话"] = self.base_url + tree.xpath("//div[@class='basin_info fl']/ul/li[9]/a/img/@src")[0]
            elif len(tree.xpath("//div[@class='basin_info fl']/ul/li")) == 8 :
                data["年代"] = ""
                data["住宅类别"] = tree.xpath("//div[@class='basin_info fl']/ul/li[4]/div[2]/span")[0].text
                data["楼层"] = tree.xpath("//div[@class='basin_info fl']/ul/li[4]/div[1]/span")[0].text
                if len(tree.xpath("//div[@class='basin_info fl']/ul/li[5]")) == 2:
                    data["朝向"] = tree.xpath("//div[@class='basin_info fl']/ul/li[5]/div[2]/span")[0].text
                else :
                    data["朝向"] = ""
                data["装修"] = tree.xpath("//div[@class='basin_info fl']/ul/li[5]/div[1]/span")[0].text
                data["小区名称"] = tree.xpath("//div[@class='basin_info fl']/ul/li[6]/div[1]/span")[0].text
                data["配套设施"] = ""
                data["联系人"] = tree.xpath("//div[@class='basin_info fl']/ul/li[7]/a/span")[0].text
                data["电话"] = self.base_url + tree.xpath("//div[@class='basin_info fl']/ul/li[7]/a/img/@src")[0]
            if data == {} :
                self.failed_num += 1
                print("第{}条房源信息获取失败".format(self.failed_num))
            else :
                self.success_num += 1
                print("第{}条房源信息获取成功".format(self.success_num))
        except Exception as e :
            print(e)
            self.failed_num += 1
            print("第{}条房源信息获取失败".format(self.failed_num))
        return data

    def export_html(self, data) :
        print("正在导出HTML")
        html = "<html><body><table><tr>"
        for key in data[0].keys() :
            html += '<th>' + key + '</th>'
        html += "</tr>"
        for value in data :
            html += "<tr>"
            print(value.values())
            for v in value.values() :
                if v is None :
                    v = ""
                if (re.match(self.base_url, v)):
                    v = "<img src=" + v + ">"
                html += "<td>" + v + "</td>"
            html += "</tr>"
        html += "</table></body></html>"
        file = open("tuitui99.html", "w", encoding = "utf-8")
        file.write(html)
        file.close()
        print("导出HTML成功")

    def main(self):
        data = []
        lists = self.get_url_lists()
        for url in lists :
            result = self.get_data(url)
            if result != {} :
                data.append(result)

        self.export_html(data)

text = Tuitui()
text.main()