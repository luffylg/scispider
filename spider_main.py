import re

import requests
import xlsReader
from bs4 import BeautifulSoup



def crawaddress(s,link):

    r=s.get("https://apps.webofknowledge.com"+link)
    fout = open('output2.html', 'w',encoding="UTF-8")
    fout.write(r.text)
    soup=BeautifulSoup(open('output2.html','r',encoding="UTF-8"),'html.parser')
    # 获取所有地址
    addresses=[addresslable.a.text.strip() for addresslable in soup.find_all("td",class_="fr_address_row2") if addresslable.a is not None]
    # 获取所有邮箱
    emails=[alables.text for alables in soup.find_all("span",class_="FR_label",text='E-mail Addresses:')[0].parent.find_all("a")]

    return addresses,emails

class SpiderMain(object):
    def __init__(self, sid, name,year):
        self.hearders={
            'Origin':'https://apps.webofknowledge.com',
            'Referer':'https://apps.webofknowledge.com/UA_GeneralSearch_input.do?product=UA&search_mode=GeneralSearch&SID=R1ZsJrXOFAcTqsL6uqh&preferencesSaved=',
            'User-Agent':"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.94 Safari/537.36",
            'Content-Type':'application/x-www-form-urlencoded'
        }
        self.form_data={
            'fieldCount':1,
            'action':'search',
            'product':'WOS',
            'search_mode':'GeneralSearch',
            'SID':sid,
            'max_field_count':25,
            'formUpdated':'true',
            'value(input1)':name,
            'value(select1)':'AU',
            'value(hidInput1)':'',
            'limitStatus':'collapsed',
            'ss_lemmatization':'On',
            'ss_spellchecking':'Suggest',
            'SinceLastVisit_UTC':'',
            'period':'Year Range',
            'range':'ALL',
            'startYear':str(year),
            'endYear':str(year),
            'update_back2search_link_param':'yes',
            'ssStatus':'display:none',
            'ss_showsuggestions':'ON',
            'ss_query_language':'auto',
            'ss_numDefaultGeneralSearchFields':1,
            'rs_sort_by':'PY.D;LD.D;SO.A;VL.D;PG.A;AU.A'
        }
        self.form_data2={
            'product':'WOS',
            'prev_search_mode':'CombineSearches',
            'search_mode':'CombineSearches',
            'SID':sid,
            'action':'remove',
            'goToPageLoc':'SearchHistoryTableBanner',
            'currUrl':'https://apps.webofknowledge.com/WOS_CombineSearches_input.do?SID='+sid+'&product=WOS&search_mode=CombineSearches',
            'x':48,
            'y':9,
            'dSet':1
        }

    def craw(self, root_url):
        try:
            s = requests.Session()
            r = s.post(root_url,data=self.form_data,headers=self.hearders)
            soup = BeautifulSoup(r.text, 'html.parser')
            result_article=soup.find_all('input', value="DocumentType_ARTICLE")
            if result_article==[]:
                article_num=0
            else:
                article_num=int(re.findall(r"\d+",result_article[0].text.replace(',',''))[0])
            result_review=soup.find_all('input', value="DocumentType_REVIEW")
            if result_review==[]:
                review_num=0
            else:
                review_num=int(re.findall(r"\d+",result_review[0].text.replace(',',''))[0])
            a_and_r=article_num+review_num
            report_link=soup.find('a', alt="View Citation Report")
            true_link="https://apps.webofknowledge.com"+report_link['href']
            r2=s.get(true_link)
            soup2= BeautifulSoup(r2.text, 'html.parser')
            refer=soup2.find_all('span',id="CR_HEADER_3")
            refer_num=int(re.findall(r"\d+",refer[0].text)[0])
            flag=0
            error='no error'
        except Exception as e:
            # 出现错误，再次try，以提高结果成功率
            try:
                s = requests.Session()
                r = s.post(root_url,data=self.form_data,headers=self.hearders)
                soup = BeautifulSoup(r.text, 'html.parser')
                result_article=soup.find_all('input', value="DocumentType_ARTICLE")
                if result_article==[]:
                    article_num=0
                else:
                    article_num=int(re.findall(r"\d+",result_article[0].text.replace(',',''))[0])
                result_review=soup.find_all('input', value="DocumentType_REVIEW")
                if result_review==[]:
                    review_num=0
                else:
                    review_num=int(re.findall(r"\d+",result_review[0].text.replace(',',''))[0])
                a_and_r=article_num+review_num
                report_link=soup.find('a', alt="View Citation Report")
                true_link="https://apps.webofknowledge.com"+report_link['href']
                r2=s.get(true_link)
                soup2= BeautifulSoup(r2.text, 'html.parser')
                refer=soup2.find_all('span',id="CR_HEADER_3")
                refer_num=int(re.findall(r"\d+",refer[0].text)[0])
                flag=0
                error='no error'

            except Exception as e:
                print(e)
                a_and_r=0
                refer_num=0
                flag=1
                error=str(e)
        return a_and_r, refer_num,flag,error
    def craw2(self,root_url,month):
        try:
            articles=[]
            article={}
            s = requests.Session()
            r = s.post(root_url,data=self.form_data,headers=self.hearders)
            fout = open('output.html', 'w',encoding="UTF-8")
            fout.write(r.text)
            soup=BeautifulSoup(open('output.html','r',encoding="UTF-8"),'html.parser')
            result_articles=soup.find_all('div', class_="search-results")
            if len(result_articles)!=0:
                results=result_articles[0].find_all('div',class_="search-results-item")
                for result in results:
                    article={}
                    date=result.find_all('span',class_="data_bold")[-1]

                    # 解析固定月份的文献
                    if month in date.value.text:
                        article["title"]=result.find_all('a',class_="smallV110")[0].value.text
                        article["journal"]=result.find_all('span',id=re.compile(r'show.+'))[0].a.value.text
                        article["addresses"],article["emails"]=crawaddress(s, result.find_all('a',class_="smallV110")[0]["href"])
                        articles.append(article)
            flag=0
            error='no error'
        except Exception as e:
            # 出现错误，再次try，以提高结果成功率
            try:
                articles=[]
                article={}
                s = requests.Session()
                r = s.post(root_url,data=self.form_data,headers=self.hearders)
                fout = open('output.html', 'w',encoding="UTF-8")
                fout.write(r.text)
                soup=BeautifulSoup(open('output.html','r',encoding="UTF-8"),'html.parser')
                result_articles=soup.find_all('div', class_="search-results")
                if len(result_articles)!=0:
                    results=result_articles[0].find_all('div',class_="search-results-item")
                    for result in results:
                        article={}
                        date=result.find_all('span',class_="data_bold")[-1]

                        # 解析固定月份的文献
                        if month in date.value.text:
                            article["title"]=result.find_all('a',class_="smallV110")[0].value.text
                            article["journal"]=result.find_all('span',id=re.compile(r'show.+'))[0].a.value.text
                            article["addresses"],article["emails"]=crawaddress(s, result.find_all('a',class_="smallV110")[0]["href"])
                            articles.append(article)
                flag=0
                error='no error'

            except Exception as e:
                print(e)
                a_and_r=0
                refer_num=0
                flag=1
                error=str(e)
        return articles, flag,error





    def delete_history(self):
        murl='https://apps.webofknowledge.com/WOS_CombineSearches.do'
        s = requests.Session()
        s.post(murl,data=self.form_data2,headers=self.hearders)


if __name__=="__main__":
    root_url = 'https://apps.webofknowledge.com/WOS_GeneralSearch.do'
    #sid='W1OZtZW2eSwnTmvSLev'

    root='http://www.webofknowledge.com/'
    s=requests.get(root)
    sid=re.findall(r'SID=\w+&',s.url)[0].replace('SID=','').replace('&','')
    mon={1:"JAN",2:"FEB",3:"MAR",4:"APR",5:"MAY",6:"JUN",7:"JUL",8:"AUG",9:"SEP",10:"OCT",11:"NOV",12:"DEC"}

    # 月份和年份在此修改
    month=mon[7]
    year=2016
    datalist=xlsReader.read()
    for index,data in enumerate(datalist):
        if index%100==0:
            # 每一百次更换sid
            s=requests.get(root)
            sid=re.findall(r'SID=\w+&',s.url)[0].replace('SID=','').replace('&','')
        csv=open('res.csv','a')
        fail=open('fail.txt','a')
        obj_spider = SpiderMain(sid,data["name"],year)
        articles, flag,error=obj_spider.craw2(root_url,month)

        if flag==0:
            print('\ncell '+str(index+1)+' '+data["name"]+' success:')
            if not articles:
                print("no result found")
            for article in articles:
                if data["email"] in article["emails"]:
                    # 如果excel读取的邮件在爬到的邮件中，直接输出title和journal
                    print("title: "+article["title"]+"\njournal: "+article["journal"])
                else:
                    print("title: "+article["title"]+"\njournal: "+article["journal"])
                    print(("\n").join(article["addresses"]))

        else:
            print('cell'+str(index+1)+' '+data["name"]+' failed'+' '+error)
            fail.write(str(index+1)+' '+data["name"]+' failed'+' '+error+'\n')

        csv.close()
        fail.close()


    # for i in range(1,nrows):
    #     if i%100==0:
    #         # 每一百次更换sid
    #         s=requests.get(root)
    #         sid=re.findall(r'SID=\w+&',s.url)[0].replace('SID=','').replace('&','')
    #     csv=open('res.csv','a')
    #     fail=open('fail.txt','a')
    #     kanming=table.cell(i,2).value
    #     obj_spider = SpiderMain(sid,kanming)
    #     ar,ref,fl,er=obj_spider.craw(root_url)
    #     csv.write(str(i+1)+"["+kanming+'['+str(ar)+'['+str(ref)+'\n')
    #     if fl==0:
    #         print('cell'+str(i+1)+' '+kanming+' finished')
    #     else:
    #         print('cell'+str(i+1)+' '+kanming+' failed'+' '+er)
    #         fail.write(str(i+1)+' '+kanming+' failed'+' '+er+'\n')
    #     csv.close()
    #     fail.close()
    #     #obj_spider.delete_history()
    #     # 更换sid后不必删除历史记录
