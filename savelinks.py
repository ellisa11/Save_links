#!/usr/bin/python
# _*_ coding: utf-8 _*_


import httplib,urllib,cookielib,urllib2
import xlrd
from xlutils.copy import copy
import os


class Save():


    def checkurlinxls(self):

       #读xls
        exl=xlrd.open_workbook(os.curdir+'/testlink.xlsx')
        table=exl.sheet_by_name('Sheet1')

        wb=copy(exl)
        ws=wb.get_sheet(0)

        nrows=table.nrows

        for i in range(nrows):
            #取第二列的链接内容
            linkurl=table.cell(i,1).value
            #去掉标题栏
            if len(linkurl)>10:
                #调用接口运行
                #result=self.httppost(linkurl)

                #post start
                host = '**.**dao.com'
                url = 'http://**.**dao.com/yws/open/memory?method=put'

                post_data = {'text': linkurl}

                post_data_urlencode = urllib.urlencode(post_data)

                #120
                head = {
                    'cookie': 'YNOTE_CSTK=-qX6MJnk; YNOTE_LOGIN=3||1490843959689; YNOTE_SESS=v2|9h7j7PMYcVQ4n4QFkLeB0gFRLkE64Jy0eyPLOEPMeLRlWnf6Zh4eS06LhLwF0MeZ0PKOMU50HwZ0eBOfg4kfwLRYG0fPKOfqz0; JSESSIONID=aaa2-4DXGXfYe4BXgYhQv',
                    'Content-Type': 'application/x-www-form-urlencoded'}

                try:
                    httpClient = httplib.HTTPConnection(host, timeout=30)

                    httpClient.request("POST", url, headers=head, body=post_data_urlencode)

                    res = httpClient.getresponse()

                    # 处理结果写入xls
                    if res.status != 200:
                        # 将状态码写入第四列中
                        ws.write(i, 3, res.status)

                    else:
                        # 将结果写入第四列中
                        s= res.read()[11:len(res.read())-1]
                        ws.write(i, 3, s)

                except Exception, e:
                    print e
                finally:
                    if httpClient:
                        httpClient.close()
                # post end

            else:
                # 是xls的标题，就什么也不做
                pass
        # 保存结果
        wb.save(os.curdir+'/testlinkresult.xls')


    # 发送请求并处理
    def httppost(self,texturl):

        host = '**.youdao.com'
        url = 'http://**/yws/open/memory?method=put'

        post_data={'text': texturl}

        post_data_urlencode = urllib.urlencode(post_data)

        head = {
            'cookie': 'YNOTE_CSTK=ml_bnH7N; YNOTE_LOGIN=3||1490581245719; YNOTE_SESS=v2|2mOPuTI-RVU5k4kM0LpLRT4kLYf64gu0qFRHTy0MPBRgLRfT40MOf0gBRLzYhHzE0JyhfkWnfgZ0JyOLOEO4kA0UW64eFRfzW0; JSESSIONID=aaaHmOuqB6Yjb_4LnFhQv',
            'Content-Type':'application/x-www-form-urlencoded'}

        try:



            httpClient = httplib.HTTPConnection(host, timeout=30)

            request=httpClient.request("POST", url, headers=head,body=post_data_urlencode)

            res = httpClient.getresponse()



        except Exception, e:
            print e
        finally:
            if httpClient:
                httpClient.close()

        return res





oo=Save()
oo.checkurlinxls()



















