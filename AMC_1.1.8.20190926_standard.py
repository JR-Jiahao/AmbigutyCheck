#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
@Time    : 2019/08/19
@Author  : JR
@FileName: AmbiguityCheck
@Software: PyCharm
"""
import os
import xlrd
import time
import smtplib
import getpass
from colorprint import *
from selenium import webdriver
from email.mime.text import MIMEText
from email.header import Header
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
start = time.clock()  # 获取开始时间
# 获取login, 作为邮箱地址前缀
get_email_address = raw_input("Please enter your login:")
# 获取登录Testrial的密码
get_login_password = getpass.getpass("Please enter your password:")
save_log = "C:/Users/"+get_email_address+"/Documents/AmbiguityCheckLog.txt"
print "Log will be saved at location: "+save_log
output_log = open(save_log, 'w')
# 获取存放TID信息的Excel表格路径，用于读取其中的TID
gt_excel_location = raw_input("Please drag your excel file to here:")

DA_name = ["Tested by Xueshan W.", "Tested by Zhou, K.", "Tested by Xudong L.",
           "Tested by Xuewei Q.", "Tested by Ruirui M.", "Tested by Tian Y.",
           "Tested by lin l.", "Tested by LiHui","Tested by Jin G.",
           "Marked by Zhou, K.", "Marked by Xudong L.", "Marked by Jin G.",
           "Marked by Xuewei Q.", "Marked by Ruirui M.", "Marked by Tian Y.",
           "Marked by lin l.", "Marked by LiHui", "Marked by Xueshan W."]

'''在使用chromdriver的时候，页面会打开运行，便于调试的时候观察，
   当程序调试稳定后，我们可以隐藏掉页面，让程序在后台静默运行。
'''
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
driver = webdriver.Chrome(chrome_options=chrome_options)

td = []          # 定义一个数组testdate,简称td,用于存放获取的，已处理过的测试日期
te = []          # 定义一个数组tester,简称te,用于存放获取的 DA NAME
cid = []         # 定义一个数组cid, 用于存放通过TID 获取到的 CID信息
excel_data = []  # 定义一个数组，用于存放读取的Excel中的TID信息
date_time = []   # 定义一个数组，用于存放获取的，未处理过的测试日期


# 发件人
sender = 'AmbiguityCheck@amazon.com'
# 收件人
email = get_email_address+'@amazon.com'


def ambiguity_check():
    wb = xlrd.open_workbook(filename=gt_excel_location)  # 打开TID excel文件
    wb.sheet_names()                # 获取表格名字
    sheet1 = wb.sheet_by_index(0)   # 通过索引获取表格
    cols = sheet1.col_values(0)     # 获取列内容
    print str(len(cols)) + " record were found from excel"
    for el in range(len(cols)):
        excel_data.append(cols[el])
    ed = 0
    print "Excel Data Lenth: " + str(len(excel_data))
    while ed <= len(excel_data)-1:
        print "\n"
        printBlue("------ Total "+str(len(cols))+" cases need to check, "
                  + str((len(cols) - ed)) + " More Waiting for Check." + "------")
        print "Start Checking TID: " + excel_data[ed][1:]
        # -----------------------------------------------进入TestRun页面------------------------------------------------
        print "TestRail is accessed"
        try:
            driver.get("https://testrail.labcollab.net/testrail/index.php?/auth/login/")
            print "waiting for login ... ..."
            driver.find_element_by_id("name").send_keys(get_email_address+"@amazon.com")
            driver.find_element_by_id("password").send_keys(get_login_password)
            driver.find_element_by_id("button_primary").click()
            if driver.find_element_by_id("button_primary"):
                printRed("Password does not correct !!!")
                get_password = getpass.getpass("Please enter your password:")
                driver.find_element_by_id("name").send_keys(get_email_address + "@amazon.com")
                driver.find_element_by_id("password").send_keys(get_password)
                driver.find_element_by_id("button_primary").click()
            if driver.find_element_by_id("button_primary"):
                printRed("Password incorrect again !!!")
                printRed("Please take sometime to memorize your password and rerun this program!")
                printRed("Program exit, Bye!")
                os._exit(1)
            else:
                printGreen("login success")
        except:
            print "No need login again"
        try:
            driver.get("https://testrail.labcollab.net/testrail/index.php?/tests/view/"+excel_data[ed][1:])
        except TimeoutException, e:
            print type(e)
            print "Timeout when loading TID page"
        printGreen("Access TID page success")
        time.sleep(20)
        driver.find_element_by_xpath("//*[@id='content-header']/div/div[1]/a").click()
        print "Obtaining CID ... ..."
        get_cid = driver.find_element_by_class_name("content-header-id").text
        printGreen("CID Obtain success")
        print "CID is: "+get_cid[1:]
        # ------------------通过Case ID(CID),直接进到对应的case页面,并获取当前页面所有的执行结果------------------------
        print "Accessing CID page ... ..."
        driver.get("https://testrail.labcollab.net/testrail/index.php?/cases/results/"+get_cid[1:])

        time.sleep(5)
        # 定位当前页面Table表格中的status字段，即执行结果的字段（PASSED/BLOCKED/PARKED/FAILED/...）
        print "Checking Ambiguity ... ..."
        tid_verify = []
        table_tid = driver.find_elements_by_xpath("//*[@id='tests']/table/tbody/tr/td/a[@class='link-noline']")
        for t_id in table_tid:
            tid_result = t_id.text
            tid_verify.append(tid_result)
        # 遍历table,并转成文本输出,可以直观的看到当前页面，当前case所有的执行结果
        excution_result = []  # 定义一个数组，用于存放每个CID的所有执行结果
        table = driver.find_elements_by_xpath("//*[@id='tests']/table/tbody/tr/td/span[@class='status']")
        for i in table:
            result = i.text
            excution_result.append(result)
        # 所有测试人员的Name都在table里，通过xpath定位name对应的span元素，并遍历转成text，然后存进数组里。
        table_name = driver.find_elements_by_xpath("//*[@id='tests']/table/tbody/tr/td/span[@class='text-soft']")
        name = []
        for n in table_name:
            t = n.text
            name.append(t)
        zipped_tid_name = zip(tid_verify, name)
        zipped_lenth = len(zipped_tid_name)
        zip_tid = 0
        while zip_tid in range(zipped_lenth):
            if excel_data[ed] == zipped_tid_name[zip_tid][0] and zipped_tid_name[zip_tid][1] in DA_name:
                '''将获取到的所有测试结果以及所有测试人员的Name打包到一起，
                   这样就能知道某条case的结果是谁执行的了，但当前包含的是所有人，
                   即包含其他Team的人员，还要进一步筛选。
                '''
                zipped = zip(excution_result, name)
                '''遍历刚才打包的结果，首先将Untested的case过滤掉，其次对有执行结果的case
                   进行人员判断，判断其对应的测试人员是否在DA_name[]数组里，如果在就存放
                   进定义好的tester[]数组里。
                '''
                zip_lenth = len(zipped)
                tester = []
                t = 0
                while t in range(zip_lenth):
                    if t == zip_lenth - 1:
                        break
                    if zipped[t][0] == "Untested":
                        t += 1
                    else:
                        print 'Skip Untested Case'
                    if zipped[t][1] in DA_name:
                        tester.append(zipped[t][1])
                        break
                    else:
                        t += 1
                if len(tester) == 0:
                    print "DA name is not CN DA"
                elif tester[0] in DA_name:
                    print tester[0]
                new_result_len = len(excution_result)
                # result列表下标最大值 = result列表长度 - 1
                new_len = new_result_len - 1
                lenth = t
                while lenth < new_len:
                    if excution_result[t] != "Passed":
                        break
                    if excution_result[t] == "Passed":
                        if excution_result[t] == excution_result[lenth + 1]:
                            print "Checking ... ..."
                        elif excution_result[lenth + 1] == "Failed" and tester[0] in DA_name:
                            te.append(tester[0])
                            cid.append(get_cid[1:])
                            tid_all = driver.find_elements_by_xpath(
                                "//*[@id='tests']/table/tbody/tr/td/a[@class='link-noline']")
                            tid_number = []
                            for t_id in tid_all:
                                tid = t_id.text
                                tid_number.append(tid)
                            test_id_number = str(tid_number[t][1:])
                            driver.get(
                                "https://testrail.labcollab.net/testrail/index.php?/tests/view/" + test_id_number)
                            while True:
                                try:
                                    date = driver.find_elements_by_xpath("//*[@class='change-meta']/p")
                                    for d in date:
                                        t = d.text
                                        date_time.append(t)
                                    if len(date_time) == 0:
                                        date_time.append("No Result")
                                    else:
                                        for k in range(len(date_time)):
                                            space = date_time[k].find(" ")
                                            td.append(date_time[0][0:space].rstrip())
                                        print "\n"
                                        if len(excel_data) == ed:
                                            printBlue("------ No More Checking ------")
                                            print "\n"
                                        else:
                                            printRed("------ Ambiguity Detected! ------")
                                            output_log.write(excel_data[ed][1:]+"\n")
                                            output_log.write("https://testrail.labcollab.net/testrail/index.php?/tests/view/"+get_cid[1:]+"\n")
                                    break
                                except:
                                    break
                            break
                        else:
                            print "Checking ... ..."
                        lenth += 1
                break
            else:
                print "No need to check, skip to next one."
                zip_tid += 1
        ed += 1
    zipped = zip(td, te, cid)
    if len(zipped) == 0:
        print "\n"
        printGreen('                Checking Complete !')
        printGreen('------------ No Ambiguity Triggered ! ! ! ------------')
        end = format(((time.clock() - start) / 60), '.2f')
        print "              Time used: " + str(end) + " minutes"
        print "\n\n"
    else:
        html = """
               <html>
               <p>Below is an audit finding which is considered as a miss from your end. It would be helpful if you could provide any data or method that makes this deviation as tolerable. Awaiting your inputs/confirmation which will help in tracking this to closure.</p>
               <table border="1">
                   <tr bgcolor=lightblue>
                       <th><strong>Audit Type</strong></th>
                       <th><strong>Execution Date</strong></th>
                       <th><strong>Executer</strong></th>
                       <th><strong>Interface</strong></th>
                   <tr>
               """
        for c, v, r in zipped:
            html += "<tr><td align='center'>{}</td>".format("Ambiguity")
            html += "<td align='center'>{}</td>".format(c)
            html += "<td align='center'>{}</td>".format(v)
            html += "<td><a href='https://testrail.labcollab.net/testrail/index.php?/cases/results/%s'>" \
                    "https://testrail.labcollab.net/testrail/index.php?/cases/results/%s</a></td></tr>" % (r, r)
        html += "</table><p>Regards</p></html>"
        message = MIMEText(html, 'HTML')
        message['From'] = sender
        message['To'] = email
        message['Cc'] = 'xueweqin@amazon.com'
        subject = 'Automated Ambiguity Audit Triggered Remind'
        message['Subject'] = Header(subject, 'utf-8')
        try:
            server = smtplib.SMTP('smtp.amazon.com', 25)
            server.sendmail(sender, email, message.as_string())
            print "\n\n"
            printRed("                                        Warning !")
            printRed("------------ Ambiguity has been triggered, email has been sent, please check ! ------------")
            end = format(((time.clock() - start)/60), '.2f')
            print "                                Time used: "+str(end)+" minutes"
        except smtplib.SMTPException:
            printRed("Error: Cannot sending email")


ambiguity_check()
