# -*- coding: UTF-8 -*-


__author__ = 'Zhang.zhiyang'

import os
import xlrd
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
import sys
import time
reload(sys)
sys.setdefaultencoding('utf-8')
fileName = raw_input('Copy the file to this path , And input the whole fileName(Like : 1.xls) :')
filePath = os.getcwd() + '\\' + fileName
url = "https://www.hn.10086.cn/Shopping/front/yiyangZone/goal/index.jsp"
driver = webdriver.Chrome(executable_path=os.getcwd()+"\\"+"chromedriver")
failedUserList = []


# get user's info from excel
def getUserInfoList(filePath):
    data = xlrd.open_workbook(filePath)
    table = data.sheets()[0]
    nrows = table.nrows
    infoList = []

    for i in range(nrows):
        userInfo = []
        referee = int(table.cell(i, 0).value)
        userInfo.append(str(referee))
        phoneNum = int(table.cell(i, 1).value)
        userInfo.append(str(phoneNum))
        userName = table.cell(i, 2).value
        userInfo.append(userName)
        phoneModel = table.cell(i, 3).value
        userInfo.append(phoneModel)
        unusedPeo = table.cell(i, 4).value
        userInfo.append(unusedPeo)
        self = table.cell(i, 5).value
        userInfo.append(self)
        address = table.cell(i, 6).value
        userInfo.append(address)

        infoList.append(userInfo)
    # print(len(infoList))
    return infoList

# open browser edit user's info and submit
def submitInfos(driver,url,infoList):
    for i in range(len(infoList)):
        driver.get(url)
        wait = WebDriverWait(driver,60)
        phoneurl = driver.find_elements_by_class_name('red_box_name')
        num = 0

        for n in range(len(phoneurl)):
            getedName = phoneurl[n].get_attribute('title')
            # print(phoneurl[i].get_attribute('title'))
            index = getedName.find(' ')
            if(index != -1):
                phoneName = getedName[0:index]+getedName[index+1:]
            else:
                phoneName = getedName
            # print(phoneName)
            # print(infoList[i][3])
            if phoneName.upper().find(infoList[i][3].upper()) != -1:
                num += 1
                print('find it')
                thisphone = phoneurl[n].find_element_by_class_name('btn_02')
                getedurl = thisphone.get_attribute('onclick')
                thisurl = "https://www.hn.10086.cn"+getedurl[getedurl.find('(')+2:getedurl.find(')')-1]
                print(getedurl)
                print(thisurl)
                driver.get(thisurl)
                wait.until(lambda x : x.find_element_by_class_name('nav').is_displayed())
                driver.find_element_by_id('userName').send_keys(unicode(infoList[i][2]))
                driver.find_element_by_id('phoneNum').send_keys(infoList[i][1])
                driver.find_element_by_id('recommendName').send_keys(infoList[i][0])
                select = driver.find_element_by_tag_name("select")
                select.click()
                option = select.find_elements_by_tag_name('option')
                for m in range(len(option)):
                    print(option[m].text)
                    if option[m].text == infoList[i][6]:
                        option[m].click()
                        break
                driver.find_element_by_class_name('greyBtn').click()
                time.sleep(2)
                if driver.current_url == thisurl:
                    wait.until(lambda x:x.find_element_by_id('alert_popup_ok').is_displayed())
                    driver.find_element_by_id('alert_popup_ok').click()
                    failedUserList.append(unicode(infoList[i][2]))
                    print('has submited')
                else:
                    wait.until(lambda x:x.find_element_by_class_name('stepUl').is_displayed)
                    print('success')
                break

        if num == 0 :
            for n in range(len(phoneurl)):
                getedName = phoneurl[n].get_attribute('title')
                # print(phoneurl[i].get_attribute('title'))
                index = getedName.find(' ')
                if(index != -1):
                    phoneName = getedName[0:index]+getedName[index+1:]
                else:
                    phoneName = getedName
                # print(phoneName)
                # print(infoList[0][3])
                if phoneName.upper().find(infoList[i][3][0:5].upper()) != -1:
                    print('find it')
                    thisphone = phoneurl[n].find_element_by_class_name('btn_02')
                    getedurl = thisphone.get_attribute('onclick')
                    thisurl = "https://www.hn.10086.cn"+getedurl[getedurl.find('(')+2:getedurl.find(')')-1]
                    # print(getedurl)
                    print(thisurl)
                    driver.get(thisurl)
                    wait.until(lambda x : x.find_element_by_class_name('nav').is_displayed())
                    driver.find_element_by_id('userName').send_keys(unicode(infoList[i][2]))
                    driver.find_element_by_id('phoneNum').send_keys(infoList[i][1])
                    driver.find_element_by_id('recommendName').send_keys(infoList[i][0])
                    select = driver.find_element_by_tag_name("select")
                    select.click()
                    option = select.find_elements_by_tag_name('option')
                    for m in range(len(option)):
                        # print(option[m].text)
                        if option[m].text == infoList[i][6]:
                            option[m].click()
                            break
                    driver.find_element_by_class_name('greyBtn').click()
                    if driver.current_url == thisurl:
                        wait.until(lambda x:x.find_element_by_id('alert_popup_ok').is_displayed())
                        driver.find_element_by_id('alert_popup_ok').click()
                        failedUserList.append(unicode(infoList[i][2]))
                        print('has submited')
                    else:
                        wait.until(lambda x:x.find_element_by_class_name('stepUl').is_displayed)
                        print('success')
                    break


# begin
userinfoList = getUserInfoList(filePath)
submitInfos(driver,url,userinfoList)

result = open('FailedUser.txt','w')
result.writelines(failedUserList)
result.close()


# over .  Recycling resources
driver.quit()
os.system('taskkill /f /t /im chromedriver.exe')
