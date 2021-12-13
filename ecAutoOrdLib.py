import time,pickle,sys,os,imaplib,email,requests
import shutil
import zipfile
import datetime as dt2
import smtplib
from bs4 import BeautifulSoup
from imap_tools import MailBox, AND,OR,NOT

from datetime import datetime as dt 
from datetime import timedelta as td

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile

import ecAutoOrdConfig
import ecAutoOrdStart

dir_path = os.path.dirname(os.path.realpath(__file__))

def login(driver,setAccount,setPassword,loginURL,adminURL,cookieName,loginChk,source,runType):
    driver.get(loginURL)
    #driver.minimize_window()
    time.sleep(4)
    #強制登入情況下 刪除cookie直接引導人工登入
    if runType=='forceLogin':
        driver.delete_all_cookies()
        os.remove(cookieName)

    #如果瀏覽器本身就已登入
    if loginChk not in driver.current_url and runType!='forceLogin':
        pickle.dump(driver.get_cookies(), open(cookieName, "wb"))
        afterLoginAction(driver,source)
        if runType == 'none' or runType == 'auto':
            print('登入完成')
            return

    elif os.path.exists(cookieName):
        #使用既有cookie登入
        driver.delete_all_cookies()
        cookies = pickle.load(open(cookieName, "rb"))
        for cookie in cookies:
            driver.add_cookie(cookie)
        driver.get(adminURL)
        time.sleep(2)
        if loginChk in driver.current_url:
            print('登入狀態已失效，重新執行')
            driver.quit()
            os.remove(cookieName)
            #login(driver,loginURL,adminURL,cookieName,loginChk)
            ecAutoOrdStart.main(source,runType)
        else:
            pickle.dump(driver.get_cookies(), open(cookieName, "wb"))
            print('確認後台登入成功')
            time.sleep(2)
            afterLoginAction(driver,source)
            
    else:
        #引導人工登入
        if loginChk in driver.current_url:
            if source=='jplabo' or source=='hey':
                email = driver.find_element(By.ID, "login-input")
                password = driver.find_element(By.ID, "password")
                #email = driver.find_element_by_id("login-input")
                #password = driver.find_element_by_id("password")
            elif source=='well':
                email = driver.find_element(By.ID, "SysLoginForm_email")
                password = driver.find_element(By.ID, "SysLoginForm_password")
                #email = driver.find_element_by_id("SysLoginForm_email")
                #password = driver.find_element_by_id("SysLoginForm_password")
            email.send_keys(setAccount)
            password.send_keys(setPassword)
        if runType == 'auto':
            sendMail(source)
            sys.exit()

        print('需要重新登入驗證，等待使用者登入中 (密碼會自動填入，無須手動填寫)')
        driver.set_window_position(500, 200) 
        driver.set_window_size(800, 600) 
        while 1==1:
            if loginChk in driver.current_url:
                continue
            else:
                break
        #driver.minimize_window()
        #儲存cookie
        pickle.dump(driver.get_cookies(), open(cookieName, "wb"))
        driver.quit()
        if runType == 'forceLogin' or runType == 'none' or runType == 'auto':
            print('登入完成')
            return
        else:    
            ecAutoOrdStart.main(source,runType)

        
def afterLoginAction(driver,source):
    '''
    if source != 'well':
        try:
            popChk = driver.find_element_by_xpath('//button[@class="closeModal"]')
        except NoSuchElementException as e:
            return e
        else:
            popChk.click()
    '''

def exportData(driver,webType,exportURL,runType):
    if webType=='cyberbiz':
        time.sleep(2)
        driver.get(exportURL+'/reportion')
        time.sleep(2)
        todayStr = str(int(time.mktime(dt.strptime(dt.strftime(dt.today(), '%m/%d/%Y'), "%m/%d/%Y").timetuple())*1000))
        if runType == 'downloadToday':
            before7Str = todayStr
        elif runType == 'download':
            before7Str = str(int(time.mktime(dt.strptime(dt.strftime(dt.now()-td(days=7), '%m/%d/%Y'), "%m/%d/%Y").timetuple())*1000))
        export = exportURL+'/report?duallist=1%2C63%2C2%2C17%2C14%2C3%2C4%2C81%2C58%2C76%2C90%2C5%2C6%2C7%2C48%2C43%2C44%2C45%2C51%2C8%2C9%2C21%2C37%2C62%2C91%2C92%2C93%2C10%2C94%2C11%2C12%2C67%2C72%2C73%2C38%2C39%2C29%2C30%2C42%2C26%2C27%2C28%2C31%2C66%2C41%2C13%2C84%2C95%2C15%2C16%2C20%2C18%2C49%2C50%2C19%2C68%2C82%2C83%2C22%2C23%2C40%2C88%2C89%2C24%2C25%2C32%2C33%2C85%2C86%2C54%2C55%2C56%2C57%2C59%2C87%2C&order_source%5B%5D=ec&order_status%5B%5D=all&financial_status%5B%5D=all&fulfillment_status%5B%5D=all&return_status%5B%5D=all&receiver_email=&order_row=true&start='+before7Str+'&end='+todayStr+'&otp_attempt='
        driver.get(export)
    elif webType=='waca':
        time.sleep(2)
        driver.get(exportURL)
        time.sleep(2)
        todayStr = dt.strftime(dt.today(), '%Y-%m-%d')
        if runType == 'downloadToday':
            before7Str = todayStr
        elif runType == 'download':
            before7Str = dt.strftime(dt.now()-td(days=7), '%Y-%m-%d')
        export = 'https://admin.waca.ec/orders/export-orders-columns-ajax?startdate='+before7Str+'+00%3A00&enddate='+todayStr+'+23%3A59&pageSize=10&shops_ispin=0&shops_unread=0&date_type=created_at&reservation='+before7Str+'+00%3A00+-+'+todayStr+'+23%3A59&keywords_type=&keywords=&order_attribute=&status=&payment_status=&payment=&shipping=&member_type=unselect&member_level=&receiver_country=all&coupon=&subscribe_status=&subscribe_cycle_times=&subscribe_cycle=&other=&columns=1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,58,59,60,61,62,63,65,66,67,68,69'
        driver.get(export)

def downloadExportData():
    mailbox = MailBox(ecAutoOrdConfig.GetConfig('imapServer'))
    mailbox.login(ecAutoOrdConfig.GetConfig('imapAccount'), ecAutoOrdConfig.GetConfig('imapPassword'))  # or mailbox.folder.set instead 3d arg
    mailData = mailbox.fetch(AND(OR(from_=["system@waca.ec", "support@cyberbiz.co"]), date=dt2.date.today()), charset='utf8')
    content = []
    subjects = []
    for msg in mailData:
        content.append(msg.html)
        subjects.append(msg.subject)
        print(msg.subject)
        # 尋找附件 (針對cyberbiz)
        for att in msg.attachments:
            open(dir_path+'\\export\\'+format(msg.subject+att.filename), 'wb').write(att.payload)
 
    index = 0
    # 輸出超連結網址 (針對WACA)
    # 再檢查一次cookie狀態
    ecAutoOrdStart.main('well','return')
    time.sleep(2)
    for mails in content:
        soup = BeautifulSoup(mails, 'html.parser')
        a_tags = soup.find_all('a')
        for tag in a_tags:
            if 'orders-' in tag.get('href'):
                url = tag.get('href')
                getDownloadLink(url)
        index+=1

    # 刪除郵件 登出
    mailbox.delete([msg.uid for msg in mailbox.fetch(AND(OR(from_=["system@waca.ec", "support@cyberbiz.co"]), date=dt2.date.today()), charset='utf8')])
    mailbox.logout()

def getDownloadLink(url):
    cookies = pickle.load(open('cookies_well.pkl', "rb"))
    ckfile = {}
    for data in cookies:
        ckfile.update({data['name']:data['value']})
    r = requests.get(url,cookies=ckfile)
    soup = BeautifulSoup(r.text, 'html.parser')
    a_tags = soup.find_all('a')
    print(cookies[1]['value'])
    print(url)
    print(r.text)
    chkFile = False
    for tag in a_tags:
        if 'download/file' in tag.get('href'):
            downloadURL = tag.get('href')
            print(downloadURL)
            if 'admin.waca.ec' in downloadURL:
                r2 = requests.get(downloadURL, allow_redirects=True,cookies=ckfile)
                filename = 'well_'+dt2.date.strftime(dt2.date.today(), '%Y-%m-%d')+'.zip'
                fullpath = dir_path+'\\export\\'+filename
                open(fullpath, 'wb').write(r2.content)
                with zipfile.ZipFile(fullpath, 'r') as zip_ref:
                    zip_ref.extractall(dir_path+'\\export')
                os.remove(fullpath)
                chkFile = True
    if chkFile == False:
        print('WACA報表輸出完成，但獲取下載連結失敗')
        input()

def initProcess():
    try:
        profile_path = 'C:\\Users\\User\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\969sgqz9.python'
        options=Options()
        options.set_preference('profile', profile_path)
        service = Service('geckodriver.exe')

        driver = Firefox(service=service, options=options)
        driver.set_window_position(500, 200) 
        driver.set_window_size(800, 600)
        return driver
    except Exception as e: 
        errorProcess(driver,e,cookieName,source,runType)

def errorProcess(driver,e,cookieName,source,runType):
    print(e)
    print('發生意外錯誤，重新執行')
    driver.quit()
    ecAutoOrdStart.main(source,runType)

def endProcess(driver,runType):
    print('完成程序')
    driver.quit()
    if runType=='return':
        return
    else:    
        sys.exit()

def sendMail(source):
    smtp=smtplib.SMTP(ecAutoOrdConfig.GetConfig('smtpServer'), ecAutoOrdConfig.GetConfig('smtpPort'))
    smtp.ehlo()
    smtp.starttls()
    smtp.login(ecAutoOrdConfig.GetConfig('smtpAccount'),ecAutoOrdConfig.GetConfig('smtpPassword'))
    msg="登入程序發生異常，請前往查看("+source+")"
    status=smtp.sendmail(ecAutoOrdConfig.GetConfig('smtpAccount'), ecAutoOrdConfig.GetConfig('smtpAccount'), msg)
    if status=={}:
        print("郵件傳送成功!")
    else:
        print("郵件傳送失敗!")
    smtp.quit()
    

