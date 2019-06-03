import pandas as pd
import datetime ,time, pysftp, os, re
import shutil
from selenium import webdriver

ADPW = pd.read_excel(r'\\vibo\nfs\CHPublic\客戶服務事業部\催收管理部\Calvin\WaiCui.pptx',
                     sheet_name='ADPW',
                     dtype=str)
# 設定帳號密碼
Account = ADPW.AD[0]
Password = ADPW.PW[0]
List = pd.read_excel(r'\\vibo\nfs\CHPublic\客戶服務事業部\催收管理部\Calvin\WaiCui.pptx',sheet_name='List',dtype='str')

path = 'C:/Users/' + Account + '/Downloads/'
today = datetime.datetime.now().strftime('%Y%m%d')

try:
    os.mkdir(path + '/WaiCui/')
except:
    time.sleep(0.1)
    
for i in range(len(List)):    
    
    # 設定基本查詢資料
    ActionId = str(List.委案別代號[i])
    CompanyID = str(List.公司代號[i])
    FTP_AD = str(List.FTP帳號[i])
    FTP_PW = str(List.FTP密碼[i])
    
    print('Dealing with: ' + List.委案別[i] + '_' + List.公司[i])
    print('Start Time Log: ' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    # 透過Selenium開啟Chrome瀏覽器
    driver = webdriver.Chrome()
    driver.set_page_load_timeout(1200)
    driver.implicitly_wait(1800)
    driver.get('http://colweb.tstartel.com/colweb/adLoginLogin')
    time.sleep(2)
    
    # 輸入帳號密碼
    driver.find_element_by_id('aduser').send_keys(Account)
    time.sleep(0.5)
    driver.find_element_by_id('password').send_keys(Password)
    time.sleep(0.5)
    driver.find_element_by_xpath('//button[@class="btn btn-default"]').click()
    time.sleep(0.5)
    
    # 展開資料夾
    time.sleep(2)
    driver.find_element_by_xpath('//i[@class="jstree-icon jstree-ocl"]').click()
    time.sleep(0.5)
    driver.find_elements_by_xpath('//i[@class="jstree-icon jstree-ocl"]')[1].click()
    time.sleep(0.5)
    driver.find_element_by_id('extRem1_anchor').click()
    time.sleep(10)

    # 選擇撈取資料的條件：委案別、委案公司與委案狀態
    webdriver.support.ui.Select(driver.find_element_by_id('actionId')).select_by_value(ActionId)
    webdriver.support.ui.Select(driver.find_element_by_id('company')).select_by_value(CompanyID)
    webdriver.support.ui.Select(driver.find_element_by_id('type')).select_by_value("OPEN")
    # 送出查詢
    driver.find_element_by_id('btnQry').click()
    time.sleep(1)
    # 全選項目並點擊下載
    driver.find_element_by_id('cb_grid').click()
    time.sleep(2)
    driver.find_element_by_id('btnDown').click()
    time.sleep(30)
    print('Download Succed!')
    
    # 搬到倉庫存放
    shutil.move(path + today + '.zip',
                '//vibo/nfs/CHPublic/客戶服務事業部/催收管理部/Calvin/WaiCui/' + today + '_' + ActionId + '_' + CompanyID + '.zip')    
    print('Move Succed!')
    
    # 將檔案上傳至FTP
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None
    sftp = pysftp.Connection(host = "172.23.1.44",
                             username = FTP_AD,
                             password = FTP_PW,
                             cnopts = cnopts)
    
    sftp.put(localpath = '//vibo/nfs/CHPublic/客戶服務事業部/催收管理部/Calvin/WaiCui/' + today + '_' + ActionId + '_' + CompanyID + '.zip',
             remotepath = '/ACCLIST/' + today + '_' + ActionId + '_' + CompanyID + '.zip')
    
    list = sftp.listdir('/ACCLIST/')
    for j in list:
        # 檢測檔案格式符合Dummy上傳的壓縮檔
        if bool(re.search('[0-9]{8}_[0-9]{2}_[0-9]{6}.*.zip', j)) == True:
            # 檢測日期是否超過3天
            if (datetime.datetime.today().date()- datetime.datetime.strptime(j[:8], "%Y%m%d").date()).days > 3:
                print('Remove: ' + j)
                sftp.remove('/ACCLIST/' + j)
                
    sftp.close()
    print('Upload SFTP Succed!')
    print('End Time Log: ' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    print('-----------------\n\n')
    driver.quit()