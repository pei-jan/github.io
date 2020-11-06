#!/usr/bin/env python
# coding: utf-8

# In[ ]:


def 課程明細表依日期下載():
    start_year = eval(input('輸入起始西元年'))
    start_month = eval(input('輸入起始月份'))
    start_day = eval(input('輸入起始日期'))
    end_month = eval(input('輸入結束月份'))

    import openpyxl
    import xlrd
    import time
    import os
    import shutil
    import pyautogui
    import datetime
    from dateutil.relativedelta import relativedelta

    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.wait import WebDriverWait

    from selenium.webdriver.chrome.options import Options
    import sys
    from selenium.webdriver.common.action_chains import ActionChains

    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9210")


    driver = webdriver.Chrome(options=options)
    driver.maximize_window()

    #移除下載項目原report
    try:
        os.remove('/Users/pei/Downloads/report.xls')
        print('已刪除下載項目/report')
    except:
        pass

    #移除修課名單資料夾原report
    for n in range(20):
        try:
            os.remove('/Users/pei/Documents/py code/合併修課名單/report/report'+str(n)+'.xls')
            print('刪除原修課名單/report'+str(n))
        except:
            pass
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    driver.switch_to.frame('topFrame')
    WebDriverWait(driver,10).until(lambda driver : driver.find_element_by_link_text('管理工具'))
    driver.find_element_by_link_text('管理工具').click()

    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('leftTreeFrame'))
    driver.find_element_by_link_text('報表查詢').click()

    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('rightContentFrame'))
    driver.find_element_by_link_text('課程報表').click()

    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('rightContentFrame'))
    WebDriverWait(driver,30).until(lambda driver : driver.find_element_by_link_text('課程明細表'))
    driver.find_element_by_link_text('課程明細表').click()

    driver.switch_to.window(driver.window_handles[0])
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('rightContentFrame'))
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('barFrame'))
    driver.find_element_by_xpath('//*[@name="listAll"]').click()

    def downloads_done():
        for i in os.listdir("/Users/pei/Downloads/"):
            if ".crdownload" in i:
                time.sleep(0.5)
                downloads_done()

    def downloads_done2():
        list_download = []
        for i in os.listdir("/Users/pei/Downloads/"):
            list_download.append(i)
        if any('.crdownload'  in s for s in list_download):
            pass
        else:
            time.sleep(1)
            downloads_done2()

    def wait5():
        try:
            driver.switch_to.window(driver.window_handles[1])
            time.sleep(3)
            wait2()
        except:
            pass

    starttime = datetime.datetime.now()

    for i in range(end_month - start_month + 1):
        start_date = datetime.date(start_year, start_month, start_day)
        first_day = datetime.date(start_date.year, start_date.month, 1)
        last_day = first_day - datetime.timedelta(days = 1) + relativedelta(months=1)
        start_month += 1

        driver.switch_to.window(driver.window_handles[0])
        driver.switch_to.frame('rightContentFrame')
        driver.switch_to.frame('barFrame')
        driver.find_element_by_xpath('//*[@name="accessBegin"]').clear()
        time.sleep(1)
        driver.find_element_by_xpath('//*[@name="accessBegin"]').send_keys(' ')
        pyautogui.hotkey('command','left')
        driver.find_element_by_xpath('//*[@name="accessBegin"]').send_keys(str(first_day.year)+str(first_day.month).zfill(2)+str(first_day.day).zfill(2))

        driver.find_element_by_xpath('//*[@name="accessEnd"]').clear()
        time.sleep(1)
        driver.find_element_by_xpath('//*[@name="accessEnd"]').send_keys(' ')

        #today = str(datetime.datetime.today().year) + str(datetime.datetime.today().month) + str(datetime.datetime.today().day)
        pyautogui.hotkey('command','left')
        driver.find_element_by_xpath('//*[@name="accessEnd"]').send_keys(str(last_day.year)+str(last_day.month).zfill(2)+str(last_day.day).zfill(2))



        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('rightContentFrame'))
        WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('barFrame'))
        driver.find_element_by_xpath('//*[@name="apply"]').click()

        time.sleep(1.5)
        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('rightContentFrame'))
        WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('bodyFrame'))
        WebDriverWait(driver,20).until(lambda driver : driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/button[2]'))
        driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/button[2]').click()
        print('等待啟動下載')
        downloads_done2()

        print('等待下載中')


        #更改檔名
        downloads_done()
        print('完成下載')
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(2)
        driver.close()
        wait5()
        oldname='/Users/pei/Downloads/report.xls'
        newname='/Users/pei/Downloads/'+'report'+str(i)+'.xls'
        os.rename(oldname,newname)
        print('重新命名：'+str(oldname+'>>>'+newname))

        #移動檔案
        shutil.move('/Users/pei/Downloads/'+'report'+str(i)+'.xls', '/Users/pei/Documents/py code/合併修課名單/report/'+'report'+str(i)+'.xls')
        time.sleep(3)

    print()
    endtime = datetime.datetime.now()
    print('檔案下載時間：'+str(endtime - starttime))

def 課程明細表統整匯出():
    #已下載完成，單獨執行excel處理
    import io
    import numpy as np
    import pandas as pd
    from pathlib import Path
    import datetime
    import xlsxwriter

    start_year = eval(input('輸入起始西元年'))
    print()
    print('＊＊＊開始合併處理excel檔＊＊＊')
    print()
    print('讀取IHR檔案')
    print()
    ihr = pd.read_excel('/Users/pei/Documents/py code/合併修課名單/IHR資料.xlsx')
    ihr['新集團員編'] = '00'+ihr['新集團員編'].astype(str).str[0:6]+'\n'
    ihr['新集團員編'] = ihr['新集團員編'].str.replace('00nan','')
    print('完成')
    print()
    print('開始讀取report檔')
    print()

    files = Path('/Users/pei/Documents/py code/合併修課名單/report/').glob('*.xls')
    starttime = datetime.datetime.now()
    df =[]

    for f in files:
        print(str(f)) 
        line_list= []
        with open(str(f)) as f2:
            for line in f2.readlines():
                if '>' not in line and '<' not in line:
                    line_list.append(line)
        list_x = line_list[34:]
        list_x = [ x for x in list_x if "(已停用)" not in x]

        for k in range(10):
            #重複執行10次，直到所有'繳'後方的(0/2)都被取代
            for i in range(len(list_x)):
                try:
                    if '繳' in list_x[i] and len(list_x[i]) ==3 and '&nbsp'  not in list_x[i+1]:
                        list_x.pop(i+1)
                except:
                    pass

        step = 30
        for k in range(10):
            #重複執行10次，直到所有時數(小時)為0都被替補(list長度會因替補增長)
            for i in range(0,len(list_x),step):
                if (str(start_year) in list_x[6+i] or str(start_year+1)in list_x[6+i]) and '&nbsp' in list_x[7+i]:
                    list_x.insert(7+i,'0')
                    i+=1
                #替補學分為空白
                try:
                    if type(eval(list_x[3+i])) != float and '課' in list_x[2+i] :
                        list_x.insert(3+i,'0.0')
                        i+=1
                except:
                    pass

        b = [list_x[i:i+step] for i in range(0,len(list_x),step)]
        df_one = pd.DataFrame(b)

        df.append(df_one)
        print('已讀取：'+str(f))
    endtime = datetime.datetime.now()
    print()
    print('讀取時間：'+str(endtime - starttime))
    print()
    print('第一步合併excel')
    print()
    starttime = datetime.datetime.now()


    df2=pd.concat(df)

    print('第一步合併完成')
    endtime = datetime.datetime.now()
    print('合併時間：'+str(endtime - starttime))
    print()

    print('刪減欄位及篩選資料')
    starttime = datetime.datetime.now()
    df2.columns = ['學習目錄','課程名稱','課程類型','學分','班次名稱','開班單位','上課日期',
                   '時數(小時)','講師','修課人數','未完成人數','免修人數','通過人數','未通過人數',
                   '退選人數','學員總體滿意度','管理員總體滿意度','員工編號','姓名','職稱','到職日',
                   '到任日','單位','必選修','修課時數','通過狀態','通過日期','費用','測驗成績','學習報告']
    df2.sort_values(by=['課程名稱','員工編號','通過日期'],inplace=True, ascending=False)
    df2.drop_duplicates(['課程名稱','員工編號'],keep='first', inplace=True)


    df2['課程名稱'] = df2['課程名稱'].str.replace('&nbsp;','')

    df2['課程名稱'] = df2['課程名稱'].str.replace('&amp;','&')
    df2['課程名稱'] = df2['課程名稱'].str.replace('#039;',"'")
    df2['課程名稱'] = df2['課程名稱'].str.replace('&lt;',"<")
    df2['課程名稱'] = df2['課程名稱'].str.replace('&gt;',">")

    df2['學習目錄'] = df2['學習目錄'].str.replace('&nbsp;','')
    df2['班次名稱'] = df2['班次名稱'].str.replace('&nbsp;','')
    df2['開班單位'] = df2['開班單位'].str.replace('&nbsp;','')
    df2['講師'] = df2['講師'].str.replace('&nbsp;','')
    df2['員工編號'] = df2['員工編號'].str.replace('&nbsp;','')
    df2['姓名'] = df2['姓名'].str.replace('&nbsp;','')
    df2['職稱'] = df2['職稱'].str.replace('&nbsp;','')
    df2['到職日'] = df2['到職日'].str.replace('&nbsp;','')
    df2['到任日'] = df2['到任日'].str.replace('&nbsp;','')
    df2['單位'] = df2['單位'].str.replace('&nbsp;','')

    #新增時數欄位
    df2['修課時數(小時)'] = df2['修課時數']
    df2['CSR用時數(小時)'] = df2['修課時數(小時)']

    #刪減欄位
    df2 = df2[['學習目錄','課程名稱','課程類型','班次名稱','上課日期',
                   '員工編號','姓名','職稱','到職日',
                   '到任日','必選修','通過狀態','通過日期','測驗成績','時數(小時)','修課時數',
                   '修課時數(小時)','CSR用時數(小時)']]

    #篩選不需公司
    filter1 = ~df2['學習目錄'].str.contains('國泰飯店觀光事業')
    df2 = df2[filter1]
    print()
    endtime = datetime.datetime.now()
    print('刪減及篩選時間：'+str(endtime - starttime))
    print()


    df_join2 = []
    n = 100000
    df3=[]
    starttime = datetime.datetime.now()
    print('開始分割檔案')
    for i in range(0,df2.shape[0],n):
        df3.append(df2[i:i+n])
    print('分割完成')
    endtime = datetime.datetime.now()
    print('檔案分割時間：'+str(endtime - starttime))
    print()


    starttime = datetime.datetime.now()


    print('補時數資料')



    for i in range(len(df3)):
        for j in range(len(df3[i]['修課時數(小時)'])):
            try:
                df3[i]['修課時數(小時)'].iloc[j] = int(df3[i]['修課時數'].iloc[j][0:2]) +  (int(df3[i]['修課時數'].iloc[j][3:5])/60) + (int(df3[i]['修課時數'].iloc[j][6:8]) / 60/60)
            except:
                pass
        print('已換算第'+str(i+1)+'個檔案修課小時')

    for i in range(len(df3)):
        for j in range(len(df3[i]['時數(小時)'])):
            try:
                df3[i]['時數(小時)'].iloc[j] = eval(df3[i]['時數(小時)'].iloc[j].replace('\n',''))
            except:
                pass
        print('已取代第'+str(i+1)+'個檔案換行符')

    for i in range(len(df3)):
        for j in range(len(df3[i]['修課時數(小時)'])):
            try:
                if df3[i]['修課時數(小時)'].iloc[j] > df3[i]['時數(小時)'].iloc[j] and df3[i]['時數(小時)'].iloc[j] != 0 :
                    df3[i]['CSR用時數(小時)'].iloc[j] = df3[i]['時數(小時)'].iloc[j]  
                else:
                    df3[i]['CSR用時數(小時)'].iloc[j] = df3[i]['修課時數(小時)'].iloc[j]
            except:
                pass
        print('已比對第'+str(i+1)+'個檔案CSR用小時')

    endtime = datetime.datetime.now()
    print('處理時間：'+str(endtime - starttime))
    print()
    print('個別串IHR')

    starttime = datetime.datetime.now()

    for i in range(len(df3)):
        df4 = pd.merge(df3[i],ihr[['公司別','新集團員編','性別','人員類別代號','人員類別名稱','人事部門名稱(現職)',
                                    '人事科別名稱(現職)','職稱代號(現職)','在職狀況名稱','本公司離職日']],left_on='員工編號',right_on='新集團員編',how='left')
        df_join2.append(df4)
        print('已串第'+str(i+1)+'個檔')

    endtime = datetime.datetime.now()
    print('串資料時間：'+str(endtime - starttime))
    print()
    print('第二步合併檔案')
    starttime = datetime.datetime.now()
    df_join = pd.concat(df_join2)
    endtime = datetime.datetime.now()
    print('合併時間：'+str(endtime - starttime))
    print()

    print('調整列順序')
    print()
    df_join = df_join[['學習目錄', '課程名稱', '課程類型' , '班次名稱', '上課日期', 
                       '必選修', '通過狀態', '通過日期','測驗成績',
           '員工編號', '姓名', '職稱',  '性別','公司別',
            '人員類別代號', '人員類別名稱', '人事部門名稱(現職)', '人事科別名稱(現職)',
           '到職日', '到任日','職稱代號(現職)', '在職狀況名稱','本公司離職日',  
            '時數(小時)', '修課時數','修課時數(小時)', 'CSR用時數(小時)']]
    df_join = df_join.reset_index(drop=True)

    print('開始匯出檔案')

    df_part1 = df_join[0:1000000]
    df_part2 = df_join[1000000:2000000]
    starttime = datetime.datetime.now()
    df_part1.to_excel('/Users/pei/Documents/py code/合併修課名單/combine/report_combine_part1.xlsx', engine='xlsxwriter')
    endtime = datetime.datetime.now()
    print('part1匯出時間：'+str(endtime - starttime))

    starttime = datetime.datetime.now()
    df_part2.to_excel('/Users/pei/Documents/py code/合併修課名單/combine/report_combine_part2.xlsx', engine='xlsxwriter')
    endtime = datetime.datetime.now()
    print('part2匯出時間：'+str(endtime - starttime))
    print()
    print('＊＊＊完成＊＊＊')
    
    
def 課程明細表下載及統整():
    start_year = eval(input('輸入起始西元年'))
    start_month = eval(input('輸入起始月份'))
    start_day = eval(input('輸入起始日期'))
    end_month = eval(input('輸入結束月份'))

    import openpyxl
    import xlrd
    import time
    import os
    import shutil
    import pyautogui
    import datetime
    from dateutil.relativedelta import relativedelta

    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.wait import WebDriverWait

    from selenium.webdriver.chrome.options import Options
    import sys
    from selenium.webdriver.common.action_chains import ActionChains

    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9210")


    driver = webdriver.Chrome(options=options)
    driver.maximize_window()


    #移除下載項目原report
    try:
        os.remove('/Users/pei/Downloads/report.xls')
        print('已刪除下載項目/report')
    except:
        pass

    #移除修課名單資料夾原report
    for n in range(20):
        try:
            os.remove('/Users/pei/Documents/py code/合併修課名單/report/report'+str(n)+'.xls')
            print('刪除原修課名單/report'+str(n))
        except:
            pass
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    driver.switch_to.frame('topFrame')
    WebDriverWait(driver,10).until(lambda driver : driver.find_element_by_link_text('管理工具'))
    driver.find_element_by_link_text('管理工具').click()

    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('leftTreeFrame'))
    driver.find_element_by_link_text('報表查詢').click()

    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('rightContentFrame'))
    driver.find_element_by_link_text('課程報表').click()

    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('rightContentFrame'))
    WebDriverWait(driver,30).until(lambda driver : driver.find_element_by_link_text('課程明細表'))
    driver.find_element_by_link_text('課程明細表').click()

    driver.switch_to.window(driver.window_handles[0])
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('rightContentFrame'))
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('barFrame'))
    driver.find_element_by_xpath('//*[@name="listAll"]').click()

    def downloads_done():
        for i in os.listdir("/Users/pei/Downloads/"):
            if ".crdownload" in i:
                time.sleep(0.5)
                downloads_done()

    def downloads_done2():
        list_download = []
        for i in os.listdir("/Users/pei/Downloads/"):
            list_download.append(i)
        if any('.crdownload'  in s for s in list_download):
            pass
        else:
            time.sleep(1)
            downloads_done2()

    def wait5():
        try:
            driver.switch_to.window(driver.window_handles[1])
            time.sleep(3)
            wait2()
        except:
            pass

    starttime = datetime.datetime.now()

    for i in range(end_month - start_month + 1):
        start_date = datetime.date(start_year, start_month, start_day)
        first_day = datetime.date(start_date.year, start_date.month, 1)
        last_day = first_day - datetime.timedelta(days = 1) + relativedelta(months=1)
        start_month += 1

        driver.switch_to.window(driver.window_handles[0])
        driver.switch_to.frame('rightContentFrame')
        driver.switch_to.frame('barFrame')
        driver.find_element_by_xpath('//*[@name="accessBegin"]').clear()
        time.sleep(1)
        driver.find_element_by_xpath('//*[@name="accessBegin"]').send_keys(' ')
        pyautogui.hotkey('command','left')
        driver.find_element_by_xpath('//*[@name="accessBegin"]').send_keys(str(first_day.year)+str(first_day.month).zfill(2)+str(first_day.day).zfill(2))

        driver.find_element_by_xpath('//*[@name="accessEnd"]').clear()
        time.sleep(1)
        driver.find_element_by_xpath('//*[@name="accessEnd"]').send_keys(' ')

        #today = str(datetime.datetime.today().year) + str(datetime.datetime.today().month) + str(datetime.datetime.today().day)
        pyautogui.hotkey('command','left')
        driver.find_element_by_xpath('//*[@name="accessEnd"]').send_keys(str(last_day.year)+str(last_day.month).zfill(2)+str(last_day.day).zfill(2))



        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('rightContentFrame'))
        WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('barFrame'))
        driver.find_element_by_xpath('//*[@name="apply"]').click()

        time.sleep(1.5)
        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('rightContentFrame'))
        WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it('bodyFrame'))
        WebDriverWait(driver,20).until(lambda driver : driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/button[2]'))
        driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/button[2]').click()
        print('等待啟動下載')
        downloads_done2()

        print('等待下載中')


        #更改檔名
        downloads_done()
        print('完成下載')
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(2)
        driver.close()
        wait5()
        oldname='/Users/pei/Downloads/report.xls'
        newname='/Users/pei/Downloads/'+'report'+str(i)+'.xls'
        os.rename(oldname,newname)
        print('重新命名：'+str(oldname+'>>>'+newname))

        #移動檔案
        shutil.move('/Users/pei/Downloads/'+'report'+str(i)+'.xls', '/Users/pei/Documents/py code/合併修課名單/report/'+'report'+str(i)+'.xls')
        time.sleep(3)

    print()
    endtime = datetime.datetime.now()
    print('檔案下載時間：'+str(endtime - starttime))

    #已下載完成，單獨執行excel處理
    import io
    import numpy as np
    import pandas as pd
    from pathlib import Path
    import datetime
    import xlsxwriter

    print()
    print('＊＊＊開始合併處理excel檔＊＊＊')
    print()
    print('讀取IHR檔案')
    print()
    ihr = pd.read_excel('/Users/pei/Documents/py code/合併修課名單/IHR資料.xlsx')
    ihr['新集團員編'] = '00'+ihr['新集團員編'].astype(str).str[0:6]+'\n'
    ihr['新集團員編'] = ihr['新集團員編'].str.replace('00nan','')
    print('完成')
    print()
    print('開始讀取report檔')
    print()

    files = Path('/Users/pei/Documents/py code/合併修課名單/report/').glob('*.xls')
    starttime = datetime.datetime.now()
    df =[]

    for f in files:
        print(str(f)) 
        line_list= []
        with open(str(f)) as f2:
            for line in f2.readlines():
                if '>' not in line and '<' not in line:
                    line_list.append(line)
        list_x = line_list[34:]
        list_x = [ x for x in list_x if "(已停用)" not in x]

        for k in range(10):
            #重複執行10次，直到所有'繳'後方的(0/2)都被取代
            for i in range(len(list_x)):
                try:
                    if '繳' in list_x[i] and len(list_x[i]) ==3 and '&nbsp'  not in list_x[i+1]:
                        list_x.pop(i+1)
                except:
                    pass

        step = 30
        for k in range(10):
            #重複執行10次，直到所有時數(小時)為0都被替補(list長度會因替補增長)
            for i in range(0,len(list_x),step):
                if (str(start_year) in list_x[6+i] or str(start_year+1)in list_x[6+i]) and '&nbsp' in list_x[7+i]:
                    list_x.insert(7+i,'0')
                    i+=1
                #替補學分為空白
                try:
                    if type(eval(list_x[3+i])) != float and '課' in list_x[2+i] :
                        list_x.insert(3+i,'0.0')
                        i+=1
                except:
                    pass

        b = [list_x[i:i+step] for i in range(0,len(list_x),step)]
        df_one = pd.DataFrame(b)

        df.append(df_one)
        print('已讀取：'+str(f))
    endtime = datetime.datetime.now()
    print()
    print('讀取時間：'+str(endtime - starttime))
    print()
    print('第一步合併excel')
    print()
    starttime = datetime.datetime.now()


    df2=pd.concat(df)

    print('第一步合併完成')
    endtime = datetime.datetime.now()
    print('合併時間：'+str(endtime - starttime))
    print()

    print('刪減欄位及篩選資料')
    starttime = datetime.datetime.now()
    df2.columns = ['學習目錄','課程名稱','課程類型','學分','班次名稱','開班單位','上課日期',
                   '時數(小時)','講師','修課人數','未完成人數','免修人數','通過人數','未通過人數',
                   '退選人數','學員總體滿意度','管理員總體滿意度','員工編號','姓名','職稱','到職日',
                   '到任日','單位','必選修','修課時數','通過狀態','通過日期','費用','測驗成績','學習報告']
    df2.sort_values(by=['課程名稱','員工編號','通過日期'],inplace=True, ascending=False)
    df2.drop_duplicates(['課程名稱','員工編號'],keep='first', inplace=True)


    df2['課程名稱'] = df2['課程名稱'].str.replace('&nbsp;','')

    df2['課程名稱'] = df2['課程名稱'].str.replace('&amp;','&')
    df2['課程名稱'] = df2['課程名稱'].str.replace('#039;',"'")
    df2['課程名稱'] = df2['課程名稱'].str.replace('&lt;',"<")
    df2['課程名稱'] = df2['課程名稱'].str.replace('&gt;',">")

    df2['學習目錄'] = df2['學習目錄'].str.replace('&nbsp;','')
    df2['班次名稱'] = df2['班次名稱'].str.replace('&nbsp;','')
    df2['開班單位'] = df2['開班單位'].str.replace('&nbsp;','')
    df2['講師'] = df2['講師'].str.replace('&nbsp;','')
    df2['員工編號'] = df2['員工編號'].str.replace('&nbsp;','')
    df2['姓名'] = df2['姓名'].str.replace('&nbsp;','')
    df2['職稱'] = df2['職稱'].str.replace('&nbsp;','')
    df2['到職日'] = df2['到職日'].str.replace('&nbsp;','')
    df2['到任日'] = df2['到任日'].str.replace('&nbsp;','')
    df2['單位'] = df2['單位'].str.replace('&nbsp;','')

    #新增時數欄位
    df2['修課時數(小時)'] = df2['修課時數']
    df2['CSR用時數(小時)'] = df2['修課時數(小時)']

    #刪減欄位
    df2 = df2[['學習目錄','課程名稱','課程類型','班次名稱','上課日期',
                   '員工編號','姓名','職稱','到職日',
                   '到任日','必選修','通過狀態','通過日期','測驗成績','時數(小時)','修課時數',
                   '修課時數(小時)','CSR用時數(小時)']]

    #篩選不需公司
    filter1 = ~df2['學習目錄'].str.contains('國泰飯店觀光事業')
    df2 = df2[filter1]
    print()
    endtime = datetime.datetime.now()
    print('刪減及篩選時間：'+str(endtime - starttime))
    print()


    df_join2 = []
    n = 100000
    df3=[]
    starttime = datetime.datetime.now()
    print('開始分割檔案')
    for i in range(0,df2.shape[0],n):
        df3.append(df2[i:i+n])
    print('分割完成')
    endtime = datetime.datetime.now()
    print('檔案分割時間：'+str(endtime - starttime))
    print()


    starttime = datetime.datetime.now()


    print('補時數資料')



    for i in range(len(df3)):
        for j in range(len(df3[i]['修課時數(小時)'])):
            try:
                df3[i]['修課時數(小時)'].iloc[j] = int(df3[i]['修課時數'].iloc[j][0:2]) +  (int(df3[i]['修課時數'].iloc[j][3:5])/60) + (int(df3[i]['修課時數'].iloc[j][6:8]) / 60/60)
            except:
                pass
        print('已換算第'+str(i+1)+'個檔案修課小時')

    for i in range(len(df3)):
        for j in range(len(df3[i]['時數(小時)'])):
            try:
                df3[i]['時數(小時)'].iloc[j] = eval(df3[i]['時數(小時)'].iloc[j].replace('\n',''))
            except:
                pass
        print('已取代第'+str(i+1)+'個檔案換行符')

    for i in range(len(df3)):
        for j in range(len(df3[i]['修課時數(小時)'])):
            try:
                if df3[i]['修課時數(小時)'].iloc[j] > df3[i]['時數(小時)'].iloc[j] and df3[i]['時數(小時)'].iloc[j] != 0 :
                    df3[i]['CSR用時數(小時)'].iloc[j] = df3[i]['時數(小時)'].iloc[j]  
                else:
                    df3[i]['CSR用時數(小時)'].iloc[j] = df3[i]['修課時數(小時)'].iloc[j]
            except:
                pass
        print('已比對第'+str(i+1)+'個檔案CSR用小時')

    endtime = datetime.datetime.now()
    print('處理時間：'+str(endtime - starttime))
    print()
    print('個別串IHR')

    starttime = datetime.datetime.now()

    for i in range(len(df3)):
        df4 = pd.merge(df3[i],ihr[['公司別','新集團員編','性別','人員類別代號','人員類別名稱','人事部門名稱(現職)',
                                    '人事科別名稱(現職)','職稱代號(現職)','在職狀況名稱','本公司離職日']],left_on='員工編號',right_on='新集團員編',how='left')
        df_join2.append(df4)
        print('已串第'+str(i+1)+'個檔')

    endtime = datetime.datetime.now()
    print('串資料時間：'+str(endtime - starttime))
    print()
    print('第二步合併檔案')
    starttime = datetime.datetime.now()
    df_join = pd.concat(df_join2)
    endtime = datetime.datetime.now()
    print('合併時間：'+str(endtime - starttime))
    print()

    print('調整列順序')
    print()
    df_join = df_join[['學習目錄', '課程名稱', '課程類型' , '班次名稱', '上課日期', 
                       '必選修', '通過狀態', '通過日期','測驗成績',
           '員工編號', '姓名', '職稱',  '性別','公司別',
            '人員類別代號', '人員類別名稱', '人事部門名稱(現職)', '人事科別名稱(現職)',
           '到職日', '到任日','職稱代號(現職)', '在職狀況名稱','本公司離職日',  
            '時數(小時)', '修課時數','修課時數(小時)', 'CSR用時數(小時)']]
    df_join = df_join.reset_index(drop=True)

    print('開始匯出檔案')

    df_part1 = df_join[0:1000000]
    df_part2 = df_join[1000000:2000000]
    starttime = datetime.datetime.now()
    df_part1.to_excel('/Users/pei/Documents/py code/合併修課名單/combine/report_combine_part1.xlsx', engine='xlsxwriter')
    endtime = datetime.datetime.now()
    print('part1匯出時間：'+str(endtime - starttime))

    starttime = datetime.datetime.now()
    df_part2.to_excel('/Users/pei/Documents/py code/合併修課名單/combine/report_combine_part2.xlsx', engine='xlsxwriter')
    endtime = datetime.datetime.now()
    print('part2匯出時間：'+str(endtime - starttime))
    print()
    print('＊＊＊完成＊＊＊')

