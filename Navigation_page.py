from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from datetime import datetime
import Global_var
from Insert_On_Datbase import insert_in_Local,create_filename
import sys, os
import ctypes
import string
import requests
import urllib.request
import urllib.parse
import re
import html
import wx
app = wx.App()
def ChromeDriver():

    chrome_options = Options()
    chrome_options.add_extension('C:\\Translation EXE\\BrowsecVPN.crx')
    browser = webdriver.Chrome(executable_path=str(f"C:\\Translation EXE\\chromedriver.exe"),chrome_options=chrome_options)
    browser.maximize_window()
    # browser.get("""https://chrome.google.com/webstore/detail/browsec-vpn-free-and-unli/omghfjlpggmjjaagoclmmobgdodcjboh?hl=en" ping="/url?sa=t&amp;source=web&amp;rct=j&amp;url=https://chrome.google.com/webstore/detail/browsec-vpn-free-and-unli/omghfjlpggmjjaagoclmmobgdodcjboh%3Fhl%3Den&amp;ved=2ahUKEwivq8rjlcHmAhVtxzgGHZ-JBMgQFjAAegQIAhAB""")
    wx.MessageBox(' -_-  Add Extension and Select Proxy Between 10 SEC -_- ', 'Info', wx.OK | wx.ICON_WARNING)
    time.sleep(15)  # WAIT UNTIL CHANGE THE MANUAL VPN SETtING
    browser.get("https://oil.gov.iq/index.php")
    wx.MessageBox(' -_-  Fill captch First -_- ', 'Info', wx.OK | wx.ICON_INFORMATION)
    browser.get("https://oil.gov.iq/?tender")
    time.sleep(2)
    
    tr = 2
    tender_href_list = []
    
    next_page = True
    while next_page == True:
        while True:
            try:
                release_date_list = browser.find_elements_by_xpath(f'//*[@id="resultTender"]/table/tbody/tr/td[4]')
                del release_date_list[0]
                break
            except:
                wx.MessageBox(' -_-  release_date_list not fount -_- ', 'error', wx.OK | wx.ICON_ERROR)
        tr = 2
        for release_date in release_date_list:
            release_date = release_date.get_attribute('innerText').strip()
            datetime_object = datetime.strptime(release_date, '%Y-%m-%d')
            publish_date = datetime_object.strftime("%Y-%m-%d")
            datetime_object_pub = datetime.strptime(publish_date, '%Y-%m-%d')
            User_Selected_date = datetime.strptime(str(Global_var.From_Date), '%Y-%m-%d')
            timedelta_obj = datetime_object_pub - User_Selected_date
            day = timedelta_obj.days
            if day >= 0:
                for tender_href in browser.find_elements_by_xpath(f'//*[@id="resultTender"]/table/tbody/tr[{str(tr)}]/td[6]/a'):
                    tender_href = tender_href.get_attribute('href').strip()
                    tender_href_list.append(tender_href)
                    tr += 1
                    break
            else:
                Scrap_data(browser, tender_href_list)
                next_page = False
                break
        next1 = True
        while next1 == True:
            try:
                for next_Page_list in browser.find_elements_by_xpath('/html/body/section/div/div/a[5]/div/i'):
                    browser.execute_script("arguments[0].scrollIntoView();", next_Page_list)
                    next_Page_list.click()
                    time.sleep(2)
                    next1 = False
                    break
            except:
                print('error on next page')
                next1 = True
                time.sleep(2)
                


def Scrap_data(browser, Tender_href):

    a = True
    while a == True:
        try:
            for href in Tender_href:
                browser.get(href)
                time.sleep(2)
                SegFeild = []
                for _ in range(45):
                    SegFeild.append('')

                get_htmlSource = ""
                for outerHTML in browser.find_elements_by_xpath('//*[@class="tenderDetails"]'):
                    get_htmlSource = outerHTML.get_attribute('outerHTML')
                    get_htmlSource = get_htmlSource.replace('href="upload/', 'href="https://oil.gov.iq/upload/')
                    break
                # Purchaser
                for Name_of_Directorate in browser.find_elements_by_xpath('//*[@class="tenderDetails"]/tbody/tr[2]/td[2]/p'):
                    Name_of_Directorate = Name_of_Directorate.get_attribute('innerText').replace('&nbsp;', '').strip()
                    if '(OCC)' in Name_of_Directorate:
                        SegFeild[2] = "Baghdad, Iraq<br>\nTel: +964 781 275 0356"
                        SegFeild[12] = Name_of_Directorate.strip()
                    elif '(TOC)' in Name_of_Directorate:
                        SegFeild[2] = "Dhi Qar - Al-Chibayish Road, Iraq<br>\nPhone: 09123276922 / 07831073600"
                        SegFeild[1] = 'info@toc.oil.gov.iq'
                        SegFeild[12] = Name_of_Directorate.strip()
                    elif '(MOO)' in Name_of_Directorate:
                        SegFeild[2] = "74 A Persian Gulf St, Kuwait City, Kuwait<br>\nTel: +965 1 858858"
                        SegFeild[8] = 'http://www.moo.gov.kw/'
                        SegFeild[1] = 'alnaft@moo.gov.kw'
                        SegFeild[12] = Name_of_Directorate.strip()
                    
                    elif '(KOTI)' in Name_of_Directorate:
                        SegFeild[2] = "Kirkuk 36001, Iraq<br>\nTel: +964 770 123 2541"
                        SegFeild[8] = 'http://www.koi.com/'
                        SegFeild[1] = 'zak@koi.com'
                        SegFeild[12] = Name_of_Directorate.strip()
                    elif '(NRC)' in Name_of_Directorate:
                        SegFeild[2] = "Baiji, Iraq<br>\n Phone: +974 7725 7608"
                        SegFeild[8] = 'http://www.nrc.oil.gov.iq/'
                        SegFeild[12] = Name_of_Directorate.strip()
                    elif '(HEESCO)' in Name_of_Directorate:
                        SegFeild[2] = "Baghdad-Daura Refinery Complex<br>\nTel: (+964) 07827836150 ,Fax: 770073"
                        SegFeild[8] = 'http://www.heesco.oil.gov.iq/en/enindex.html'
                        SegFeild[12] = Name_of_Directorate.strip()
                    elif '(BAIOTI)' in Name_of_Directorate:
                        SegFeild[2] = "Baghdad-Daura Refinery Complex<br>\nTel: (+964) 07827836150 ,Fax: 770073"
                        SegFeild[8] = 'http://www.heesco.oil.gov.iq/en/enindex.html'
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(SOMO)' in Name_of_Directorate:
                        SegFeild[2] = "Baghdad, Iraq<br>\nPhone: "
                        SegFeild[8] = "https://somooil.gov.iq/en/index.php"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(OEC)' in Name_of_Directorate:
                        SegFeild[2] = "Port Said Street,baghdad,Iraq<br>\nPhone: +9647832593219 "
                        SegFeild[8] = "http://oec.oil.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(BAS-OTI)' in Name_of_Directorate:
                        SegFeild[2] = "Basrah, Iraq<br>\nPhone: +964 780 911 4735"
                        SegFeild[8] = "http://bsroti.oil.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(AWLCO)' in Name_of_Directorate:
                        SegFeild[2] = "Building :39,street : 15,District: 309,Baghdad,Iraq<br>\nPhone: +964-780- 911 6291, Fax: +964 8825643"
                        SegFeild[8] = "http://www.awlco.net/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(IOTC)' in Name_of_Directorate:
                        SegFeild[2] = "Basrah, Iraq<br>\nPhone:"
                        SegFeild[8] = "http://iotc.oil.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(BOC)' in Name_of_Directorate:
                        SegFeild[2] = "Hey Al Kafaat, Basrah, Iraq<br>\nPhone:"
                        SegFeild[8] = "http://boc.oil.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(MOC)' in Name_of_Directorate:
                        SegFeild[2] = "Amarah, Iraq<br>\nPhone:"
                        SegFeild[8] = "http://moc.oil.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif 'شركة مصافي الجنوب ()' in Name_of_Directorate:
                        SegFeild[2] = "Iraq-basrah<br>\nPhone:0096440614713"
                        SegFeild[8] = "http://www.src.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(NOC)' in Name_of_Directorate:
                        SegFeild[2] = "Phone: 07481492275, Fax: 0096450255399"
                        SegFeild[8] = ""
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(NGC)' in Name_of_Directorate:
                        SegFeild[2] = "Kirkuk, Iraq<br>\nPhone: 07481492275,Fax: 0096450255399"
                        SegFeild[8] = "http://ngc.oil.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(BOTI)' in Name_of_Directorate:
                        SegFeild[2] = "Baghdad, Iraq<br>\n Phone: 00964-1 4250363,Fax: 00964-1 4257235"
                        SegFeild[8] = "http://www.boti.oil.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(IDC)' in Name_of_Directorate:
                        SegFeild[2] = "Baghdad/al-nidhal,park Al-saadoun,kirkuk, Iraq<br>\n Phone: 719173 - 7198278,Fax: 7178285"
                        SegFeild[8] = "http://www.idc.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(PRDC)' in Name_of_Directorate:
                        SegFeild[2] = "Baghdad - Bub Al Sham/ near Al-Sumood Station, Iraq<br>\n Phone: 0740 0233 637"
                        SegFeild[8] = "http://prdc.oil.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(MDOC)' in Name_of_Directorate:
                        SegFeild[2] = "Iraq<br>\nPhone: +20 1001797986"
                        SegFeild[8] = "https://imog-summit.com/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(MRC)' in Name_of_Directorate:
                        SegFeild[2] = "Baghdad, Iraq<br>\nPhone: +964 1 775 0300"
                        SegFeild[8] = "https://mrc.oil.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(OPDC)' in Name_of_Directorate:
                        SegFeild[2] = "Masafi St, Baghdad, Iraq<br>\nPhone: +964 790 145 9594"
                        SegFeild[8] = "http://www.opdc.oil.gov.iq/"
                        SegFeild[12] = Name_of_Directorate.strip()

                    elif '(OPC)' in Name_of_Directorate:
                        SegFeild[2] = "Iraq"
                        SegFeild[12] = Name_of_Directorate.strip()
                    elif '(GFC)' in Name_of_Directorate:
                        SegFeild[2] = "Taji, Iraq<br>\nPhone:"
                        SegFeild[8] = ""
                        SegFeild[12] = Name_of_Directorate.strip()
                    else:
                        SegFeild[12] = Name_of_Directorate.strip()
                    break

                # Title
                for Tender_Subject in browser.find_elements_by_xpath('//*[@class="tenderDetails"]/tbody/tr[1]/td[2]/p'):
                    Tender_Subject = Tender_Subject.get_attribute('innerText').replace('&nbsp;', '').strip()
                    Tender_Subject = string.capwords(str(Tender_Subject)).strip()
                    SegFeild[19] = Tender_Subject
                    break

                # # Email
                # for Email in browser.find_elements_by_xpath('/html/body/div[2]/center/table/tbody/tr[7]/td/table/tbody/tr/td[2]/center/table/tbody/tr[4]/td[2]/div'):
                #     Email = Email.get_attribute('innerText').replace('&nbsp;', '').replace('&nbsp;', '').strip().replace(' ','')
                #     SegFeild[1] = Email.strip()
                #     break

                # tender NO
                for Bid_number in browser.find_elements_by_xpath('//*[@class="tenderDetails"]/tbody/tr[3]/td[2]'):
                    Bid_number = Bid_number.get_attribute('innerText').replace('&nbsp;', '').strip()
                    SegFeild[13] = Bid_number.strip()
                    break

                # Release Date
                Release_Date = ""
                for Release_Date in browser.find_elements_by_xpath('//*[@class="tenderDetails"]/tbody/tr[4]/td[2]'):
                    Release_Date = Release_Date.get_attribute('innerText').replace('&nbsp;', '').strip()
                    break

                # Extention Date
                # Extention_Date = ""
                # for Extention_Date in browser.find_elements_by_xpath('/html/body/div[2]/center/table/tbody/tr[7]/td/table/tbody/tr/td[2]/center/table/tbody/tr[8]/td[2]'):
                #     Extention_Date = Extention_Date.get_attribute('innerText').replace('&nbsp;', '').strip()
                #     if Extention_Date == "لايوجد تمديد":
                #         Extention_Date = ""
                #     break

                # Document
                # for Attachment in browser.find_elements_by_xpath('/html/body/div[2]/center/table/tbody/tr[7]/td/table/tbody/tr/td[2]/center/table/tbody/tr[9]/td[2]/a'):
                #     Attachment = Attachment.get_attribute('href').strip()
                #     SegFeild[5] = Attachment
                #     break

                # Close Date
                try:
                    for Close_Date in browser.find_elements_by_xpath('//*[@class="tenderDetails"]/tbody/tr[5]/td[2]'):
                        Close_Date = Close_Date.get_attribute('innerText').strip()
                        datetime_object = datetime.strptime(Close_Date, "%Y-%m-%d")
                        mydate = datetime_object.strftime("%Y-%m-%d")
                        SegFeild[24] = mydate
                except:
                    SegFeild[24] = ""

                SegFeild[18] = "موضوع المناقصة: " + str(SegFeild[19]) + "<br>\n""اسم المديرية: " + str(SegFeild[12]) + "<br>\n""تاريخ الاصدار: " + str(Release_Date) + "<br>\n""تاريخ الاغلاق: " + str(SegFeild[24])

                SegFeild[7] = "IQ"

                # notice type
                SegFeild[14] = "2"

                SegFeild[22] = "0"

                SegFeild[26] = "0.0"

                SegFeild[27] = "0"  # Financier

                SegFeild[28] = str(href)

                # Source Name
                SegFeild[31] = 'oil.gov.iq'

                SegFeild[20] = ""
                SegFeild[21] = "" 
                SegFeild[42] = SegFeild[7]
                SegFeild[43] = "" 

                for SegIndex in range(len(SegFeild)):
                    print(SegIndex, end=' ')
                    print(SegFeild[SegIndex])
                    SegFeild[SegIndex] = html.unescape(str(SegFeild[SegIndex]))
                    SegFeild[SegIndex] = str(SegFeild[SegIndex]).replace("'", "''")
                a = False
                check_date(get_htmlSource, SegFeild)
                print(" Total: " + str(len(Tender_href)) + " Duplicate: " + str(Global_var.duplicate) + " Expired: " + str(Global_var.expired) + " Inserted: " + str(Global_var.inserted) + " Skipped: " + str(Global_var.skipped) + " Deadline Not given: " + str(Global_var.deadline_Not_given) + " QC Tenders: " + str(Global_var.QC_Tender),"\n")
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : ", sys._getframe().f_code.co_name + "--> " + str(e), "\n", exc_type, "\n", fname, "\n", exc_tb.tb_lineno)
            a = True
        ctypes.windll.user32.MessageBoxW(0, "Total: " + str(len(Tender_href)) + "\n""Duplicate: " + str(
            Global_var.duplicate) + "\n""Expired: " + str(Global_var.expired) + "\n""Inserted: " + str(
            Global_var.inserted) + "\n""Skipped: " + str(
            Global_var.skipped) + "\n""Deadline Not given: " + str(
            Global_var.deadline_Not_given) + "\n""QC Tenders: " + str(Global_var.QC_Tender) + "",
                                         "oil.gov.iq", 1)
        Global_var.Process_End()
        browser.quit()
        sys.exit()


def check_date(get_htmlSource, SegFeild):
    tender_date = str(SegFeild[24])
    nowdate = datetime.now()
    date2 = nowdate.strftime("%Y-%m-%d")
    try:
        if tender_date != '':
            deadline = time.strptime(tender_date , "%Y-%m-%d")
            currentdate = time.strptime(date2 , "%Y-%m-%d")
            if deadline > currentdate:
                insert_in_Local(get_htmlSource, SegFeild)
            else:
                print("Tender Expired")
                Global_var.expired += 1
        else:
            print("Deadline was not given")
            Global_var.deadline_Not_given += 1
    except Exception as e:
        exc_type , exc_obj , exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" , fname , "\n" , exc_tb.tb_lineno)


ChromeDriver()