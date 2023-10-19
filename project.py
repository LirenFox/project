import tkinter as tk
from tkinter import messagebox

import io
import pandas as pd
#pip install pandas
#pip install openpyxl
import msoffcrypto
#pip install msoffcrypto-tool

import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
#from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

#driver = webdriver.Chrome(ChromeDriverManager().install())

def add_insurance():
    global user_choice
    user_choice = "加保"
    window.destroy()
    add_window = tk.Tk()
    add_window.title("加保")
    label = tk.Label(add_window, text="被加/退保人之代號：")
    label.pack(pady=10)
    entry = tk.Entry(add_window)
    entry.pack(pady=10)
    label2 = tk.Label(add_window, text="請輸入自然人憑證pin碼：")
    label2.pack(pady=10)
    entry2 = tk.Entry(add_window, show="*")
    entry2.pack(pady=10)
    #label3 = tk.Label(cancel_window, text="請輸入自然人憑證之身分證號：")
    #label3.pack(pady=10)
    #entry3 = tk.Entry(cancel_window, show="*")
    #entry3.pack(pady=10)
    confirm_button = tk.Button(add_window, text="確定", command=lambda: save_identifier(entry.get(), entry2.get()))
    #confirm_button = tk.Button(add_window, text="確定", command=lambda: save_identifier(entry.get(), entry2.get(), entry3.get())) #若要新增身分證號的輸入
    confirm_button.pack(pady=10)

def cancel_insurance():
    global user_choice
    user_choice = "退保"
    window.destroy()
    cancel_window = tk.Tk()
    cancel_window.title("退保")
    label = tk.Label(cancel_window, text="被加/退保人之代號：")
    label.pack(pady=10)
    entry = tk.Entry(cancel_window)
    entry.pack(pady=10)
    label2 = tk.Label(cancel_window, text="請輸入自然人憑證pin碼：")
    label2.pack(pady=10)
    entry2 = tk.Entry(cancel_window, show="*")
    entry2.pack(pady=10)
    #label3 = tk.Label(cancel_window, text="請輸入自然人憑證之身分證號：")
    #label3.pack(pady=10)
    #entry3 = tk.Entry(cancel_window, show="*")
    #entry3.pack(pady=10)
    confirm_button = tk.Button(cancel_window, text="確定", command=lambda: save_identifier(entry.get(), entry2.get()))
    #confirm_button = tk.Button(add_window, text="確定", command=lambda: save_identifier(entry.get(), entry2.get(), entry3.get())) #若要新增身分證號的輸入
    confirm_button.pack(pady=10)

saved_identifier = ""
saved_pin = 0
user_choice = ""
def save_identifier(identifier, pin):
    global saved_identifier, saved_pin
    saved_identifier = identifier
    saved_pin = pin
    window.quit()     # 結束視窗的事件循環
    print("選擇：", user_choice)  # 將選擇結果輸出在終端
    print("被加/退保人之代號：", identifier)

#若要新增身分證號的輸入
#saved_identity_Num = ""
# def save_identifier(identifier, pin, identity_Num): 
#     global saved_identifier, saved_pin, saved_identity_Num
#     saved_identity_Num = identity_Num
#     saved_identifier = identifier
#     saved_pin = pin
#     window.quit()     # 結束視窗的事件循環
#     print("選擇：", user_choice)  # 將選擇結果輸出在終端
#     print("被加/退保人之代號：", identifier)

# 建立主視窗
window = tk.Tk()
window.title("保險服務")

# 初始化使用者選擇
user_choice = ""

# 加保按鈕
add_button = tk.Button(window, text="加保", command=add_insurance)
add_button.pack(pady=10)

# 退保按鈕
cancel_button = tk.Button(window, text="退保", command=cancel_insurance)
cancel_button.pack(pady=10)

# 開始執行視窗迴圈
window.mainloop()

#讀取預設輸入
temp= io.BytesIO()
with open("input.xlsx", 'rb') as file:
    excel = msoffcrypto.OfficeFile(file)
    excel.load_key('123')
    excel.decrypt(temp)

df = pd.read_excel(temp)
print("user choice = ", user_choice)

chrome_options = webdriver.ChromeOptions()
chrome_options.add_extension('./0.0.1.3_0.crx')
driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
#if version error
#pip install --upgrade webdriver_manager

driver.get("https://edesk.bli.gov.tw/cpa/")
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入身分證號']").send_keys(df.at[2, saved_identifier]) #使用被加退保人之自然人憑證
#driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入身分證號']").send_keys("saved_identity_Num") #使用其他人之自然人憑證
driver.find_element(By.CSS_SELECTOR, "input[placeholder='含檢查碼共9位']").send_keys("15118422B") #保險證號
driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入6-8碼密碼']").send_keys(saved_pin) #自然人憑證pin碼
driver.find_element(By.CSS_SELECTOR, "button[aria-label='Close']").click()
driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
time.sleep(10)
#略過導覽
driver.find_element(By.CSS_SELECTOR, "div[class='d-sm-none d-md-flex justify-content-between align-items-center mt-7 pt-7'] div[role='button']").click()
#功能選單
driver.find_element(By.CSS_SELECTOR, "div[class='d-flex flex-shrink-0 justify-content-center align-items-center ml-5'] span").click()
#勞保申辦作業
driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/app-menu[1]/div[1]/nav[1]/div[2]/div[2]/ul[1]/li[1]/div[1]").click()
#單筆申報
driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/app-menu[1]/div[1]/nav[1]/div[2]/div[2]/ul[1]/li[1]/ul[1]/li[1]/div[1]/p[1]").click()


if user_choice == "加保":
   driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/app-menu[1]/div[1]/nav[1]/div[2]/div[2]/ul[1]/li[1]/ul[1]/li[1]/ul[2]/li[1]").click()
   time.sleep(1)
   driver.find_element(By.TAG_NAME, "button").click()
   time.sleep(2)
   if df.at[1, saved_identifier] == "外籍":#身分
      driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/div[1]/div[3]/div[1]/div[1]/app-meca0w01001[1]/app-bli-stepper[1]/app-apply-form[1]/form[1]/div[2]/div[3]/label[1]").click()
   elif df.at[1, saved_identifier] == "外籍配偶":
      driver.find_element(By.CSS_SELECTOR, "label[for='forn+2']").click()
   elif df.at[1, saved_identifier] == "大陸配偶":
      driver.find_element(By.CSS_SELECTOR, "label[for='forn+3']").click()

   if df.at[1, saved_identifier] != "本國人":
      if df.at[5, saved_identifier] == "男" :
         driver.find_element(By.CSS_SELECTOR, "label[for='tsex+0']").click()
      else :
         driver.find_element(By.CSS_SELECTOR, "label[for='tsex+1']").click()

   driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入姓名']").send_keys(df.at[0, saved_identifier]) #姓名
   driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入身分證號']").send_keys(df.at[2, saved_identifier]) #身分證號
   driver.find_element(By.CSS_SELECTOR, "input[placeholder='民國60年1月1日，請輸入0600101']").send_keys( "0"+str(df.at[4, saved_identifier]) ) #生日
   date = driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/div[1]/div[3]/div[1]/div[1]/app-meca0w01001[1]/app-bli-stepper[1]/app-apply-form[1]/form[1]/div[8]/div[1]/div[2]/select[1]")
   Select(date).select_by_index(df.at[6, saved_identifier]) #加保日
   special_identity = driver.find_element(By.XPATH,"/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/div[1]/div[3]/div[1]/div[1]/app-meca0w01001[1]/app-bli-stepper[1]/app-apply-form[1]/form[1]/div[9]/div[1]/div[2]/select[1]")
   Select(special_identity).select_by_index( df.at[7, saved_identifier] ) #特殊身份別
   driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入工資']").send_keys(df.at[8, saved_identifier]) #月薪資總額

   LaoG_special_identity = driver.find_element(By.XPATH,"/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/div[1]/div[3]/div[1]/div[1]/app-meca0w01001[1]/app-bli-stepper[1]/app-apply-form[1]/form[1]/div[13]/div[1]/div[2]/select[1]")
   Select(LaoG_special_identity).select_by_index(df.at[9, saved_identifier]) #勞基特殊身分別
   if df.at[9, saved_identifier] != 3: 
      driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='urate']").clear()
      driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='urate']").send_keys(df.at[10, saved_identifier]) #雇主提繳率
      driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='prate']").send_keys(df.at[11, saved_identifier]) #個人自願提繳率
      if df.at[12, saved_identifier] == 1:
         driver.find_element(By.CSS_SELECTOR, "label[for='efdmk']").click() #勞退提繳日期與(勞/就/職)加保日不同
         driver.find_element(By.CSS_SELECTOR, "#pefdte").send_keys(df.at[13, saved_identifier])
   
   if df.at[1, saved_identifier] == "本國人":
      driver.find_element(By.CSS_SELECTOR, "#hrdte").send_keys(df.at[14, saved_identifier]) #合於健保投保日期
   driver.find_element(By.CSS_SELECTOR, "#pin").send_keys(saved_pin) #自然人憑證pin碼
   driver.find_element(By.CSS_SELECTOR, ".btn.btn-primary.ml-sm-2.ml-md-5.flex-grow-1.flex-md-grow-0") #申報

else :
   driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/app-menu[1]/div[1]/nav[1]/div[2]/div[2]/ul[1]/li[1]/ul[1]/li[1]/ul[2]/li[2]").click()
   time.sleep(1)
   driver.find_element(By.CSS_SELECTOR, ".btn.btn-primary.mx-5").click()
   time.sleep(1)

   if df.at[1, saved_identifier] == "外籍":#身分
      driver.find_element(By.CSS_SELECTOR, "label[for='forn+1']").click()
      driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入居留證統一證號']").send_keys(df.at[3, saved_identifier])
   elif df.at[1, saved_identifier] == "外籍配偶":
      driver.find_element(By.CSS_SELECTOR, "label[for='forn+2']").click()
      driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入居留證統一證號']").send_keys(df.at[3, saved_identifier])
   elif df.at[1, saved_identifier] == "大陸配偶":
      driver.find_element(By.CSS_SELECTOR, "label[for='forn+3']").click()
      driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入居留證統一證號']").send_keys(df.at[3, saved_identifier])

   driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入姓名']").send_keys(df.at[0, saved_identifier]) #姓名
   driver.find_element(By.CSS_SELECTOR, "input[placeholder='請輸入身分證號']").send_keys(df.at[2, saved_identifier]) #身分證號
   driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='brDte']").send_keys( "0"+str(df.at[4, saved_identifier]) ) #生日
   date = driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/div[1]/div[3]/div[1]/div[1]/app-meca0w02001[1]/app-bli-stepper[1]/app-delete-form[1]/form[1]/div[8]/div[1]/div[2]/select[1]")
   Select(date).select_by_index(df.at[18, saved_identifier]) #退保日
   driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/div[1]/div[3]/div[1]/div[1]/app-meca0w02001[1]/app-bli-stepper[1]/app-delete-form[1]/form[1]/div[13]/div[2]/div[1]/input[1]").send_keys(df.at[15, saved_identifier]) #健保轉出(退保)日
   if df.at[16, saved_identifier] == "退保":
      driver.find_element(By.CSS_SELECTOR, "label[for='hqrTyp+1']").click() #原因(1)
   reason = driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/app-pages[1]/app-layout[1]/span[1]/div[1]/div[3]/div[1]/div[1]/app-meca0w02001[1]/app-bli-stepper[1]/app-delete-form[1]/form[1]/div[14]/div[2]/div[3]/div[1]/select[1]")
   Select(reason).select_by_index(df.at[17, saved_identifier]) #原因(2)
   driver.find_element(By.CSS_SELECTOR, "#pin").send_keys(saved_pin) #自然人憑證pin碼
   driver.find_element(By.CSS_SELECTOR, ".btn.btn-primary.mx-5") #申報

   