from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from time import sleep
import os
import sys


#SBI証券に自動ログイン、指定した数値以上の利益が出ている株のみを売却　又は指定以下を売却するプログラム

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

driver = webdriver.Chrome(resource_path('./driver/chromedriver.exe'))
    #options.add_argument('--headless')
url = "https://www.sbisec.co.jp/ETGate"
options = Options()
driver.implicitly_wait(10)
driver.maximize_window()
response = driver.get(url)

print("全て売却を行いますか？\ny/n")
all_sell = input()

user_id = input("ユーザー名を入力してください\n入力を間違えた場合はCtrlとCを同時押ししてプログラムを終了させてください:")

user_pw = input("ログインパスワードを入力してください:")

trade_pw = input("取引パスワードを入力してください:")
"""
ユーザー名:###
ログインパスワード:###
取引パスワード:###
"""
id_textbox = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/form/div/div/div/dl/dd[1]/div/input")
if id_textbox:
    id_textbox.click()
    id_textbox.send_keys(user_id)
    pw_textbox = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/form/div/div/div/dl/dd[2]/div/input")
    if pw_textbox:
        pw_textbox.click()
        pw_textbox.send_keys(user_pw)
        login_button = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/form/div/div/div/p[2]/a/input")
        if login_button:
            login_button.click()
            sleep(2)
    
account_manage = driver.find_element_by_xpath("/html/body/div[1]/div[1]/div[2]/div/ul/li[3]/a/img")

list_assessment = []

if account_manage:
    account_manage.click()
    
                                                
assessment_table = driver.find_element_by_xpath("/html/body/div[1]/table/tbody/tr/td[1]/table/tbody/tr[2]/td/table[1]/tbody/tr/td/form/table[2]/tbody/tr[1]/td[2]/table[5]/tbody/tr/td[3]/table[4]")
                                                
assessment_value = assessment_table.find_elements(By.TAG_NAME, "font")

if all_sell == "n":
    print("+値以上の設定ができます。半角数字を入力して下さい")
    int_profit = input()
    inter = int(int_profit)
r = 3
k = 0
for assessment in list(assessment_value):
    assessment_text = assessment
    list_assessment.append(assessment_text)
    
for i in list_assessment:

    list_text = i.text
    k += 1
    r += 2   
    sleep(1)
    if k >= 6:
        int_list_text = list_text[1:]
        intin_list_text = int(int_list_text)
        try:
        
            if all_sell == "y":
                print(list_text)                              
                sell_capital = driver.find_element_by_xpath("/html/body/div[1]/table/tbody/tr/td[1]/table/tbody/tr[2]/td/table[1]/tbody/tr/td/form/table[2]/tbody/tr[1]/td[2]/table[5]/tbody/tr/td[3]/table[4]/tbody/tr[{}]/td[2]/a[2]".format(r)).get_attribute("href")
                driver.execute_script("window.open()") #make new tab
                driver.switch_to.window(driver.window_handles[1]) #switch new tab
                driver.get(sell_capital)
                sleep(1)

                all_sell_btn = driver.find_element_by_xpath("/html/body/table/tbody/tr/td[1]/form/div[4]/div[1]/div[3]/table/tbody/tr[7]/td/label[2]")
                all_sell_btn.click()
                sleep(1)

                input_text_pass = driver.find_element_by_xpath("/html/body/table/tbody/tr/td[1]/form/div[4]/div[1]/div[7]/div/div/p[2]/input[3]")
                sleep(1)

                input_text_pass.click()
                input_text_pass.send_keys(trade_pw)
                sleep(1)

                order_submid = driver.find_element_by_name("ACT_estimate")
                order_submid.click()
                sleep(1)

                if driver.find_element_by_xpath("/html/body/table/tbody/tr/td[1]/form/div[6]/p/a/img"):
                    order_click = driver.find_element_by_xpath("/html/body/table/tbody/tr/td[1]/form/div[6]/p/a/img")
                    order_click.click()
                    sleep(1)

                    driver.close()
                    driver.switch_to.window(driver.window_handles[0]) 
            

            else:
                if intin_list_text >= inter:
                    print(list_text)
                    sell_capital = driver.find_element_by_xpath("/html/body/div[1]/table/tbody/tr/td[1]/table/tbody/tr[2]/td/table[1]/tbody/tr/td/form/table[2]/tbody/tr[1]/td[2]/table[5]/tbody/tr/td[3]/table[4]/tbody/tr[{}]/td[2]/a[2]".format(r)).get_attribute("href")
                    driver.execute_script("window.open()") #make new tab
                    driver.switch_to.window(driver.window_handles[1]) #switch new tab
                    driver.get(sell_capital)
                    sleep(1)

                    all_sell_btn = driver.find_element_by_xpath("/html/body/table/tbody/tr/td[1]/form/div[4]/div[1]/div[3]/table/tbody/tr[7]/td/label[2]")
                    all_sell_btn.click()
                    sleep(1)

                    input_text_pass = driver.find_element_by_xpath("/html/body/table/tbody/tr/td[1]/form/div[4]/div[1]/div[7]/div/div/p[2]/input[3]")
                    sleep(1)

                    input_text_pass.click()
                    input_text_pass.send_keys(trade_pw)
                    sleep(1)

                    order_submid = driver.find_element_by_name("ACT_estimate")
                    order_submid.click()
                    sleep(1)

                    if driver.find_element_by_xpath("/html/body/table/tbody/tr/td[1]/form/div[6]/p/a/img"):
                        order_click = driver.find_element_by_xpath("/html/body/table/tbody/tr/td[1]/form/div[6]/p/a/img")
                        sleep(1)

                        order_click.click()
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        
        
                            
        except NoSuchElementException:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
