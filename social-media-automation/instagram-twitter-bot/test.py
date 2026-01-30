import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import subprocess

#chrome task kill
result = subprocess.run('taskkill /f /im chrome.exe 2> nul',encoding='shift jis',shell=True,stdout=subprocess.PIPE)

#headless background
option = Options()

#ログイン情報を維持するための設定 
# 参考→https://rabbitfoot.xyz/selenium-chrome-profile
PROFILE_PATH ="C:\\Users\\user\\AppData\\Local\\Google\\Chrome\\User Data\\" # 変更
option.add_argument('--user-data-dir=' + PROFILE_PATH)
option.add_argument('--profile-directory=Default')# 変更
#Getting Default Adapter failed error message
option.add_experimental_option('excludeSwitches', ['enable-logging'])

# ブラウザを開く。 #options=option background
driver = webdriver.Chrome(executable_path=ChromeDriverManager().install() ,options=option)

#URL
URL= "https://twitter.com/home"

# URLを開く。
driver.get(URL)

#待機時間
time.sleep(3)

#Twitter_text_areaををクリック
Twitter_text_area = driver.find_element('xpath','//*[@id="react-root"]/div/div/div[2]/main/div/div/div/div/div/div[3]/div/div[2]/div[1]/div/div/div/div[2]/div[1]/div/div/div/div/div/div/div/div/div/div/label/div[1]/div/div')
Twitter_text_area.click()

time.sleep(2)
#Twitter_text_areaをのテキスト入力
Twitter_text_area = driver.find_element('xpath','//*[@id="react-root"]/div/div/div[2]/main/div/div/div/div/div/div[3]/div/div[2]/div[1]/div/div/div/div[2]/div[1]/div/div/div/div/div/div[2]/div/div/div/div/label/div[1]/div/div/div/div/div/div[2]/div/div/div/div')
Twitter_text_area.send_keys('test')

time.sleep(2)
#ツィートボタンをクリック
Button = driver.find_element('xpath','//*[@id="react-root"]/div/div/div[2]/main/div/div/div/div/div/div[3]/div/div[2]/div[1]/div/div/div/div[2]/div[3]/div/div/div[2]/div[3]/div/span/span')
Button.click()

time.sleep(4)

driver.close()
