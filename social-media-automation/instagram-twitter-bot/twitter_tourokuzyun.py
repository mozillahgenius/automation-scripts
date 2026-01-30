import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import subprocess
import pyperclip
import random
from datetime import datetime as dt, date, timedelta

#今の日付時：分：秒
strDate= dt.now().strftime("%H:%M:%S")
print(strDate)

#テキストファイル読み込み
txt_file=r"C:\Users\test\Documents\Twitter\登録文サンプル (1).txt" #変更

#TXTファイル書き込み
txt_write_file=r"C:\Users\test\Documents\test\test.txt"

#chrome task kill
result = subprocess.run('taskkill /f /im chrome.exe 2> nul',encoding='shift jis',shell=True,stdout=subprocess.PIPE)

#headless background
option = Options()

#ログイン情報を維持するための設定　
# 参考→https://rabbitfoot.xyz/selenium-chrome-profile
PROFILE_PATH ="C:\\Users\\user\\AppData\\Local\\Google\\Chrome\\User Data\\" # 変更
option.add_argument('--user-data-dir=' + PROFILE_PATH)
option.add_argument('--profile-directory=Defult')# 変更
#Getting Default Adapter failed error message
option.add_experimental_option('excludeSwitches', ['enable-logging'])

#もしファイルがなければ、0追加
try:
    with open(txt_write_file, "r", encoding="utf-8") as f:
        data = f.readlines()[0] 
        #print(data)
except :
    print("File New")
    f = open(txt_write_file, 'w', encoding='UTF-8')
    data=0
    f.write(str(data))

    f.close()

    print(data)


with open(txt_file, "r", encoding="utf-8") as f:
    line = f.readlines()
    count = 0
    list=[]
    for item in line:

        #行数をカウント
        count += 1
        
        #文字列の置換
        item_mod = item.replace('<br>', '\n')
        #item_mod = item.encode('latin-1', 'backslashreplace').decode('unicode_escape')
        
        #リストに追加
        list.append(item_mod)
        
    #ランダム表示
    #random.shuffle(list)
    #新しいリストへ追加
    new_list=list[int(data)]+"\n"+strDate
    print(new_list)

    #行数1を引く
    count=count-1
    print(data)

    #行数を確認 全ての行数と一致すればリセット
    if int(data)==count:
        print("reset data 0")
        f = open(txt_write_file, 'w', encoding='UTF-8')
        lines=0
        f.write(str(lines))

        f.close()
    else:
        #行数1足し算
        data=int(data)+1
        f = open(txt_write_file, 'w', encoding='UTF-8')

        f.write(str(data))

        f.close()

    #ブラウザを開く。 #options=option background
    driver = webdriver.Chrome(executable_path=ChromeDriverManager().install() ,options=option)

    #URL
    URL= "https://twitter.com/home"
    
    #URLを開く。
    driver.get(URL)
    
    #待機時間 3秒
    time.sleep(3)
    #Twitter_text_areaををクリック
    Twitter_text_area = driver.find_element('xpath','//*[@id="react-root"]/div/div/div[2]/main/div/div/div/div/div/div[3]/div/div[2]/div[1]/div/div/div/div[2]/div[1]/div/div/div/div/div/div/div/div/div/div/label/div[1]/div/div')
    Twitter_text_area.click()
    
    #待機時間 4秒
    time.sleep(4)
    #Twitter_text_areaをのテキスト入力
    Twitter_text_area = driver.find_element('xpath','//*[@id="react-root"]/div/div/div[2]/main/div/div/div/div/div/div[3]/div/div[2]/div[1]/div/div/div/div[2]/div[1]/div/div/div/div/div/div[2]/div/div/div/div/label/div[1]/div/div/div/div/div/div[2]/div/div/div/div')
    #Twitter_text_area.send_keys(new_list)

    pyperclip.copy(new_list)

    Twitter_text_area.send_keys(Keys.CONTROL,"v")

    #待機時間 4秒
    time.sleep(4)
    #ツィートボタンをクリック
    Button = driver.find_element('xpath','//*[@id="react-root"]/div/div/div[2]/main/div/div/div/div/div/div[3]/div/div[2]/div[1]/div/div/div/div[2]/div[3]/div/div/div[2]/div[3]/div/span/span')
    Button.click()
    
    time.sleep(4)
    driver.close()
