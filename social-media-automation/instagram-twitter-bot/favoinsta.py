 # coding: UTF-8
# selenium version 3 
# pip install selenium==3
#
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.action_chains import ActionChains
import time
import random



def login():
	try:
		driver.get('https://www.instagram.com/accounts/login/?source=auth_switcher')
		f = open('insta.txt','a')
		f.write("instagramにアクセスしました\n")
		f.close()
		time.sleep(1)

		#メアドと、パスワードを入力
		driver.find_element_by_name('username').send_keys('YOUR_USERNAME')
		time.sleep(1)
		driver.find_element_by_name('password').send_keys('YOUR_PASSWORD')
		time.sleep(5)

		#ログインボタンを押す
		driver.find_elements_by_css_selector('button[type="submit"]')[0].click()
		time.sleep(random.randint(2, 5))
		f = open('insta.txt','a')
		f.write("instagramにログインしました\n")
		f.close()
		time.sleep(1)
	except:
		pass	

def tagsearch(tag):
	try:		
		instaurl = 'https://www.instagram.com/explore/tags/'
		driver.get(instaurl + tag)
		time.sleep(random.randint(2, 10))
		f = open('insta.txt','a')
		f.write("listtagより、tagで検索を行いました\n")
		f.close()
		time.sleep(10)
	except:
		pass	

def clicknice():

	niceclick_factor = random.randint(3, 5)
	try:
		if random.randint(0,1) == 0:
			# 人気投稿
			driver.find_elements_by_css_selector('div._aagw')[0].click()
		else:
			# 最新投稿
			driver.find_elements_by_css_selector('article > div:nth-child(3)')[0].find_elements_by_css_selector('div._aagw')[0].click()
		
		time.sleep(random.randint(2, 10))
		f = open('insta.txt','a')
		f.write("投稿をクリックしました\n")
		f.close()
		time.sleep(1)
		firstflag = True
		while True:
			nicebutton = driver.find_elements_by_css_selector('span._aamw > button')[0]
			if inner_clicknice(nicebutton) == True:		
				niceclick_factor -= 1
				f = open('insta.txt','a')
				f.write("投稿をいいねしました\n")
				f.close()
				time.sleep(1)
			if niceclick_factor == 0:
				break
			
			prevnextbutton = driver.find_elements_by_css_selector('div._aaqh > button')
			if prevnextbutton:
				# ２回目以降のループで行く戻るボタンが１つの場合は最後と判断
				if firstflag == False and len(prevnextbutton) == 1:
					break
				#取得された前次ボタンの最後が次なのでそちらをクリック
				prevnextbutton[-1].click()
				firstflag = False
				f = open('insta.txt','a')
				f.write("次の投稿へ移動しました\n")
				f.close()
				time.sleep(random.randint(random.randint(2, 5), random.randint(10, 15)))
	except:
		pass

def inner_clicknice(nicebutton):
	if nicebutton :
		# 「いいね！」したものを２度目のクリックで無効化しないようにする対処
		if nicebutton.find_elements_by_css_selector('svg')[0].get_attribute('aria-label') == 'いいね！':
			nicebutton.click()
			return (True)
	return (False)


if __name__ == '__main__':
	
	taglist = ['起業','独立']

	while True:
		driver = webdriver.Chrome()
		driver.implicitly_wait(60)
		driver.set_page_load_timeout(480)
		time.sleep(1)
		login()

		tagsearch(random.choice(taglist))
		clicknice()

		driver.close()

		abc = random.randint(random.randint(1200, 1800), random.randint(2400, 3000))
		f = open('insta.txt','a')
		f.write(str(abc)+"秒待機します\n")
		f.close()
		time.sleep(abc)
