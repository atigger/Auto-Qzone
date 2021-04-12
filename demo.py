from selenium import webdriver

UA = 'user-agent="Mozilla/5.0 (Linux; Android 5.0; SM-N9100 Build/LRX21V) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/37.0.0.0 Mobile Safari/537.36 V1_AND_SQ_5.3.1_196_YYB_D QQ/5.3.1.2335 NetType/WIFI"'
headers = {"User-Agent": UA}

chrome_driver = "C:/Users/64995/AppData/Local/Google/Chrome/Application/chromedriver.exe"
chromeOptions = webdriver.ChromeOptions()
chromeOptions.add_argument(UA)
driver = webdriver.Chrome(executable_path=chrome_driver, options=chromeOptions)
driver.set_window_position(20, 40)
driver.set_window_size(1080, 1920)
driver.get()
