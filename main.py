import os
import random
from datetime import datetime
from urllib.request import urlretrieve
from PIL import Image
import openpyxl as openpyxl
from selenium import webdriver
from time import sleep
import requests
from bs4 import BeautifulSoup
from selenium.webdriver import ActionChains
import logging

url = "https://user.qzone.qq.com/649953543/"
# url = "http://httpbin.org/ip"
chrome_driver = "C:/Program Files/Google/Chrome/Application/chromedriver.exe"
file_url = "D:/文件/QQ资料.xlsx"

qq_ok = 0
qq_err = 0
err_qq = []


def load_file():
    workbook = openpyxl.load_workbook(file_url, data_only=True)
    worksheet = workbook['Sheet1']
    x_column = worksheet['B']
    x_len = len(x_column)  # 获取X轴长度
    qq = [None] * (x_len - 1)  # 定义数组长度
    tag_x = 0
    for i in range(x_len)[1:]:
        qq[tag_x] = x_column[i].value
        tag_x += 1

    y_column = worksheet['C']
    y_len = len(y_column)  # 获取Y轴长度
    pwd = [None] * (y_len - 1)  # 定义数组长度
    tag_y = 0
    for i in range(y_len)[1:]:
        pwd[tag_y] = y_column[i].value
        tag_y += 1
    return qq, pwd


def login(qq, pwd):
    chromeOptions = webdriver.ChromeOptions()
    driver = webdriver.Chrome(executable_path=chrome_driver, options=chromeOptions)
    driver.delete_all_cookies()
    # 设置浏览器窗口的位置和大小
    # driver.minimize_window()
    driver.set_window_position(20, 40)
    driver.set_window_size(1100, 700)
    # 打开一个页面（QQ空间登录页）
    driver.get(url)
    # 登录表单在页面的框架中，所以要切换到该框架
    driver.switch_to.frame("login_frame")
    # 通过使用选择器选择到表单元素进行模拟输入和点击按钮提交
    driver.find_element_by_id("switcher_plogin").click()
    driver.find_element_by_id("u").clear()
    driver.find_element_by_id("u").send_keys(qq)
    driver.find_element_by_id("p").clear()
    driver.find_element_by_id("p").send_keys(pwd)
    sleep(1)
    driver.find_element_by_id("login_button").click()
    sleep(3)
    tag = True
    for i in range(60):
        if login_status(driver):
            sj = random.randint(1, 5)
            sleep(sj)
            tag = False
            break
        sleep(1)
    if tag:
        global err_qq
        err_qq.append(qq)
        print("登录失败")
        logging.info(str(datetime.now()) + "--->登录失败")
    driver.quit()
    # 判断是否进入主页
    # login_success(driver)
    return


def login_status(driver):
    tag = False
    try:
        a = driver.switch_to.frame("tcaptcha_iframe")  # 切换到iframe

    except Exception as e:
        e = None

    try:
        a = driver.find_element_by_id("tcaptcha_drag_thumb")
        if a is not None:
            Tencent().pj(driver)
    except Exception as e:
        e = None

    try:
        a = driver.find_element_by_class_name("link-text")
        print("登录成功")
        logging.info(str(datetime.now()) + "--->登录成功")
        global qq_ok
        qq_ok += 1
        tag = True
    except Exception as e:
        e = None
    return tag


# 关闭其他应用程序
# pro_name:将要关闭的程序
def end_program(pro_name):
    os.system('%s%s' % ("taskkill /F /IM ", pro_name))
    print("结束进程:", pro_name)
    logging.info(str(datetime.now()) + "--->结束进程:" + str(pro_name))


class Tencent():
    def get_img(self):
        """
        获取验证码阴影图和原图
        :return:
        """
        # driver.switch_to.frame('tcaptcha_iframe')
        # 获取有阴影的图片
        src = self.driver.find_element_by_id('slideBg').get_attribute('src')
        # 分析图片地址，发现原图地址可以通过阴影图地址改动获取 只需要修改一下图片路径中的index
        # print(src)
        src_bg = src.replace('index=1', 'index=0')
        src_bg = src_bg.replace('img_index=1', 'img_index=0')
        # print(src_bg)
        # 将图片下载到本地
        urlretrieve(src, 'img1.png')
        urlretrieve(src_bg, 'img2.png')
        # 读取本地图片

        captcha1 = Image.open('img1.png')
        captcha2 = Image.open('img2.png')
        if captcha1 and captcha2 is not None:
            print("验证码读取完成")
            logging.info(str(datetime.now()) + "--->验证码读取完成")
        return captcha1, captcha2

    def resize_img(self, img):
        """
        下载的图片把网页中的图片进行了放大，所以将图片还原成原尺寸
        :param img: 图片
        :return: 返回还原后的图片
        """
        # 通过本地图片与原网页图片的比较，计算出的缩放比例 原图（680x390）缩小图（280x161）
        a = 2.428
        (x, y) = img.size
        x_resize = int(x // a)
        y_resize = int(y // a)
        """
        Image.NEAREST ：低质量
        Image.BILINEAR：双线性
        Image.BICUBIC ：三次样条插值
        Image.ANTIALIAS：高质量
        """
        img = img.resize((x_resize, y_resize), Image.ANTIALIAS)
        return img

    def is_pixel_equal(self, img1, img2, x, y):
        """
        比较两张图片同一点上的像数值，差距大于设置标准返回False
        :param img1: 阴影图
        :param img2: 原图
        :param x: 横坐标
        :param y: 纵坐标
        :return: 是否相等
        """
        pixel1, pixel2 = img1.load()[x, y], img2.load()[x, y]
        sub_index = 100
        # 比较RGB各分量的值
        if abs(pixel1[0] - pixel2[0]) < sub_index and abs(pixel1[1] - pixel2[1]) < sub_index and abs(
                pixel1[2] - pixel2[2]) < sub_index:
            return True
        else:
            return False

    def get_gap_offset(self, img1, img2):
        """
        获取缺口的偏移量
        """
        distance = 70
        for i in range(distance, img1.size[0]):
            for j in range(img1.size[1]):
                # 两张图片对比,(i,j)像素点的RGB差距，过大则该x为偏移值
                if not self.is_pixel_equal(img1, img2, i, j):
                    distance = i
                    return distance
        return distance

    def get_track(self, distance):
        """
        计算滑块的移动轨迹
        """
        # 通过观察发现滑块并不是从0开始移动，有一个初始值
        distance -= 30
        a = distance / 4
        track = [a, a, a, a]
        return track

    def shake_mouse(self):
        """
        模拟人手释放鼠标抖动
        """
        ActionChains(self.driver).move_by_offset(xoffset=-2, yoffset=0).perform()
        ActionChains(self.driver).move_by_offset(xoffset=2, yoffset=0).perform()

    def operate_slider(self, track):
        """
        拖动滑块
        当你调用ActionChains的方法时，不会立即执行，而是会将所有的操作按顺序存放在一个队列里，当你调用perform()方法时，队列中的时间会依次执行。
        :param track: 运动轨迹
        :return:
        """
        # 定位到拖动按钮
        slider_bt = self.driver.find_element_by_xpath('//div[@class="tc-drag-thumb"]')
        # 点击拖动按钮不放
        ActionChains(self.driver).click_and_hold(slider_bt).perform()
        # 按正向轨迹移动
        # move_by_offset函数是会延续上一步的结束的地方开始移动
        print("正在滑动")
        logging.info(str(datetime.now()) + "--->正在滑动")
        for i in track:
            ActionChains(self.driver).move_by_offset(xoffset=i, yoffset=0).perform()
            sleep(random.random() / 100)  # 每移动一次随机停顿0-1/100秒之间骗过了极验，通过率很高
        sleep(random.random())
        # 按逆向轨迹移动
        back_tracks = [-1, -0.5, -1]
        for i in back_tracks:
            sleep(random.random() / 100)
            ActionChains(self.driver).move_by_offset(xoffset=i, yoffset=0).perform()
        # 模拟人手抖动
        self.shake_mouse()
        sleep(random.random())
        # 松开滑块按钮
        ActionChains(self.driver).release().perform()

    def pj(self, driver):
        self.driver = driver
        a, b = self.get_img()
        a = self.resize_img(a)
        b = self.resize_img(b)
        distance = self.get_gap_offset(a, b)
        track = self.get_track(distance)
        self.operate_slider(track)


if __name__ == '__main__':
    logging.basicConfig(filename='qqkj.log', level=logging.INFO)
    logging.info(str(datetime.now()) + "--->开始运行")
    oldtime = datetime.now()
    qq, pwd = load_file()
    for i in range(len(qq)):
        print("QQ:", qq[i], "密码:", pwd[i])
        logging.info(str(datetime.now()) + "--->QQ:" + str(qq[i]) + "密码:" + str(pwd[i]))
        login(qq[i], pwd[i])
    qq_err = len(qq) - qq_ok
    newtime = datetime.now()
    date1 = newtime - oldtime
    end_program("chromedriver.exe")
    print("运行完成！成功", qq_ok, "个---失败", qq_err, "个 共用时", date1.seconds, "秒")
    logging.info(
        str(datetime.now()) + "--->运行完成！成功" + str(qq_ok) + "个---失败" + str(qq_err) + "个 共用时" + str(date1.seconds) + "秒")
    if qq_err > 0:
        print("失败的QQ号：", err_qq)
        logging.info(str(datetime.now()) + "--->失败的QQ号：" + str(err_qq))
