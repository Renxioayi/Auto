# coding = utf-8
import xlrd
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import random
import datetime
import getpass

class Auto:

    # 加载chrome配置
    def __init__(self):
        option = webdriver.ChromeOptions()
        # 实际选择与chrome所在文件路径有关
        option.add_argument('--user-data-dir=C:\\Users\\' +  getpass.getuser()  + '\AppData\Local\Google\Chrome\\User Data')
        self.driver = webdriver.Chrome(options=option)
        self.driver.implicitly_wait(2)
        self.driver.maximize_window()
        self.h = self.driver.current_window_handle

    # 通过link_text定位
    def find_link_text(self, str):
        try:
            mouse = self.driver.find_element_by_link_text(str)
            ActionChains(self.driver).move_to_element(mouse).perform()
            mouse.click()
            return mouse
        except Exception as e:
            print(e)

    # 通过id定位
    def find_id(self, str):
        try:
            mouse = self.driver.find_element_by_id(str)
            ActionChains(self.driver).move_to_element(mouse).perform()
            mouse.click()
            return mouse
        except Exception as e:
            print(e)

    # 通过class_name定位
    def find_class_name(self, str):
        try:
            mouse = self.driver.find_element_by_class_name(str)
            ActionChains(self.driver).move_to_element(mouse).perform()
            mouse.click()
            return mouse
        except Exception as e:
            print(e)

    # 通过css_selector定位
    def find_css_selector(self, str):
        try:
            mouse = self.driver.find_element_by_css_selector(str)
            ActionChains(self.driver).move_to_element(mouse).perform()
            mouse.click()
            return mouse
        except Exception as e:
            print(e)

    # 通过xpath定位
    def find_xpath(self, str):
        try:
            mouse = self.driver.find_element_by_xpath(str)
            ActionChains(self.driver).move_to_element(mouse).perform()
            mouse.click()
            return mouse
        except Exception as e:
            print(e)

    # 创建等待ID方法
    def wait_id(self, str):
        try:
            WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.ID, str))
            )
        except Exception as e:
            print(e)

    # 创建等待text方法
    def wait_text(self, str):
        try:
            WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.LINK_TEXT, str))
            )
        except Exception as e:
            print(e)

    # 创建等待class_name方法
    def wait_class_name(self, str):
        try:
            WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.CLASS_NAME, str))
            )
        except Exception as e:
            print(e)

    # 创建等待css_selector方法
    def wait_selector(self, str):
        try:
            WebDriverWait(self.driver, 1).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, str))
            )
        except Exception as e:
            print(e)

    # 创建等待xpath方法
    def wait_xpath(self, str):
        try:
            WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, str))
            )
        except Exception as e:
            print(e)

    # 切换到新打开的窗口
    def changeCurrentWindow(self):
        try:
            all_h = self.driver.window_handles
            self.driver.switch_to.window(all_h[-1])
        except Exception:
            print("等待元素时异常")

    # 关闭窗口
    def closeWindow(self):
        try:
            self.driver.close()
        except Exception:
            print("关闭窗口异常")

    # 判断微信认证提示是否存在
    def is_confirm_exist(self):
        try:
            time.sleep(1)
            s = self.driver.find_elements_by_css_selector(
                css_selector='#app > div:nth-child(6) > div.weui-desktop-dialog__wrp.weui-desktop-dialog_simple > div > div.weui-desktop-dialog__ft > button.weui-desktop-btn.weui-desktop-btn_default')
            if len(s) == 0:
                print("认证提示不存在")
            elif len(s) == 1:
                print("认证提示存在")
                self.wait_selector(
                    '#app > div:nth-child(6) > div.weui-desktop-dialog__wrp.weui-desktop-dialog_simple > div > div.weui-desktop-dialog__ft > button.weui-desktop-btn.weui-desktop-btn_default')
                self.find_css_selector(
                    '#app > div:nth-child(6) > div.weui-desktop-dialog__wrp.weui-desktop-dialog_simple > div > div.weui-desktop-dialog__ft > button.weui-desktop-btn.weui-desktop-btn_default')
        except Exception:
            return True

    # 判断深色模式公告是否存在
    def is_notice_exist(self):
        try:
            time.sleep(1)
            s = self.driver.find_elements_by_css_selector(
                css_selector='#vue_app > div:nth-child(6) > div.weui-desktop-dialog__wrp.weui-desktop-dialog_darkmode-notice > div > div.weui-desktop-dialog__ft > button')
            if len(s) == 0:
                print("深色提示不存在")
            elif len(s) == 1:
                print("深色提示存在")
                self.wait_selector('#vue_app > div:nth-child(6) > div.weui-desktop-dialog__wrp.weui-desktop-dialog_darkmode-notice > div > div.weui-desktop-dialog__ft > button')
                self.find_css_selector('#vue_app > div:nth-child(6) > div.weui-desktop-dialog__wrp.weui-desktop-dialog_darkmode-notice > div > div.weui-desktop-dialog__ft > button')
        except Exception:
            return True

    # 进入一键登录主页
    def login(self):
        try:
            self.driver.get("http://account.mokasaas.com/admin/wechataccount/accountlist?ref=addtabs=1")
        except Exception:
            print("跳转到一键登录异常")

    # 传入采集链接
    def put_link(self,link):
        try:
            # 传入采集链接
            self.wait_class_name("ant-input")
            op = self.find_class_name('ant-input')

            # 调用键盘事件ctrl+a backspace
            op.send_keys(Keys.CONTROL, 'a')
            op.send_keys(Keys.BACKSPACE)
            op.send_keys(link)

        except Exception:
            print("传入采集链接异常")

    # 传入账号
    def put_account(self, gh_id):
        try:
            # 点击原始id
            self.wait_xpath('//*[@id="one"]/div/div[1]/div[1]/form/fieldset/div/div[1]/div/input[2]')
            op = self.find_xpath('//*[@id="one"]/div/div[1]/div[1]/form/fieldset/div/div[1]/div/input[2]')

            # 调用键盘事件ctrl+a backspace
            op.send_keys(Keys.CONTROL, 'a')
            op.send_keys(Keys.BACKSPACE)
            op.send_keys(gh_id)

            # 点击提交
            self.wait_xpath('//*[@id="one"]/div/div[1]/div[1]/form/fieldset/div/div[27]/div/button[1]')
            self.find_xpath('//*[@id="one"]/div/div[1]/div[1]/form/fieldset/div/div[27]/div/button[1]')

            # 等待公众号查询成功
            time.sleep(2)

            # 点击一键登录(根据页面具体的展示来确定xpath,暂时先用link_text来处理)
            self.wait_text('一键登录')
            self.find_link_text('一键登录')

            # 等待页面加载1.5s
            time.sleep(1.5)

        except Exception:
            print("传入公众号名称异常")

    # 进入新建图文消息页面
    def enter_new_page(self):
        try:
            # 判断认证确认是否存在
            self.is_confirm_exist()

            # 点击素材管理
            self.wait_text("素材管理")
            self.find_link_text("素材管理")
            time.sleep(0.5)

            # 点击新建图文消息
            self.wait_xpath('//*[@id="js_main"]/div[3]/div[1]/div[2]/button')
            self.find_xpath('//*[@id="js_main"]/div[3]/div[1]/div[2]/button')

        except Exception:
            print("进入新建图文消息页面异常")

    # 新建次条7图文
    def new_secondary_article(self,lastArticles):
        # 新建消息2
        self.wait_xpath('//*[@id="js_add_appmsg"]')
        self.find_xpath('//*[@id="js_add_appmsg"]')

        # 图文消息
        self.wait_class_name('js_create_article')
        self.find_class_name('js_create_article')

        # 点击采集文章
        self.wait_class_name("icondaoruwenzhang")
        self.find_class_name("icondaoruwenzhang")
        time.sleep(1.5)

        # 传入采集链接
        self.put_link(random.choice(lastArticles))

        # 点击确定
        self.wait_class_name("ant-btn-confirm")
        self.find_class_name("ant-btn-confirm")

    # 文案采集主方法
    def upload_main(self,firstArticles,lastArticles):
        try:
            time.sleep(1.5)
            # 关闭窗口
            self.closeWindow()

            # 切换到新打开的窗口
            self.changeCurrentWindow()
            time.sleep(1)

            # 判断深色模式提醒
            self.is_notice_exist()

            # 点击采集文章
            self.wait_class_name("icondaoruwenzhang")
            self.find_class_name("icondaoruwenzhang")
            time.sleep(1.5)

            # 传入采集链接
            self.put_link(random.choice(firstArticles))

            # 点击确定
            self.wait_class_name("ant-btn-confirm")
            self.find_class_name("ant-btn-confirm")

            # 二条
            self.new_secondary_article(lastArticles)

            # 三条
            self.new_secondary_article(lastArticles)

            # 四条
            self.new_secondary_article(lastArticles)

            # 五条
            self.new_secondary_article(lastArticles)

            # 六条
            self.new_secondary_article(lastArticles)

            # 七条
            self.new_secondary_article(lastArticles)

            # 八条
            self.new_secondary_article(lastArticles)

            # 保存
            self.wait_id('js_submit')
            self.find_id('js_submit')

            time.sleep(2)

        except Exception:
            print("文案上传主方法异常")

    # 前置方法
    def auto_before(self, gh_id):
        # 进入一键登录主页
        self.login()

        # 传入gh_id
        self.put_account(gh_id)
        time.sleep(1.5)

        # 关闭窗口
        self.closeWindow()

        # 切换到新打开的窗口
        self.changeCurrentWindow()

        # 进入文案上传界面
        self.enter_new_page()

    # 后置方法
    def auto_after(self,firstArticles,lastArticles):
        # 文案上传主方法
        self.upload_main(firstArticles,lastArticles)

def getFirstArticles():
    try:
        # 打开Excel表格
        wb = xlrd.open_workbook('账号.xlsx')
        # 获取当前正在显示的sheet
        sheet = wb.sheet_by_name('parameter')

        firstArticles = [(sheet.cell_value(i, 1))for i in range(1, sheet.nrows)]

        while '' in firstArticles:
            firstArticles.remove('')
    except FileNotFoundError:
        print("账号.xlsx文件不存在")
    return firstArticles

def getLastArticles():
    try:
        # 打开Excel表格
        wb = xlrd.open_workbook('账号.xlsx')
        # 获取当前正在显示的sheet
        sheet = wb.sheet_by_name('parameter')

        lastArticles = [(sheet.cell_value(i, 2))for i in range(1, sheet.nrows)]

        while '' in lastArticles:
            lastArticles.remove('')
    except FileNotFoundError:
        print("账号.xlsx文件不存在")
    return lastArticles

def getGh_ids():
    try:
        # 打开Excel表格
        wb = xlrd.open_workbook('账号.xlsx')
        # 获取当前正在显示的sheet
        sheet = wb.sheet_by_name('parameter')

        gh_ids = [(sheet.cell_value(i, 0))for i in range(1, sheet.nrows)]

        while '' in gh_ids:
            gh_ids.remove('')
    except FileNotFoundError:
        print("账号.xlsx文件不存在")
    return gh_ids

def run_main():
    firstArticles = getFirstArticles()
    lastArticles = getLastArticles()
    gh_ids = getGh_ids()

    auto = Auto()

    for gh_id in gh_ids:
        print(datetime.datetime.now().strftime('%Y-%m-%d  %H-%M-%S') + "    正在上传    " + gh_id)
        auto.auto_before(gh_id)
        auto.auto_after(firstArticles,lastArticles)

    input("=======================文案上传完成=======================")


if __name__ == '__main__':
    print("=======================欢迎使用微信公众号文案上传=======================")
    run_main()
