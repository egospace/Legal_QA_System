from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import math
from loguru import logger
import os,sys
import time

wd = webdriver.Firefox(r'Utils\geckodriver')
wd.implicitly_wait(10)
ac = ActionChains(wd)
sys.path.append(r'data')
logger.add("interface_log_{time}.log", rotation="500MB", encoding="utf-8", enqueue=True, compression="zip", retention="10 days")

def init():
    wd.get("https://wenshu.court.gov.cn/website/wenshu/181010CARHS5BS3C/index.html?open=login")
    wd.switch_to_frame("contentIframe")
    wd.find_element_by_name('username').send_keys("账号")
    wd.find_element_by_name('password').send_keys("密码\n")
    wd.switch_to_default_content()
    element = wd.find_element_by_id("_view_1540966814000")
    element1 = element.find_element_by_class_name("search-middle").find_element_by_tag_name("input")
    element1.send_keys("消费者\n")


def clear_init():
    wd.find_element_by_xpath('//*[@id="clear-Btn"]').click()


def search_init(court):
    time.sleep(2)
    clear_init()
    a_search = wd.find_element_by_xpath('//*[@id="_view_1545034775000"]/div/div[1]/div[1]')
    a_search.click()
    send1 = a_search.find_element_by_xpath('//*[@id="flyj"]')
    # send1 = a_search.find_element_by_xpath('//*[@id="qbValue"]')
    send1.clear()
    # send1.send_keys("消费者权益")
    send1.send_keys("消费者权益保护法")
    send2 = a_search.find_element_by_xpath('//*[@id="s2"]')
    send2.clear()
    send2.send_keys(court)
    a_search.find_element_by_xpath('//*[@id="selectCon_other_ajlx"]').click()
    a_search.find_element_by_xpath('//*[@id="gjjs_ajlx"]/li[4]').click()
    a_search.find_element_by_xpath('//*[@id="s9"]').click()
    a_search.find_element_by_xpath('//*[@id="0301_anchor"]').click()
    a_search.find_element_by_xpath('//*[@id="s6"]').click()
    a_search.find_element_by_xpath('//*[@id="gjjs_wslx"]/li[3]').click()
    a_search.find_element_by_xpath('//*[@id="searchBtn"]').click()


def download():
    time.sleep(5)
    conditions = eval(wd.find_element_by_xpath('//*[@id="_view_1545184311000"]/div/div[2]/span').text)
    logger.info(conditions)
    conditions = conditions
    if int(conditions) == 0:
        return
    elif 0 < int(math.ceil(conditions / 15)) <= 40:
        conditions = int(math.ceil(conditions / 15))
    else:
        conditions = eval(wd.find_element_by_xpath('//*[@id="_view_1545184311000"]/div/div[2]/span').text)
        if int(conditions) == 0:
            return
        elif 0 < int(math.ceil(conditions / 15)) <= 40:
            conditions = int(math.ceil(conditions / 15))
        else:
            conditions = 40
    wd.find_element_by_css_selector('.pageSizeSelect').click()
    wd.find_element_by_css_selector('.pageSizeSelect > option:nth-child(3)').click()
    time.sleep(1)
    for index in range(conditions):
        try:
            '''全选的点击'''
            time.sleep(3)
            wd.find_element_by_xpath('//div[@class="LM_tool clearfix"]/div[4]/a[1]/label').click()
            time.sleep(1)
            wd.find_element_by_xpath(
                '//html/body/div/div[4]/div[2]//div[@class="LM_tool clearfix"]/div[4]/a[3]').click()
            logger.info(f'{time.strftime("%d{d}%H{h}%M{z}%S{s}:").format(d="日", h="时", z="分", s="秒")} '
                  f'第{index+1}页下载成功>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>')
            """下一页的点击"""
            time.sleep(1)
            wd.find_element_by_xpath('//div[@class="left_7_3"]/a[last()]').click()
        except:
            try:  # 特殊情况
                '''全选的点击'''
                time.sleep(3)
                wd.find_element_by_xpath('//div[@class="LM_tool clearfix"]/div[4]/a[1]/label').click()

                '''批量下载的点击'''
                wd.find_element_by_xpath('//html/body/div/div[4]/div[2]//div[@class="LM_tool clearfix"]/div[4]/a[3]').click()
            except:
                logger.add(sys.stderr, format="{time}{level}{message}", filter="my_module", level="INFO")
                logger.warning('第%d页可能出现缺失，缺失地区：' % index)
                time.sleep(2)
                """下一页的点击"""
                time.sleep(1)
                wd.find_element_by_xpath('//div[@class="left_7_3"]/a[last()]').click()
                '''全选的点击'''
                wd.find_element_by_xpath('//div[@class="LM_tool clearfix"]/div[4]/a[1]/label').click()
                '''批量下载的点击'''
                wd.find_element_by_xpath('//html/body/div/div[4]/div[2]//div[@class="LM_tool clearfix"]/div[4]/a[3]').click()


if __name__ == "__main__":
    init()
    time.sleep(30)
    with open('法院名称.txt', 'r', encoding="utf-8") as file_handle:
        for i in file_handle.readlines():
            court = i.strip('\n')
            logger.info(court)
            search_init(court)
            download()
