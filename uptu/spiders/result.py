import os
import scrapy
from selenium.webdriver.common.action_chains import ActionChains
from scrapy import signals
from scrapy.xlib.pydispatch import dispatcher
from scrapy.http import TextResponse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import xlwt
from Tkinter import *
from ocr.testing import read_captcha
import cv2,urllib
import numpy as np


def url_to_image(url):
    # download the image, convert it to a NumPy array, and then read
    # it into OpenCV format
    resp = urllib.urlopen(url)
    image = np.asarray(bytearray(resp.read()), dtype="uint8")
    image = cv2.imdecode(image, cv2.IMREAD_COLOR)
 
    # return the image
    return image


class Result(scrapy.Spider):
    name = "btech"
    roll = 1409110901
    sheet = None
    workbook = None
    count = 1
    top = [["",0],["",1000]]
    allowed_domains = ['http://new.aktu.co.in/']
    start_urls = ['http://new.aktu.co.in/']

    def __init__(self, filename=None):
    	self.workbook = xlwt.Workbook()
        self.sheet = self.workbook.add_sheet('Sheet_1')
        
        self.driver = webdriver.Firefox()
        dispatcher.connect(self.spider_closed, signals.spider_closed)

    def spider_closed(self, spider):
        self.workbook.save('result.xls')
        self.driver.close()

    def add_in_sheet(self,item):
        self.count += 1
        if self.count == 2:
            self.sheet.write(0,0,'Name')
            self.sheet.write(0,1,"Father's Name")
            self.sheet.write(0,2,'Roll No.')
            self.sheet.write(0,3,'Enrollment No.')
            self.sheet.write(0,4,'Branch')
            self.sheet.write(0,5,'College')
            self.sheet.write(0,6,item['s1'])
            self.sheet.write(0,7,item['s2'])
            self.sheet.write(0,8,item['s3'])
            self.sheet.write(0,9,item['s4'])
            self.sheet.write(0,10,item['s5'])
            self.sheet.write(0,11,item['s6'])
            self.sheet.write(0,12,'GP')
            self.sheet.write(0,13,'Total')
        self.sheet.write(self.count,0,item['name'])
        self.sheet.write(self.count,1,item['father'])
        self.sheet.write(self.count,2,item['roll'])
        self.sheet.write(self.count,3,item['enroll'])
        self.sheet.write(self.count,4,item['branch'])
        self.sheet.write(self.count,5,item['clg'])
        self.sheet.write(self.count,6,item[item['s1']])
        self.sheet.write(self.count,7,item[item['s2']])
        self.sheet.write(self.count,8,item[item['s3']])
        self.sheet.write(self.count,9,item[item['s4']])
        self.sheet.write(self.count,10,item[item['s5']])
        self.sheet.write(self.count,11,item[item['s6']])
        self.sheet.write(self.count,12,item['gp'])
        self.sheet.write(self.count,13,item['tot'])
        if int(item['tot']) > self.top[0][1]:
            self.top[0][1] = int(item['tot'])
            self.top[0][0] = item['name']
        if int(item['tot']) < self.top[1][1]:
            self.top[1][1] = int(item['tot']) 
            self.top[1][0] = item['name']
        

    def parse_result(self, response):
        item = {}
        # Load the current page into Selenium
        
        self.driver.get(response)
        try:
            WebDriverWait(self.driver, 40).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_imgstud"]')))
        except TimeoutException:
            print "Time out"
            return
        # Sync scrapy and selenium so they agree on the page we're looking at then let scrapy take over
        resp = TextResponse(url=self.driver.current_url, body=self.driver.page_source, encoding='utf-8');
        temp = format(resp.xpath('//*[@id="lblname"]/text()').extract())
        item['name'] = temp[3:-2]

        temp = format(resp.xpath('//*[@id="lblfname"]/text()').extract())
        item['father'] = temp[3:-2]

        temp = format(resp.xpath('//*[@id="lblrollno"]/text()').extract())
        item['roll'] = temp[3:-2]

        temp = format(resp.xpath('//*[@id="lblenrollno"]/text()').extract())
        item['enroll'] = temp[3:-2]

        temp = format(resp.xpath('//*[@id="lblbranch"]/text()').extract())
        item['branch'] = temp[3:-2]

        temp = format(resp.xpath('//*[@id="lblcollegename"]/text()').extract())
        item['clg'] = temp[3:-2]

        temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/b/text()').extract())
        item['s1'] = temp[3:-2]

        t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/b/text()').extract())
        t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[4]/b/text()').extract())
        item[item['s1']] = t1[3:-2] + ' , '  + t2[3:-2]

        temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/b/text()').extract())
        item['s2'] = temp[3:-2]
        t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[3]/b/text()').extract())
        t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[4]/b/text()').extract())
        item[item['s2']] = t1[3:-2] + ' , '  + t2[3:-2]


        temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/b/text()').extract())
        item['s3'] = temp[3:-2]
        t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[3]/b/text()').extract())
        t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[4]/b/text()').extract())
        item[item['s3']] = t1[3:-2] + ' , '  + t2[3:-2]    

        temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/b/text()').extract())
        item['s4'] = temp[3:-2]
        t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td[3]/b/text()').extract())
        t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td[4]/b/text()').extract())
        item[item['s4']] = t1[3:-2] + ' , '  + t2[3:-2]

        temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[6]/td[2]/b/text()').extract())
        item['s5'] = temp[3:-2]
        t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[6]/td[3]/b/text()').extract())
        t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[6]/td[4]/b/text()').extract())
        item[item['s5']] = t1[3:-2] + ' , '  + t2[3:-2]

        temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[7]/td[2]/b/text()').extract())
        item['s6'] = temp[3:-2]
        t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[7]/td[3]/b/text()').extract())
        t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[7]/td[4]/b/text()').extract())
        item[item['s6']] = t1[3:-2] + ' , '  + t2[3:-2]

        temp = format(resp.xpath('//*[@id="ctl00_ContentPlaceHolder1_tr1"]/td[3]/text()').extract())
        item['gp'] = temp[3:-2]
        temp = format(resp.xpath('//*[@id="Pane0_content"]/table[3]/tbody/tr[2]/td[3]/text()').extract())
        print temp[5:-7]
        item['tot'] = temp[5:-7]
        self.add_in_sheet(item)

    def parse(self, response):
        while self.roll < 1409110903:
            self.driver.get('http://new.aktu.co.in/')
            try:
                WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_divSearchRes"]/center/table/tbody/tr[4]/td/center/div/div/img')))
            except:
                continue
	        # Sync scrapy and selenium so they agree on the page we're looking at then let scrapy take over
            resp = TextResponse(url=self.driver.current_url, body=self.driver.page_source, encoding='utf-8');
            rollno = self.driver.find_element_by_name('ctl00$ContentPlaceHolder1$TextBox1')
            rollno.send_keys(self.roll)
            captcha_url = format(resp.xpath('//*[@id="ctl00_ContentPlaceHolder1_divSearchRes"]/center/table/tbody/tr[4]/td/center/div/div/img/@src').extract())
            url = "http://new.aktu.co.in/" + captcha_url[3:-2]
            print url
            captcha = url_to_image(url)
            captcha_value = read_captcha(captcha)
            print captcha_value
            captcha_input = self.driver.find_element_by_name('ctl00$ContentPlaceHolder1$txtCaptcha')
            captcha_input.send_keys(captcha_value)
            input()
            submit = self.driver.find_element_by_name('ctl00$ContentPlaceHolder1$btnSubmit')
            actions = ActionChains(self.driver)
            actions.click(submit)
            actions.perform()
            resp = TextResponse(url=self.driver.current_url, body=self.driver.page_source, encoding='utf-8');
            if "Incorrect Code" in format(resp.xpath('*').extract()):
                continue
            self.parse_result(self.driver.current_url)
            self.roll += 1
        self.count +=3
        self.sheet.write(self.count,0,"First")
        self.sheet.write(self.count,1,self.top[0][0])
        self.sheet.write(self.count+1,0,"Last")
        self.sheet.write(self.count+1,1,self.top[1][0])
        return

