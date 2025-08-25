from selenium import webdriver
from time import sleep
from Excels import *
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys

import openpyxl
driver=webdriver.Chrome()
driver.implicitly_wait(10)
driver.get('https://testautomationpractice.blogspot.com/')
driver.maximize_window()
sleep(5)
file=r"C:\Users\yanna\Desktop\QSPIDERS FILES\Information-datapy.xlsx"

rows=get_row_num(file,"Sheet1")
for row in range(2,rows+1):
    name=read_cell_data(file,"Sheet1",row,1)
    email=read_cell_data(file,"Sheet1",row,2)
    phone=read_cell_data(file,"Sheet1",row,3)
    address=read_cell_data(file,"Sheet1",row,4)
    gender=read_cell_data(file,"Sheet1",row,5)
    days=read_cell_data(file,"Sheet1",row,6)
    country=read_cell_data(file,"Sheet1",row,7)
    colors=read_cell_data(file,"Sheet1",row,8)
    sorted_list=read_cell_data(file,"Sheet1",row,9)

    print(name,email,phone,address,gender,days,country,colors,sorted_list)

    #--- testing  web application with excel data ....
    driver.find_element("id","name").send_keys(name)
    sleep(2)
    driver.find_element("id","email").send_keys(email)
    sleep(2)
    driver.find_element("id","phone").send_keys(phone)
    sleep(2)
    driver.find_element("id","textarea").send_keys(address)
    sleep(5)
    #-----------------
    if gender=='Male':
        driver.find_element("id","male").click()
    else:
        driver.find_element("id","female").click()
    sleep(5)
    #"------------------"
    days=driver.find_elements("xpath", "//input[@type='checkbox']/../label")
    for day in days:
        day.click()
    sleep(5)
    #------------------
    D1=driver.find_element("id","country")
    obj1=Select(D1)
    obj1.select_by_visible_text(country)
    sleep(5)
    #------------------
    D2 = driver.find_element("id", "colors")
    obj2 = Select(D2)
    obj2.select_by_visible_text(colors)
    sleep(5)
    #---------------------
    D3 = driver.find_element("id", "animals")
    obj3 = Select(D3)
    obj3.select_by_visible_text(sorted_list)
    sleep(5)
    driver.refresh()
    sleep(2)
driver.quit()

