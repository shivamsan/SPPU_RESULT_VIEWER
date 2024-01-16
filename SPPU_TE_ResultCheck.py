from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import openpyxl

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

def read_data(Start ,End):
    workbook = openpyxl.load_workbook('TE.xlsx')
    sheet = workbook.active

    data =[]
    for row in sheet[Start:End]:
        data.append([cell.value for cell in row])
    return data

data = read_data('A1','A158')

def open_res(seat_num):
    web = webdriver.Chrome(options=options)
    web.get('https://onlineresults.unipune.ac.in/Result/Dashboard/Default')

    time.sleep(1)

    BatchSelect = web.find_element("xpath",'//*[@id="tblRVList"]/tbody/tr[2]/td[4]/a/input')
    BatchSelect.click()

    time.sleep(0.5)

    Seat = web.find_element("xpath",'//*[@id="SeatNo"]')

    Seat.send_keys('T190554'+seat_num)

    Mom = web.find_element("xpath",'//*[@id="MotherName"]')
    S= int(seat_num)
    Mom.send_keys(data[S-201])

    GoRes = web.find_element("xpath",'//*[@id="btn"]')
    GoRes.click()

    print("Oprning...")
    print("Result Declared")

seat_num = input("Enter last 3 digits of your Exam Seat Number: ")
open_res(seat_num)

