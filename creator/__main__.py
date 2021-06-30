from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.common.exceptions import NoSuchElementException
from openpyxl import Workbook
from selenium.webdriver.firefox.options import Options

def main():

    sborka_id = input('Enter sborka id: ')

    options = Options()
    options.headless = True
    driver = webdriver.Firefox(options=options)
    driver.get("https://iis.dclink.org.ua/Sborki.aspx")

    driver.find_element(By.XPATH, "//h2[contains(text(), 'Выполнить вход')]")
    driver.find_element(By.XPATH, "//input[@name='ctl00$MainContent$LoginUser$UserName']").send_keys('balaban')
    driver.find_element(By.XPATH, "//input[@name='ctl00$MainContent$LoginUser$Password']").send_keys('1sG2Q6wl')
    driver.find_element(By.XPATH, "//input[@type='submit']").click()

    window_before = driver.window_handles[0]
    # driver.find_element(By.XPATH, "//a[contains(@href, '{}')]".format(sborka_id)).click()
    driver.get("https://iis.dclink.org.ua/SborkaForm.aspx?ID={}".format(sborka_id))
    # window_after = driver.window_handles[1]
    # driver.switch_to.window(window_after)



    SNs = driver.find_elements(By.XPATH, "//span[contains(@id, 'SN')]")
    cleaned_SNs = []
    for SN in SNs:
        cleaned_SNs.append(int(SN.text))
    # RN = driver.find_element(By.XPATH, "//span[contains(@id, 'НомерПН')]")
    # PART = driver.find_element(By.XPATH, "//span[contains(@id, 'Описание')]")

    driver.quit()



    workbook = Workbook()
    sheet = workbook.active


    for counter, sn in zip(range(1, len(cleaned_SNs)), cleaned_SNs):
        sheet["A{}".format(counter)] = sn

    workbook.save(filename='../Sns.xlsx')


if __name__ == '__main__':
    main()
