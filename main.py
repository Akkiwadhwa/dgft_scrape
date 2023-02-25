import time
from fake_useragent import UserAgent
import cv2
import numpy as np
import pandas as pd
import pytesseract
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

ua = UserAgent()
def get_cap(a):
    with open('filename.png', 'wb') as file:
        file.write(a.screenshot_as_png)
    image = cv2.imread('filename.png')
    image = ~image
    kernel = np.array([[0, -1, 0],
                       [-1, 5, -1],
                       [0, -1, 0]])

    image_sharp = cv2.filter2D(src=image, ddepth=-1, kernel=kernel)
    # a = cv2.fastNlMeansDenoisingColored(image, None, 10, 10, 7, 21)
    filename = "{}.png".format("temp")
    cv2.imwrite(filename, image_sharp)
    img = Image.open('temp.png')
    cap = pytesseract.image_to_string(img)
    return cap


url = "https://www.enquiry.icegate.gov.in/enquiryatices/dgftTrackIEC"
options = Options()
options.add_argument(f"user-agent={ua.random}")
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

df = pd.read_excel("test.xlsx")
for i in df["iecode"]:
    li_dict = []
    while True:
        driver.get(url)
        driver.find_element(By.NAME, "iecNO").send_keys(i)
        while True:
            driver.find_element(By.XPATH, '//*[@id="pagetable"]/tbody/tr[5]/td[2]/p/a/img').click()
            a = driver.find_element(By.ID, "capimg")
            cap = get_cap(a)
            if cap:
                break

        driver.find_element(By.XPATH, '//*[@id="pagetable"]/tbody/tr[3]/td[2]/img').click()

        driver.find_element(By.XPATH, '//*[@id="datepicker"]/table/tbody/tr[2]/td[2]/select').click()

        driver.find_element(By.XPATH, '//*[@id="datepicker"]/table/tbody/tr[2]/td[2]/select').find_elements(By.TAG_NAME,"option")[9].click()

        driver.find_element(By.XPATH, '//*[@id="datepicker"]/table/tbody/tr[2]/td[1]/select').click()

        driver.find_element(By.XPATH, '//*[@id="datepicker"]/table/tbody/tr[2]/td[1]/select').find_elements(
            By.TAG_NAME,
            "option")[
            0].click()

        driver.find_element(By.XPATH, '//*[@id="datepicker"]/table/tbody/tr[4]/td[7]').click()

        driver.find_element(By.XPATH, '//*[@id="pagetable"]/tbody/tr[4]/td[2]/img').click()

        driver.find_element(By.XPATH, '//*[@id="datepicker"]/table/tbody/tr[2]/td[1]/select').click()
        driver.find_element(By.XPATH,
                            '//*[@id="datepicker"]/table/tbody/tr[2]/td[1]/select').find_elements(
            By.TAG_NAME,
            "option")[
            1].click()

        driver.find_element(By.XPATH, '//*[@id="datepicker"]/table/tbody/tr[4]/td[4]').click()

        driver.find_element(By.ID, "captchaResp").send_keys(cap)
        driver.find_element(By.ID, "SubB").click()
        if driver.find_element(By.ID, "pagetable").find_elements(By.TAG_NAME, "th")[0].text == "LOCATION":
            list1 = driver.find_elements(By.CLASS_NAME, "tdText")
            for i1 in list1:
                f = i1.find_elements(By.TAG_NAME, "td")
                d = {
                    "LOCATION": f[0].text,
                    "IEC": f[1].text,
                    "SHIPPING BILL NO": f[2].text,
                    "SHIPPING BILL DATE": f[3].text,
                    "EGM DATE": f[4].text,
                    "CUSTOMS FILE NAME": f[5].text,
                    "ERROR CODE": f[6].text
                }
                li_dict.append(d)
            df = pd.DataFrame(li_dict)
            df.to_csv(f"{i}.csv")
            break
        else:
            pass
