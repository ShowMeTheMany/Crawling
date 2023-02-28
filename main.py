import time

from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException

from openpyxl import Workbook
wb = Workbook()
ws = wb.create_sheet('유아용품')
wb.remove_sheet(wb['Sheet'])
ws.append(['상품명', '가격', '분류', '재고'])

i = 1

while True:
    options = webdriver.ChromeOptions()

    options.add_argument('--headless')
    UserAgent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36'
    options.add_argument('user-agent=' + UserAgent)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.get(url='https://www.coupang.com/np/categories/221934?page=' + str(i))
    time.sleep(5)

    try:
        product = driver.find_element(By.ID, 'productList')
        lis = product.find_elements(By.CLASS_NAME, 'baby-product')
        print('*' * 50 + ' ' + str(i) + 'Page Start!' + ' ' + '*' * 50)

        for li in lis:
            try:
                product = li.find_element(By.CLASS_NAME, 'name').text
                price = li.find_element(By.CLASS_NAME, 'price-value').text

                print('Product: ' + product)
                print('Price: ' + price)

                ws.append([product, price, '유아용품', str(10)])

            except Exception:
                pass

        print('*' * 50 + ' ' + str(i) + 'Page End!' + ' ' + '*' * 50)
        time.sleep(5)
        i += 1
        driver.quit()

    except NoSuchElementException:
        wb.save('/Users/aaron/Desktop/coupang_data/유아용품.xlsx')
        wb.close()
        exit(0)