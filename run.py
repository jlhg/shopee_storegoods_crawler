#!/usr/bin/env python
import logging
import os
import sys

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import xlsxwriter

STORE_URL_FILE = 'store_url.txt'
LOG_FILE = 'log.txt'

# Set logging
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(name)-12s %(levelname)-8s %(message)s',
                    datefmt='%m-%d %H:%M',
                    handlers=[logging.FileHandler(LOG_FILE, 'w', 'utf-8'), ])
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
console.setFormatter(formatter)
logging.getLogger('').addHandler(console)


def main():
    if not os.path.isfile(STORE_URL_FILE):
        logging.error('Can not find store url file: {}'.format(STORE_URL_FILE))
        sys.exit()

    with open(STORE_URL_FILE) as f:
        store_url = f.readline().strip()

    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('user-agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36"')
    driver = webdriver.Chrome(options=chrome_options)
    driver.set_page_load_timeout(30)
    wait = WebDriverWait(driver, 30)

    items = {}
    driver.get(store_url)

    while True:
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'shopee-page-controller')))
        elements = driver.find_elements_by_css_selector('.shopee-item-card--link')
        for element in elements:
            title = element.get_attribute('title')
            url = element.get_attribute('href')
            logging.info('title={}, url={}'.format(title, url))
            items[url] = title

        latest_inactive_page_number = driver.find_elements_by_css_selector('.shopee-page-controller > .shopee-button-no-outline')[-1].text
        if latest_inactive_page_number == '...':
            latest_inactive_page_number = driver.find_elements_by_css_selector('.shopee-page-controller > .shopee-button-no-outline')[-2].text
        latest_inactive_page_number = int(latest_inactive_page_number)
        active_page_number = int(driver.find_element_by_css_selector('.shopee-page-controller > .shopee-button-solid--primary').text)
        if active_page_number < latest_inactive_page_number:
            page_url = '{}?page={}&sortBy=pop'.format(store_url, active_page_number)
            logging.info('Go to next page: {}'.format(page_url))
            driver.get(page_url)
        else:
            break

    logging.info('Total items: {}'.format(len(items)))

    workbook = xlsxwriter.Workbook('output.xlsx')
    worksheet = workbook.add_worksheet()

    p_title = 1
    p_url = 2
    p_blank = 3
    for k, v in items.items():
        worksheet.write_string('A{}'.format(p_title), '品名')
        worksheet.write_string('B{}'.format(p_title), v)
        worksheet.write_string('A{}'.format(p_url), '網址')
        worksheet.write_string('B{}'.format(p_url), k)
        p_title = p_blank + 1
        p_url = p_title + 1
        p_blank = p_url + 1
    workbook.close()


if __name__ == '__main__':
    main()
