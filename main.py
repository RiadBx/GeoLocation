import time
import pandas as pd
import lxml.html
import selenium
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from selenium import webdriver

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
import re


def scrape(store_id, title, address, suburb, state, phone, postcode):
    info = title + " " + address + " " + suburb + " " + state

    str = f'https://www.google.com/maps/search/{info}'
    driver.get(str)
    timeout = 0

    try:
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'IdOfMyElement')))
        print("Page is ready!")
    except TimeoutException:
        print("Loading took some time!")

    time.sleep(5)
    try:
        tree = lxml.html.fromstring(driver.page_source)
        results = tree.xpath(
            '/html/body/div[3]/div[9]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]')
        for text in results:
            print(text.text_content())

        loc = re.findall(r"@[-]*\d+.\d+,[-]*\d+.\d+", driver.current_url)[0].replace("@", "").split(",")
        latitude, langitude = loc[0], loc[1]

        region_dict = {'id': store_id,
                       'data.title': title,
                       'data.address': address,
                       'data.suburb': suburb,
                       'data.state': state,
                       'data.postcode': postcode,
                       'data.phone': phone,
                       'data.location': {'latitude': latitude, 'longitude': langitude},
                       'latitude': latitude,
                       'longitude': langitude,
                       "error": "No"
                       }

        return region_dict
    except selenium.common.exceptions.InvalidArgumentException:

        loc = re.findall(r"@[-]*\d+.\d+,[-]*\d+.\d+", driver.current_url)[0].replace("@", "").split(",")
        latitude, langitude = loc[0], loc[1]

        region_dict = {'id': store_id,
                       'data.title': title,
                       'data.address': address,
                       'data.suburb': suburb,
                       'data.state': state,
                       'data.postcode': postcode,
                       'data.phone': phone,
                       'data.location': {'latitude': latitude, 'longitude': langitude},
                       'latitude': latitude,
                       'longitude': langitude,
                       'error': "Yes"
                       }
        return region_dict


def excel_file_read(file_name):
    dataframe = pd.read_excel(file_name, skiprows=1)
    data = dataframe.to_dict()
    store_id = list(data['id'].values())
    title = list(data['data.title'].values())
    address = list(data['data.address'].values())
    suburb = list(data['data.suburb'].values())
    state = list(data['data.state'].values())
    phone = list(data['data.phone'].values())
    postcode = list(data['data.postcode'].values())

    link = []
    length = len(title)
    for i in range(0, length):
        location = title[i].rstrip() + " " + address[i].rstrip() + " " + suburb[i].rstrip() + " " + state[i].rstrip()
        link.append(location)
    return store_id, title, address, suburb, state, phone, postcode


def write_excel(data):
    df = pd.DataFrame(data=data)

    # convert into excel
    df.to_excel("final.xlsx", index=False)


if __name__ == '__main__':

    # write_excel()

    driver = webdriver.Firefox()
    hit_count = 0
    store_id, title, address, suburb, state, phone, postcode = excel_file_read('ryco-distributors-au (1).xlsx')
    data = []
    for i in range(len(title[:50])):
        data.append(
            scrape(store_id[i].rstrip(), title[i].rstrip(), address[i].rstrip(), suburb[i].rstrip(), state[i].rstrip(),
                   phone[i], postcode[i]))
    write_excel(data)
    driver.quit()



