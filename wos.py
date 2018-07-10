import os
import time
import string
import random
import subprocess as sp
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from PIL import Image
from PIL import ImageEnhance

random.seed(1024)

os.environ["TESSDATA_PREFIX"] = os.path.join(os.getcwd(), "tesseract-4.0.0-alpha")



def orc(img, tesseract=os.path.join(os.getcwd(), 'tesseract-4.0.0-alpha/tesseract.exe')):
    # 识别简单验证码
    if isinstance(img, str):
        img = Image.open(img)
    gray = img.convert('L')
    contrast = ImageEnhance.Contrast(gray)
    ctgray = contrast.enhance(3.0)
    bw = ctgray.point(lambda x: 0 if x < 1 else 255)
    bw.save('captcha_threasholded.png')
    process = sp.Popen(
        [tesseract, 'captcha_threasholded.png', 'out', '--psm 7', '--tessdata-dir ' + os.path.dirname(tesseract)],
        shell=True
    )
    process.wait()
    with open('out.txt', 'r') as f:
        words = ''.join(list(f.readlines())).rstrip()
    words = ''.join(c for c in words if c in string.ascii_letters + string.digits).lower()
    return words


def search_settings(driver, university, start, end):
    # 点选 database
    driver.find_element_by_class_name('select2-selection__arrow').click()
    driver.find_element_by_id(
        'select2-databases-results'
    ).find_elements_by_tag_name(
        'li'
    )[1].click()

    # 点选搜索类别
    if driver.find_element_by_id('select2-select1-container').text != 'Organization-Enhanced':
        driver.find_elements_by_class_name('select2-selection__arrow')[1].click()
        driver.find_element_by_id(
            'select2-select1-results'
        ).find_elements_by_tag_name('li')[10].click()

    # input
    searchinput = driver.find_element_by_id('value(input1)')
    driver.execute_script('arguments[0].value = arguments[1]', searchinput, university)

    # 时间
    timespan = driver.find_element_by_id('timespan').find_element_by_class_name('select2-selection__rendered')
    if timespan.text != 'Custom year range':
        timespan.click()
        driver.find_element_by_class_name(
            'select2-results__options'
        ).find_elements_by_tag_name('li')[6].click()
    yeararrows = driver.find_element_by_class_name('timespan_custom').find_elements_by_class_name('select2-selection')
    for i in range(len(yeararrows)):
        yeararrows[i].click()
        yearoptions = driver.find_elements_by_class_name('select2-results__option')
        for yop in yearoptions:
            if yop.text == '{0}'.format([start, end][i]):
                yop.click()
                break

    # more settings
    if 'fa-caret-down' in driver.find_element_by_id('settings-arrow').get_attribute('class'):
        driver.find_element_by_id('settings-arrow').click()

    checkboxes = driver.find_elements_by_class_name('wos-style-checkbox')
    for i in range(3):
        if not checkboxes[i].get_property('checked'):
            checkboxes[i].click()
    for i in range(3, 8):
        if checkboxes[i].get_property('checked'):
            checkboxes[i].click()


def start_search(driver):
    searchbutton = driver.find_element_by_class_name('searchButton').find_element_by_tag_name('button')
    searchbutton.click()


def check_search(driver):
    try:
        driver.find_element_by_id('noRecordsDiv')
        return False
    except NoSuchElementException:
        return True


def analysisresult(driver):
    # analysis result
    WebDriverWait(driver, 4).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'create-cite-report'))
    )
    driver.find_element_by_class_name(
        'create-cite-report'
    ).find_element_by_class_name(
        'snowplow-analyze-results'
    ).click()


def analysispage(driver):
    # 选择显示数量
    driver.find_element_by_id('select2-refineMaxRows-container').click()
    shownums = driver.find_element_by_id(
        'select2-refineMaxRows-results'
    ).find_elements_by_class_name('select2-results__option')
    for nop in shownums:
        if nop.text == '500':
            nop.click()
            break
    # 数据
    evenrows = driver.find_element_by_class_name(
        'RA-NEWresultsSectionTable'
    ).find_elements_by_class_name('RA-NEWRAresultsEvenRow')
    oddrows = driver.find_element_by_class_name(
        'RA-NEWresultsSectionTable'
    ).find_elements_by_class_name('RA-NEWRAresultsOddRow')
    rows = evenrows + oddrows
    result = list()
    for row in rows:
        result.append(
            [row.find_elements_by_tag_name('td')[1].text.strip(), row.find_elements_by_tag_name('td')[2].text.strip()]
        )
    return result


def crawler(driver, university, start, end, filepath):
    page = 'https://apps.webofknowledge.com/'
    driver.maximize_window()
    for year in range(start, end + 1):
        driver.get(page)
        # 设置搜索选项
        search_settings(driver, university, year, year)
        # 开始搜索
        start_search(driver)
        if not check_search(driver):
            continue
        time.sleep(1)
        # 科目分析页面
        analysisresult(driver)
        # 下载数据
        result = analysispage(driver)
        with open(filepath, 'ab') as f:
            f.write('\n'.join('{0},{1},{2},{3}'.format(university, year, x[0], x[1]) for x in result).encode('utf-8'))
            f.write('\n'.encode('utf-8'))
        time.sleep(random.randint(0, 1))


profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.dir', os.path.join(os.getcwd(), 'filedata'))
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/x--tagged; charset=utf-8')
driver = webdriver.Firefox(firefox_profile=profile, executable_path=r'./geckodriver.exe')


university = pd.read_table('university.txt', header=0)

for row in university.loc[404:600, ].iterrows():
    print('{0}: {1}'.format(row[0], row[1]['学校名称']))
    crawler(
        driver,
        university=row[1]['英文名称'],
        start=2001,
        end=2017,
        filepath=os.path.join('./filedata', row[1]['学校名称'] + '.csv')
    )

crawler(driver, university='Peking University', start=2002, end=2017, filepath=os.path.join('./filedata', '北京大学.csv'))