import os
import time
import string
import random
import subprocess as sp
from selenium import webdriver
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
    startyear = driver.find_element_by_class_name('timespan_custom').find_element_by_class_name('startyear')
    while True:
        if startyear.get_property('value') == '{0}'.format(start):
            break
        else:
            startyear.send_keys('{0}'.format(start))
    endyear = driver.find_element_by_class_name('timespan_custom').find_element_by_class_name('endyear')
    while True:
        if endyear.get_property('value') == '{0}'.format(end):
            break
        else:
            endyear.send_keys('{0}'.format(end))

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


def analysisresult(driver):
    # analysis result
    driver.find_element_by_class_name(
        'create-cite-report'
    ).find_element_by_class_name(
        'snowplow-analyze-results'
    ).click()


def analysispage(driver):
    # 选择显示数量
    shownum = driver.find_element_by_id('refineMaxRows')
    while True:
        if shownum.get_property('value') == '500':
            break
        else:
            shownum.send_keys('500')
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


crawler(driver, university='Peking University', start=2011, end=2011, filepath=os.path.join('./filedata', '北京大学.csv'))