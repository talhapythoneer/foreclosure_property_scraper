from selenium import webdriver
from time import sleep
from selenium.webdriver.chrome.options import Options
from shutil import which
from scrapy import Selector
import xlsxwriter


URL = "https://www.foreclosure.com/listing/search?q=California&lc=foreclosure&lc=preforeclosure&pt=mf&lc=foreclosure&loc=California&view=list"
email = ""
password = ""

workbook = xlsxwriter.Workbook('Data.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, "Price")
worksheet.write(0, 1, "Beds")
worksheet.write(0, 2, "Baths")
worksheet.write(0, 3, "Area")
worksheet.write(0, 4, "PropType")
worksheet.write(0, 5, "listingType")
worksheet.write(0, 6, "Address")
rowN = 1

url = URL

if "&ps=" not in url:
    url = url + "&ps=100"


def botInitialization():
    # Initialize the Bot
    chromeOptions = Options()
    # chromeOptions.add_argument("--headless")
    chromePath = which("chromedriver")
    driver = webdriver.Chrome(executable_path=chromePath, options=chromeOptions)
    driver.maximize_window()
    return driver


driver = botInitialization()
print("Logging in... ")

driver.get("https://www.foreclosure.com/login")
emailBox = driver.find_element_by_css_selector("input[placeholder='Email']")
passBox = driver.find_element_by_css_selector("input[placeholder='Password']")
emailBox.send_keys(email)
passBox.send_keys(password)
loginButton = driver.find_element_by_css_selector("input#btnLoginUsernamePassword")
driver.execute_script("arguments[0].click();", loginButton)
sleep(2)

print("Getting the main page... ")
driver.get(url)
sleep(2)
count = 1
while True:
    response = Selector(text=driver.page_source)
    properties = response.css("div.listingRow")
    for prop in properties:
        price = prop.css("span.tdprice > strong::text").extract_first()
        address = prop.css("a.address > span::text").extract()
        address = ", ".join(address)
        beds = prop.css("div.bedroomsbox::text").extract_first()
        baths = prop.css("div.barhroomsbox::text").extract_first()
        area = prop.css("div.sizebox::text").extract_first()
        propertyType = prop.css("div.ptypebox::text").extract_first()
        listingType = prop.css("div.messajeType::text").extract()[3].strip()

        worksheet.write(rowN, 0, price)
        worksheet.write(rowN, 1, beds)
        worksheet.write(rowN, 2, baths)
        worksheet.write(rowN, 3, area)
        worksheet.write(rowN, 4, propertyType)
        worksheet.write(rowN, 5, listingType)
        worksheet.write(rowN, 6, address)
        rowN += 1

    nextPage = driver.find_element_by_css_selector("a#pageNextBottom")
    print("Pages Extracted: " + str(count))
    count += 1
    if nextPage.get_attribute("data-nextpage") == "-1":
        break
    else:
        driver.execute_script("arguments[0].click();", nextPage)
        sleep(5)

workbook.close()
driver.close()
driver.quit()
