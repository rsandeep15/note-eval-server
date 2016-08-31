import warnings,requests, xml.etree.ElementTree as ET, openpyxl
import urllib
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class RealEstate:
    # Finds all the columns with the vital information needed to query APIs
    def findColumns(self, sheet):
        mapping = {}
        keys = ["Street Address", "Zip", "Zillow", "Trulia"]
        for col in range(1, sheet.max_column):
            cell = sheet.cell(row = 3, column = col)
        # Trimming possible white spaces on edge before comparison
            if cell.value is not None and cell.value.strip() in keys:
                mapping[cell.value.strip()] = col
        # Returns a dictionary of the mappings for look up
        return mapping

    # Queries Zillow API for a price of a home given its address and zipcode
    def getZillowPrice(self, address, zipcode):
        ZILLOW_API_KEY = 'X1-ZWz1fcy5pfuk97_5upah'
        ZILLOW_URL = 'http://www.zillow.com/webservice/GetSearchResults.htm'

        data = {'zws-id':ZILLOW_API_KEY, 'address':address,
            'citystatezip':zipcode}

        r = requests.post(ZILLOW_URL, params=data)

        root = ET.fromstring(r.text)
        element = root.find('./response/results/result/zestimate/amount')
        if element is None:
            return "Not found"
        else:
            if (str(element.text) == "None"):
                return "Not found"
            else:
                return "$" + str(element.text)

    # Queries Trulia API for a price of a home given its address and zipcode
    def getTruliaPrice(self, driver, address, zipcode):
        TRULIA_URL = "http://trulia.com"
        wait = WebDriverWait(driver, 20)
        driver.get(TRULIA_URL)
        element = wait.until(EC.presence_of_element_located ((By.ID,"searchbox_form_location")))
        element.clear()
        element.send_keys(str(address) + " " + str(zipcode))
        element.send_keys(Keys.RETURN)
        price = "Not Found"
        try:
            content = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.typeDeemphasize > span:nth-child(1) > span:nth-child(2)")))
            price = "$" + content.text
        finally:
            return price

    def run(self, publicURL):
        # Pull the remote file into a local directory
        FILE_PATH = "data/data.xlsx"
        urllib.urlretrieve(publicURL, FILE_PATH)

        warnings.simplefilter('ignore')
        wb = openpyxl.load_workbook('data/data.xlsx')
        sheet = wb.get_sheet_by_name('Jan')

        columnsMap = self.findColumns(sheet);
        print columnsMap
        address_col = columnsMap["Street Address"]
        zipcode_col = columnsMap["Zip"]

        zillow_col = columnsMap["Zillow"]
        trulia_col = columnsMap["Trulia"]
        #
        # driver = webdriver.Firefox()
        # for x in range(4, sheet.max_row):
        #     address = sheet[address_col + str(x)].value
        #     zipcode = sheet[zipcode_col + str(x)].value
        #
        # #Skip invalid addresses and zipcodes
        # if address is None or zipcode is None:
        #     continue
        #
            zillowPrice = getZillowPrice(address, zipcode)
        #     truliaPrice = getTruliaPrice(driver, address, zipcode)
        #
        if zillowPrice != "Not found":
            sheet[zillow_col + str(x)] = zillowPrice
        # if truliaPrice != "Not found":
        #     sheet[trulia_col + str(x)] = truliaPrice
        #
        output = str(address).ljust(30) + str(zipcode).ljust(15)
        # output += str(zillowPrice).ljust(30) + str(truliaPrice)
        print output
        #
        # wb.save('data/data.xlsx')
