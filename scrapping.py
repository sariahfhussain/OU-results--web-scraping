import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from xlwt import Workbook
import pandas as pd

class UniversityResult:
    result_link = "https://www.osmania.ac.in/res07/20240511.jsp"
    pre_link = result_link + "?mbstatus&htno="
    row = 1
    collegeCodes = ["1604"]
    fieldCodes = ["733"]
    year = "23"

    def __init__(self):
        self.wb, self.sheet = self.getNewExcelSheet()
        self.initiateFindingResult()

    def initiateFindingResult(self):
        options = Options()
        options.headless = True
        driver = webdriver.Chrome(options=options)

        for fieldCode in self.fieldCodes:
            for collegeCode in self.collegeCodes:
                # Iterate through hall ticket numbers 1 to 180
                for index in range(1, 181):
                    hall_ticket = collegeCode + self.year + fieldCode + str(index).zfill(3)
                    self.findResult(driver, fieldCode, collegeCode, hall_ticket, index)

        excelFileName = f"Results__{fieldCode}_{self.year}.xls"
        self.wb.save(excelFileName)
        driver.quit()

    def findResult(self, driver, fieldCode, collegeCode, hall_ticket, index):
        try:
            resultLink = self.pre_link + hall_ticket
            driver.get(resultLink)
            time.sleep(0.4)  # Add a small delay to allow the page to load

            marks_element = driver.find_element(By.CSS_SELECTOR, "#AutoNumber5 tr:nth-child(3) td:nth-child(2)")
            marks = marks_element.text

            name_element = driver.find_element(By.CSS_SELECTOR, "#AutoNumber3 tr:nth-child(3) td:nth-child(2)")
            name = name_element.text

            # Process and write data to Excel sheet
            self.sheet.write(self.row, 0, self.row)
            self.sheet.write(self.row, 2, str(collegeCode))
            self.sheet.write(self.row, 3, str(fieldCode))
            self.sheet.write(self.row, 4, hall_ticket[-3:])
            self.sheet.write(self.row, 5, marks)
            self.sheet.write(self.row, 6, name)

            self.row += 1
            print(f"Row Number: {self.row - 1}")
            print(f"{name} {hall_ticket} {marks}")

        except Exception as e:
            print(f"Error processing hall ticket {hall_ticket}: {e}")

    def getNewExcelSheet(self):
        wb = Workbook()
        print("Creating Excel sheet...")
        sheet1 = wb.add_sheet('Sheet 1')
        sheet1.write(0, 0, 'S.No')
        sheet1.write(0, 1, 'Rank')
        sheet1.write(0, 2, 'Code')
        sheet1.write(0, 3, 'Field')
        sheet1.write(0, 4, 'R.No')
        sheet1.write(0, 5, "CGPA")
        sheet1.write(0, 6, "Name")
        return wb, sheet1

# Initialize and run the script
UniversityResult()
