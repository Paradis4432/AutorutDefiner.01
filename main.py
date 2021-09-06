# ---IMPORTS---
import datetime
import time
from datetime import datetime

import dateutil.parser
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.select import Select
import sys

# ---GLOBAL VARIABLES---
link = "https://secure-ausomxbga.crmondemand.com/OnDemand/logon.jsp?type=normal&lang=esn&reason=logoff"
linkPractice = "https://www.tutorialspoint.com/selenium/selenium_automation_practice.htm"
user = "EQUIFAX1/SALESL_SUPERVISOR"
passw = "Sales.2021"
nameExcel = "100_ruts.xlsx"
driver = None

df = pd.read_excel(nameExcel)
workbook = load_workbook(filename=nameExcel)
sheet = workbook.active

start = time.time()
# ---BASIC LOGIC---

try:
    driver = webdriver.Chrome("/opt/homebrew/bin/chromedriver") #macOS
except:
    print("chrome file not found")
else:
    print("chrome file found, using macOS")

try:
    driver = webdriver.Chrome("C:/Users/LucasUser/Downloads/chromedriver.exe")  # Win10
except:
    print("chrome file not found")
else:
    print("chrome file found, using Windows")

if driver is None:
    print("Something went wrong while trying to setup chrome. stopping program")
    sys.exit()

driver.get(link)
driver.implicitly_wait(1)  # gives an implicit wait for 1 second


# driver.maximize_window() # For maximizing window

# noinspection PyMethodMayBeStatic
class web:
    def login(self):
        elem = driver.find_element_by_id("IPT_SC_SignIn")
        elem.send_keys(user)
        elem = driver.find_element_by_id("IPT_SC_Password")
        elem.send_keys(passw)
        driver.find_element_by_id("BTN_SC_Login").click()
        assert "No results found." not in driver.page_source

    def chooseCuentasFromDropdown(self):
        sel = Select(driver.find_element_by_xpath("//select[@name='GlobalSearchMultiField.Search Object']"))
        sel.select_by_visible_text("Cuentas")

    def findRutInSearchBar(self, rut):
        elem = driver.find_element_by_id("GlobalSearchMultiField.Location_Shadow")
        elem.send_keys(rut)
        driver.find_element_by_id("BTN_AP_Global_spcSearch_spcMulti_spcField_Go").click()

    def selectAccountAndActivities(self, id, id2):
        driver.find_element_by_id(id).click()  # if error, rut is null
        driver.find_element_by_id(id2).click()

    def getdate(self):
        time.sleep(1)
        source = driver.page_source
        split1 = source.split("<a class=\"stdFont\">")
        cont = 0
        date = []
        for i in split1:
            if cont == 0:
                cont = cont + 1
                continue
            date.append(i[0:10])
        print("returning date \n" + str(date))
        return date[0]

    def checkDaysDate(self, date):
        dateParsed = dateutil.parser.parse(date, dayfirst=True)
        timeNow = datetime.now() - dateParsed
        print(timeNow)

        if timeNow.days >= 60:
            print("returning asignar")
            return "asignar"
        else:
            print("returning carterizado")
            return "carterizado"


# noinspection PyMethodMayBeStatic
class excel:
    def checkIfBufferFileExists(self):
        try:
            with open('currentRutProcessing.txt', 'x') as f:
                print("currentRutProcessing.txt not found, creating one")
                f.close()
        except:
            pass
        try:
            with open('currentPosEditing.txt', 'x') as f:
                print("currentPosEditing.txt not found, creating one")
                f.close()
        except:
            pass

    def getNextRut(self):
        current = int(self.getCurrentRutProcessing())
        rutList = []
        for value in df['RUT']:
            rutList.append(str(value))
        # print(rutList)

        if len(rutList) - 1 < (rutList.index(current)):
            return -1
        return rutList[rutList.index(current) + 1]

    def getCurrentRutProcessing(self):
        with open('currentRutProcessing.txt', 'r') as f:
            data = f.read()
            f.close()
            return data

    def updateCurrentRutProcessing(self, rut):
        with open('currentRutProcessing.txt', 'w') as f:
            f.write(rut)
        f.close()

    def createStatusColumn(self):
        sheet["H1"] = "STATUS"
        workbook.save(filename=nameExcel)
        with open('currentPosEditing.txt', 'w') as f:
            f.write('H,1')
        f.close()

    def updateStatusColumn(self, status):
        column, row = self.getNextPosForStatus()
        print("current status pos: " + column + str(row))
        sheet[column + str(row)] = status
        workbook.save(filename=nameExcel)

    def getNextPosForStatus(self):
        with open('currentPosEditing.txt', 'r') as f:
            pos = f.readline()
            posSplit = pos.split(",")
        f.close()
        return posSplit[0], int(posSplit[1]) + 1

    def updateNextPosForStatus(self):
        with open('currentPosEditing.txt', 'r') as f:
            pos = f.readline()
            posSplit = pos.split(",")
            pos1 = int(posSplit[1])
            pos1 += 1
            newPos = posSplit[0] + "," + str(pos1)
            f.close()
        with open('currentPosEditing.txt', 'w') as f:
            f.write(newPos)
            f.close()


class autoProcess:
    def getToDate(self):
        web().login()
        web().chooseCuentasFromDropdown()
        currentRut = excel().getCurrentRutProcessing()
        web().findRutInSearchBar(currentRut)
        web().selectAccountAndActivities("_rtid_0", "LNK_HD_ActivityClosedChildList")
        web().getdate()

    def login(self):
        web().login()

    def autoLogic(self):
        web().login()
        web().chooseCuentasFromDropdown()

        excel().createStatusColumn()

        rutList = []
        listaUnica = []

        for value in df['RUT']:
            rutList.append(str(value))
            print("appending value " + str(value))
        print("found a total of " + str(len(rutList)) + " elements")
        for data in range(0, len(rutList)):
            validRut = True
            # update current rut
            excel().updateCurrentRutProcessing(rutList[data])
            # get date from current rut
            currentRut = excel().getCurrentRutProcessing()
            # si el current rut esta en la lista unica return duplicado

            #print("UNIQUE LIST DEBUG: current data in rutList: " + str(rutList[data]))
            #print("UNIQUE LIST DEBUG: current listaUnica: " + str(listaUnica))
            #print("UNIQUE LIST DEBUG: rutList[data] in listaUnica: " + str(bool(rutList[data] in listaUnica)))

            if rutList[data] in listaUnica:
                # devolver duplicado
                excel().updateStatusColumn("duplicado")
                excel().updateNextPosForStatus()
                print("Duplicado")
                print("round finished rut: " + excel().getCurrentRutProcessing())
                print("\n\n")
                continue

            listaUnica.append(excel().getCurrentRutProcessing())
            web().findRutInSearchBar(currentRut)
            try:
                web().selectAccountAndActivities("_rtid_0", "LNK_HD_ActivityClosedChildList")
            except:
                # rut is null
                validRut = False
            if validRut:
                try:
                    date = web().getdate()
                except:
                    excel().updateStatusColumn("asignar")
                    print("Asignar")
                    excel().updateNextPosForStatus()
                    print("round finished rut: " + excel().getCurrentRutProcessing())
                    print("\n\n")
                    continue
                statusOfRut = web().checkDaysDate(date)
                excel().updateStatusColumn(statusOfRut)
            else:
                # rut null return crear
                excel().updateStatusColumn("crear")
                driver.find_element_by_id("GlobalSearchMultiField.Location_Shadow").clear()
                print("Creando Rut")
            excel().updateNextPosForStatus()
            print("round finished rut: " + excel().getCurrentRutProcessing())
            print("\n\n")



            # logic based on date of current rut
            # if rut is null return crear
            # if rut is not null and date > 60 days return asignar
            # if rut is not null and date < 60 days return carterizado
            # if rut is not null and date is

    # loop toma la lista de RUTs del excel
    # actualiza current RUT
    # llama a getDate
    # logica en base a getDate

    # si el rut no esta creado return "crear"
    # si el dato esta creado y la fecha es mayor a 60 dias return "asignar"
    # si la fecha es menos a 60 dias return "carterizado"


# print(excel().createStatusColumn())
autoProcess().autoLogic()
# autoProcess().login()
# excel().checkIfBufferFileExists()

end = time.time()

print("Runtime of the program is: " + str(end - start))
