# ---IMPORTS---
import datetime
import time
from datetime import datetime

import dateutil.parser
import pandas as pd
import selenium.common.exceptions
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.select import Select
import sys

# -- COPYRIGHT --
__author__ = "David Alexander Avilés Brun, Lucas Ezequiel Rodriguez"
__copyright__ = "Copyright (C) 2021 by David & Lucas"
__license__ = "All rights reserved."
__version__ = "Beta_Salesland"

# ---GLOBAL VARIABLES---
link = "https://secure-ausomxbga.crmondemand.com/OnDemand/logon.jsp?type=normal&lang=esn&reason=logoff"
linkPractice = "https://www.tutorialspoint.com/selenium/selenium_automation_practice.htm"
user = "EQUIFAX1/SALESL_SUPERVISOR"
passw = "Sales.2021"
nameExcel = "libro4.xlsx"
driver = None

df = pd.read_excel(nameExcel)
workbook = load_workbook(filename=nameExcel)
sheet = workbook.active

start = time.time()
# ---BASIC LOGIC---

try: driver = webdriver.Chrome("/opt/homebrew/bin/chromedriver")  # macOS
except: print("chrome file not found with macOS path")
else: print("chrome file found, using macOS")

try: driver = webdriver.Chrome("C:/Users/LucasUser/Downloads/chromedriver.exe")  # Win10
except: print("chrome file not found with windows path")
else: print("chrome file found, using Windows")

if driver is None:
    print("Something went wrong while trying to setup chrome. stopping program")
    sys.exit()

driver.get(link)
driver.implicitly_wait(0.5)  # gives an implicit wait for 1 second


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
        time.sleep(0.1)
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

    def checkOwner(self):
        element = driver.find_element_by_id("_rtid_1")
        val = element.get_attribute("innerText")
        return val == "SALESL_FFVV"


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
        print("UPDATE STARTING")
        column, row = self.getNextPosForStatus()
        print("current status pos: " + column + str(row))
        sheet[column + str(row)] = status
        print("UPDATE ASIGNING STATUS")

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

        cont = 0

        for value in df['RUT']:
            rutList.append(str(value))
            # print("appending value " + str(value))
        print("found a total of " + str(len(rutList)) + " elements")
        for data in range(0, len(rutList)):
            print("STARTING")
            validRut = True
            # update current rut
            excel().updateCurrentRutProcessing(rutList[data])
            print("UPDATED")
            # get date from current rut
            currentRut = excel().getCurrentRutProcessing()
            print("ASSIGNED CURRENT RUT")
            # si el current rut esta en la lista unica return duplicado

            # print("UNIQUE LIST DEBUG: current data in rutList: " + str(rutList[data]))
            # print("UNIQUE LIST DEBUG: current listaUnica: " + str(listaUnica))
            # print("UNIQUE LIST DEBUG: rutList[data] in listaUnica: " + str(bool(rutList[data] in listaUnica)))

            if rutList[data] in listaUnica:
                print("RUTLIST[data] IN LISTAUNICA")
                # devolver duplicado
                excel().updateStatusColumn("duplicado")
                excel().updateNextPosForStatus()
                print("Duplicado")
                print("round finished rut: " + excel().getCurrentRutProcessing())
                print("\n\n")
                continue

            print("APPEND LISTA UNICA")
            listaUnica.append(excel().getCurrentRutProcessing())
            print("APPEND DONE")
            print("SEARCHING")
            web().findRutInSearchBar(currentRut)
            print("READY")
            try:
                web().selectAccountAndActivities("_rtid_0", "LNK_HD_ActivityClosedChildList")  # click en ir y cuenta
                print("CLCIKED IR AND CUENTA")
            except:
                # rut is null
                validRut = False
            try:
                if web().checkOwner():
                    print("OPENING")
                    excel().updateStatusColumn("SALESL")
                    print("UPDATED COLUMS SALESL")
                    excel().updateNextPosForStatus()

                    continue
            except selenium.common.exceptions.NoSuchElementException:
                driver.refresh()

            if validRut:
                try:
                    date = web().getdate()
                except:
                    print("UPDATING ASIGNAR")
                    excel().updateStatusColumn("asignar")
                    print("Asignar")
                    excel().updateNextPosForStatus()
                    print("round finished rut: " + excel().getCurrentRutProcessing())
                    print("\n\n")
                    continue
                statusOfRut = web().checkDaysDate(date)
                print("UPDATING STATIS OF COLUMN")
                excel().updateStatusColumn(statusOfRut)
                print("UPDATED")
            else:
                try:
                    # rut null return crear
                    print("UPDATING CREAR")
                    excel().updateStatusColumn("crear")
                    print("UPDATED CREAR")
                    print("FINDING ELEMENT")
                    driver.find_element_by_id(
                        "GlobalSearchMultiField.Location_Shadow").clear()  # borra barra de busqueda
                    print("FOUND")
                    print("Creando Rut")
                except selenium.common.exceptions.NoSuchElementException:
                    driver.refresh()
            print("UPDATING FOR NEXT POS")
            excel().updateNextPosForStatus()
            print("FINISHED")
            print("round finished rut: " + excel().getCurrentRutProcessing())
            print("\n\n")
            # contador
            # si contador > 100 guardar
            if cont > 100:
                cont = 0
                print("SAVING")
                workbook.save(filename=nameExcel)
                print("SAVED")
            cont += 1
            print("CURRENT CONT: " + str(cont))

            # logic based on date of current rut
            # if rut is null return crear
            # if rut is not null and date > 60 days return asignar
            # if rut is not null and date < 60 days return carterizado
            # if rut is not null and date is
        print("SAVING")
        workbook.save(filename=nameExcel)
        print("SAVED")

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

f = '{0:.3g}'.format(int(end - start) / 60)
print("Runtime of the program is: " + str(f))
