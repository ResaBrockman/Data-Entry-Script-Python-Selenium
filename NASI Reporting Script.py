import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome()


#Things that need to be updated before running:
#Month reporting
month = 'September'
monthnum = 9


#Open Excel Spreadsheet(s)
#Must be named "NASI Import [current month]
#Must contain two worksheets: Journeymen & Apprentices
wb = openpyxl.load_workbook('NASI Import ' + month + '.xlsx')
journey = wb['Journeymen']
appr = wb['Apprentices']
#label values will be going through (class, SSN, worker hours, worker wages)
curRow = 1
locClass = journey['A'+str(curRow)]
ssn = journey['E'+str(curRow)]
hours = journey['F'+str(curRow)]
wage = journey['G'+str(curRow)]
#apprentices
AcurRow = 1
AlocClass = appr['A'+str(curRow)]
Assn = appr['E'+str(curRow)]
Ahours = appr['F'+str(curRow)]
Awage = appr['G'+str(curRow)]

# FUNCTION SECTION

def goToRow(rowNum):
    global curRow
    global locClass
    global ssn
    global hours
    global wage

    curRow = rowNum
    locClass = journey['A'+str(curRow)]
    ssn = journey['E'+str(curRow)]
    hours = journey['F'+str(curRow)]
    wage = journey['G'+str(curRow)]
    print('locClass.value: '+str(locClass.value)+'  ssn.value: '+str(ssn.value))

def AgoToRow(rowNum):
    global AcurRow
    global AlocClass
    global Assn
    global Ahours
    global Awage

    AcurRow = rowNum
    AlocClass = appr['A'+str(AcurRow)]
    Assn = appr['E'+str(AcurRow)]
    Ahours = appr['F'+str(AcurRow)]
    Awage = appr['G'+str(AcurRow)]
    print('AlocClass.value: '+str(AlocClass.value)+'  Assn.value: '+str(Assn.value))

def advanceRow():
    global curRow
    global locClass
    global ssn
    global hours
    global wage
    
    curRow = curRow + 1
    locClass = journey['A'+str(curRow)]
    ssn = journey['E'+str(curRow)]
    hours = journey['F'+str(curRow)]
    wage = journey['G'+str(curRow)]

def AadvanceRow():
    global AcurRow
    global AlocClass
    global Assn
    global Ahours
    global Awage
    
    AcurRow = AcurRow + 1
    AlocClass = appr['A'+str(AcurRow)]
    Assn = appr['E'+str(AcurRow)]
    Ahours = appr['F'+str(AcurRow)]
    Awage = appr['G'+str(AcurRow)]

def goToNextClass():
    
    global locClass
    global curRow
    global ssn
    global hours
    global wage
    
    advanceRow()
    
    while (locClass.value == None or hours.value == None) and curRow < 500:
        advanceRow()
        
    result = locClass.value
    if curRow >= 500:
        result = "curRow too high"
    return result

def AgoToNextClass():
    
    global AlocClass
    global AcurRow
    global Assn
    global Ahours
    global Awage
    
    AadvanceRow()
    
    while (AlocClass.value == None or Ahours.value == None) and AcurRow < 500:
        AadvanceRow()
        
    result = AlocClass.value
    if AcurRow >= 500:
        result = "curRow too high"
    return result

def removeAllWorkers():
    number = int(driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table[1]/tbody/tr/td[1]').text)    
        
    while number > 0:        
        
## UNCOMMENT THIS SECTION IF DELETING DATA WITH NON-ZERO VALUES
##        if driver.find_element_by_id('hoursTotal').text != '0.00':
##            driver.find_element_by_xpath('//*[@id="hours0"]').send_keys(Keys.CONTROL+"a")
##            driver.find_element_by_xpath('//*[@id="hours0"]').send_keys(Keys.DELETE)
##            driver.find_element_by_xpath('//*[@id="hours0"]').send_keys("0.00")
##            driver.find_element_by_xpath('//*[@id="gross0"]').send_keys(Keys.CONTROL+"a")
##            driver.find_element_by_xpath('//*[@id="gross0"]').send_keys(Keys.DELETE)
##            driver.find_element_by_xpath('//*[@id="gross0"]').send_keys("0.00")
        
        #end of section for data with values

        driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table[2]/tbody/tr[4]/td[7]/span/input').click()
        alert_obj = driver.switch_to.alert
        alert_obj.accept()  
        number = int(driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table[1]/tbody/tr/td[1]').text)


def enterWorkers(classCode):
 
    global locClass
    global curRow
    global ssn
    global hours
    global wage

    if classCode == locClass.value:
        print('Entering class: ' + classCode)
    else:
        raise Exception('class mismatch')
   
    while locClass.value == None or locClass.value == classCode:
        while hours.value== "" and (locClass.value == None or locClass.value == classCode):
            advanceRow()
            continue
        driver.find_element_by_id("ssn").send_keys(ssn.value)
        driver.find_element_by_id("hours").send_keys(Keys.CONTROL+"a")
        driver.find_element_by_id("hours").send_keys(Keys.DELETE)
        driver.find_element_by_id("hours").send_keys(hours.value)
        driver.find_element_by_id("gross").send_keys(Keys.CONTROL+"a")
        driver.find_element_by_id("gross").send_keys(Keys.DELETE)
        driver.find_element_by_id("gross").send_keys(wage.value)
        driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table[2]/tbody/tr[2]/td[7]/span/input').click()

        print(driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table[2]/tbody/tr[1]/td').text)
                                     
        #increase amount
        if ssn.value != "":
            advanceRow()

    driver.find_element_by_xpath('//*[@id="content"]/div/div[1]/form/span[1]/input').click()
    print("Done enteringclass: " + classCode)

def enterApprentices(classCode):
 
    global AlocClass
    global AcurRow
    global Assn
    global Ahours
    global Awage
    '''
    if classCode == AlocClass.value:
        print('Entering class: ' + classCode)
    else:
        raise Exception('class mismatch')
   '''
    while AlocClass.value == None or AlocClass.value == classCode:
        while Ahours.value== "" and (AlocClass.value == None or AlocClass.value == classCode):
            AadvanceRow()
            continue
        driver.find_element_by_id("ssn").send_keys(Assn.value)
        driver.find_element_by_id("hours").send_keys(Keys.CONTROL+"a")
        if checkForPopup():
            driver.find_element_by_id("ssn").send_keys(Keys.CONTROL+"a")
            driver.find_element_by_id("ssn").send_keys(Keys.DELETE)
            print("Skipped: " + str(appr['B'+str(AcurRow)].value) + " SSN: " + str(ssn.value))
            AadvanceRow()
            continue
        driver.find_element_by_id("hours").send_keys(Keys.DELETE)
        if checkForPopup():
            driver.find_element_by_id("ssn").send_keys(Keys.CONTROL+"a")
            driver.find_element_by_id("ssn").send_keys(Keys.DELETE)
            print("Skipped: " + str(appr['B'+str(AcurRow)]) + " SSN: " + str(ssn.value))
            AadvanceRow()
            continue
        driver.find_element_by_id("hours").send_keys(Ahours.value)
        driver.find_element_by_id("gross").send_keys(Keys.CONTROL+"a")
        driver.find_element_by_id("gross").send_keys(Keys.DELETE)
        driver.find_element_by_id("gross").send_keys(Awage.value)
        driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table[2]/tbody/tr[2]/td[7]/span/input').click()

        print(driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table[2]/tbody/tr[1]/td').text)
        if driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table[2]/tbody/tr[1]/td').text != "":
            print("Error on: " + str(appr['B'+str(AcurRow)].value) + " SSN: " + str(ssn.value))
                                     
        #increase amount
        if Assn.value != "":
            AadvanceRow()
   
    driver.find_element_by_xpath('//*[@id="content"]/div/div[1]/form/span[1]/input').click()

def checkForPopup():
    try:
        alert_obj = driver.switch_to.alert
        alert_obj.accept()
        return True
    except:
        return False
        
    
# END FUNCTION SECTION



#Start browser & login to NASI
driver.get("https://www.webremit-nasifund.org/webRemittance/login/login")
driver.find_element_by_id("userId").send_keys("REDACTED")
driver.find_element_by_id("password").send_keys("REDACTED")
driver.find_element_by_name("_action_doLogin").click()

#Choose reporting month dropdown
#month = 1 + correct month number
driver.find_element_by_xpath('//*[@id="month"]/option['+str(monthnum+1)+']').click()
driver.find_element_by_name('_action_save').click()


#set initial class and row value
currentClass = locClass
row = 1

#Some Testing - NASI main page table starts at row 1
#element = driver.find_element_by_xpath('//*[@id="content"]/div[2]/div/table/tbody/tr['+str(3)+']/td[2]')
#print(element.text)


#find on main page
y=1
while True:
    try:
        #look for element
        #check class
        element = driver.find_element_by_xpath('//*[@id="content"]/div[2]/div/table/tbody/tr['+str(y)+']/td[2]')
        if element.text == locClass.value:
            #print ("found locClass")
            driver.find_element_by_xpath('//*[@id="content"]/div[2]/div/table/tbody/tr['+str(y)+']/td[4]/form/span/input').click()
            break
        else:
            y=y+1

    except:
        y=1
        #print("call goToNextClass()")
        goToNextClass()
        if curRow >= 500:
            break
        

# Enter Data

#How many workers?
workers = int(driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table[1]/tbody/tr/td[1]').text)

# Remove all workers
removeAllWorkers()

# What class are we entering?
classCode = locClass.value
enterWorkers(classCode)

# loop throught the rest
while curRow < 400:
    goToNextClass()
    print(locClass.value)
    try:
        driver.find_element_by_partial_link_text('669 / '+str(locClass.value)).click()
        removeAllWorkers()
        enterWorkers(locClass.value)
    except:
        print('Could not find link to enter class code: '+str(locClass.value))
print('Done!')


#Enter Apprentices

while AcurRow < 500:
    print(AlocClass.value)

        try:
            driver.find_element_by_partial_link_text('669 / '+str(AlocClass.value)).click()
        except:
            driver.find_element_by_partial_link_text('669/'+str(AlocClass.value)).click()
        removeAllWorkers()
        enterApprentices(AlocClass.value)
    except:
        print('Could not find link to enter class code: '+str(AlocClass.value))
    AgoToNextClass()
print('Done!')


AgoToRow(7)
try:
    driver.find_element_by_partial_link_text('669 / '+"03").click()
except:
    driver.find_element_by_partial_link_text('669/'+"03").click()
enterApprentices("03")


        
