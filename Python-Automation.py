# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from unit_test import AutomationTest

df = pd.read_excel(r'')

'Opening Wordpress--------------------------------------------------------------'
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(chrome_options=chrome_options)
driver.get('')

loginbutton = driver.find_element_by_xpath('//*[@id="jetpack-sso-wrap"]/a[1]')
loginbutton.click()

'Username and Password inputting-------------------------------------------------'
UserIN = input('Enter your Username: ')
username = driver.find_element_by_xpath('//*[@id="user_login"]')
username.send_keys(UserIN)
time.sleep(1)

PassIN = input('Enter your Password: ')
password = driver.find_element_by_xpath('//*[@id="user_pass"]')
password.send_keys(PassIN)

'Test if login parameters were inputted correctly'
AutomationTest.test_login(UserIN,PassIN,username,password)

login = driver.find_element_by_xpath('//*[@id="wp-submit"]')
login.click()

'Start Cell integration / Main Loop ----------------------------------------------'
row = 0
while (1):

    programTitleCell = df.iloc[row,0]
    
    'Break loop after last event'
    if programTitleCell == 'END':
        break

    'Identify cell columns'
    descriptionCell = df.iloc[row,1]
    startTimeCell = df.iloc[row,2]
    endTimeCell = df.iloc[row,3]
    dateCell = df.iloc[row,4]
    repeatDateCell = df.iloc[row,5]
    endRepeatDateCell = df.iloc[row,6]
    locationCell = df.iloc[row,7]
    organizerCell = df.iloc[row,8]
    programTypeCell = df.iloc[row,9]
    programTypeCell_B = df.iloc[row,10]
    languageCell = df.iloc[row,11]
    linkCell = df.iloc[row,12]
    
    'Event Creation'
    time.sleep(0.5)
    title = driver.find_element_by_xpath('//*[@id="title"]')
    title.send_keys(programTitleCell)

    textbutton = driver.find_element_by_xpath('//*[@id="content-html"]')
    textbutton.click()
    
    'Insert pre-written description template with event name, description and link'
    description = driver.find_element_by_xpath('//*[@id="content"]')
    description.send_keys('<strong>Description:</strong>')
    description.send_keys(Keys.ENTER)
    description.send_keys(Keys.ENTER)
    description.send_keys(descriptionCell)
    
    'Gather description up to this point for testing'
    totalDescription = description.get_attribute("value")
    
    description.send_keys(Keys.ENTER)
    description.send_keys(Keys.ENTER)
    description.send_keys('<strong>Registration:</strong>')
    description.send_keys(Keys.ENTER)
    description.send_keys(Keys.ENTER)
    description.send_keys('To access this program click <a href='+linkCell+'>here.</a>')
    
    'Test if description from excel is contained within the field'
    AutomationTest.test_description(descriptionCell,totalDescription)
    
    sdate = driver.find_element_by_xpath('//*[@id="mec_start_date"]')
    sdate.send_keys(dateCell[:10])
    
    edate = driver.find_element_by_xpath('//*[@id="mec_end_date"]')
    edate.send_keys(dateCell[:10])
    
    'Test correct event date inputting'
    AutomationTest.test_startEndDate(sdate,edate,dateCell)
    
    'Function for time dropdown menus'
    def calendarDropdown():
        
        'In order to combine amHours and pmHours into one for loop, 12 must be placed as the last index in the pmHoursArray'
        'This is due to 12 being the last indexed xpath value on wordpress'
        'A placeholder of 00 is used in the amHoursArray to ensure we have equal length arrays'
        '00 will never be selected by amHours since wordpress does not directly input 24 hour values into the dropdown menu'
        amHoursArray = ['01','02','03','04','05','06','07','08','09','10','11','00']
        pmHoursArray = ['13','14','15','16','17','18','19','20','21','22','23','12']
        minsArray = ['00','05','10','15','20','25','30','35','40','45','50','55']
        
        'Since both amHoursArray and pmHoursArray have the same length, we can index only amHoursArray for both possibilities'
        i=0
        for i in range(len(amHoursArray)):
            if (startTimeCell[0:2]) == amHoursArray[i]:
                starttimeH = driver.find_element_by_xpath('//*[@id="mec_start_hour"]/option['+str(i+1)+']')
                starttimeH.click()
                starttimeK = driver.find_element_by_xpath('//*[@id="mec_start_ampm"]/option[1]')
                starttimeK.click()
                break
            
            elif (startTimeCell[0:2]) == pmHoursArray[i]:
                starttimeH = driver.find_element_by_xpath('//*[@id="mec_start_hour"]/option['+str(i+1)+']')
                starttimeH.click()
                starttimeK = driver.find_element_by_xpath('//*[@id="mec_start_ampm"]/option[2]')
                starttimeK.click()
                break
            
        for i in range(len(minsArray)):
            if (startTimeCell[3:5] == minsArray[i]):
                 starttimeH = driver.find_element_by_xpath('//*[@id="mec_start_minutes"]/option['+str(i+1)+']')
                 starttimeH.click()
                 break
                
            
        '-------------------------------------------------------------------------------------------------'
        
        for i in range(len(amHoursArray)):
            if (endTimeCell[0:2]) == amHoursArray[i]:
                starttimeH = driver.find_element_by_xpath('//*[@id="mec_end_hour"]/option['+str(i+1)+']')
                starttimeH.click()
                starttimeK = driver.find_element_by_xpath('//*[@id="mec_end_ampm"]/option[1]')
                starttimeK.click()
                break
            
            elif (endTimeCell[0:2]) == pmHoursArray[i]:
                starttimeH = driver.find_element_by_xpath('//*[@id="mec_end_hour"]/option['+str(i+1)+']')
                starttimeH.click()
                starttimeK = driver.find_element_by_xpath('//*[@id="mec_end_ampm"]/option[2]')
                starttimeK.click()
                break
            
        for i in range(len(minsArray)):
            if (endTimeCell[3:5] == minsArray[i]):
                 starttimeH = driver.find_element_by_xpath('//*[@id="mec_end_minutes"]/option['+str(i+1)+']')
                 starttimeH.click()
                 break
           
            
    calendarDropdown()
    
    'Switch wordpress tabs and add location and organizer'
    tabswitchLocation = driver.find_element_by_xpath('//*[@id="mec_metabox_details"]/div[2]/div/div[1]/a[5]')
    tabswitchLocation.click()
    
    locationDrop = driver.find_element_by_xpath('//*[@id="select2-mec_location_id-container"]')
    locationDrop.click()
    
    locationSet = driver.find_element_by_xpath('/html/body/span/span/span[1]/input')
    locationSet.send_keys(locationCell)
    locationSet.send_keys(Keys.ENTER)
    
    tabswitchOrganizer = driver.find_element_by_xpath('//*[@id="mec_metabox_details"]/div[2]/div/div[1]/a[7]')
    tabswitchOrganizer.click()
    
    organizerDrop = driver.find_element_by_xpath('//*[@id="select2-mec_organizer_id-container"]')
    organizerDrop.click()
    
    organizerInput = driver.find_element_by_xpath('/html/body/span/span/span[1]/input')
    organizerInput.send_keys(organizerCell)
    organizerInput.send_keys(Keys.ENTER)
    
    'Function for Identifying primary and secondary program categories, as well as their associated colours'
    def ProgramCategories():
        
        i=0
        programCategoryArray = ['Online','In Centre','Outdoor']
        programCategoryArrayXpath = ['//*[@id="in-mec_category-166"]','//*[@id="in-mec_category-168"]','//*[@id="in-mec_category-167"]']
        programCategoryArray_B = ['EarlyON Outdoor','Exploring Nature','Exploring the Community','Family Time In Centre','Family Time Online','Healthy Eating In Centre','Healthy Eating Online','Indigenous-Led In Centre','Indigenous-Led Online','Mother Goose In Centre','Mother Goose Online','Music and Movement In Centre','Music and Movement Online','Parent Group & Resources In Centre','Parent Group & Resources Online','Story Time In Centre','Story Time Online']
        programCategoryArrayXpath_B = ['//*[@id="in-mec_category-250"]','//*[@id="in-mec_category-251"]','//*[@id="in-mec_category-249"]','//*[@id="in-mec_category-161"]','//*[@id="in-mec_category-316"]','//*[@id="in-mec_category-162"]','//*[@id="in-mec_category-317"]','//*[@id="in-mec_category-163"]','//*[@id="in-mec_category-318"]','//*[@id="in-mec_category-159"]','//*[@id="in-mec_category-319"]','//*[@id="in-mec_category-160"]','//*[@id="in-mec_category-320"]','//*[@id="in-mec_category-164"]','//*[@id="in-mec_category-321"]','//*[@id="in-mec_category-158"]','//*[@id="in-mec_category-322"]']
        programColourArray = ['#a8507c','#a8507c','#a8507c','#a8507c','#a8507c','#906fce','#906fce','#11637c','#11637c','#8ce580','#8ce580','#44bb99','#44bb99','#b391c9','#b391c9','#77aadd','#77aadd']
        
        for i in range(len(programCategoryArray)):
            if (programTypeCell) == programCategoryArray[i]:
                category1 = driver.find_element_by_xpath(programCategoryArrayXpath[i])
                category1.click()
                break
            
        for i in range(len(programCategoryArray_B)):
            if (programTypeCell_B) == programCategoryArray_B[i]:
                category2 = driver.find_element_by_xpath(programCategoryArrayXpath_B[i])
                category2.click()
                colour = driver.find_element_by_xpath('//*[@id="mec_metabox_color"]/div[2]/div/div[1]/div/button')
                colour.click()
                colourSearch = driver.find_element_by_xpath('//*[@id="mec_event_color"]')
                colourSearch.clear()
                colourSearch.send_keys(programColourArray[i])
                break
            
    ProgramCategories()
    
    'Since there may be events that only occur once, we can ignore the extra steps required to repeat an event below'
    if (repeatDateCell) == 'None':
        pass
        
    else:
        
        tabswitch1 = driver.find_element_by_xpath('//*[@id="mec_metabox_details"]/div[2]/div/div[1]/a[2]')
        tabswitch1.click()
        
        eventRepeat = driver.find_element_by_xpath('//*[@id="mec_repeat"]')
        eventRepeat.click()
        
        certainWeekdays = driver.find_element_by_xpath('//*[@id="mec_repeat_type"]/option[4]')
        certainWeekdays.click()
        
        'Function for the days of the week that the event occurs on'
        def weekdays():
            
            weekdaysArray = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
            
            'The xpath index has an offset of 2 since this is how its represented in wordpress'
            i=0
            for i in range(len(weekdaysArray)):
                if (repeatDateCell) == weekdaysArray[i]:
                    weekday = driver.find_element_by_xpath('//*[@id="mec_repeat_certain_weekdays_container"]/label['+str(i+2)+']')
                    weekday.click()
                    break
        
        weekdays()
        
        onButton = driver.find_element_by_xpath('//*[@id="mec_repeat_ends_date"]')
        onButton.click()
        
        endRepeat = driver.find_element_by_xpath('//*[@id="mec_date_repeat_end_at_date"]')
        endRepeat.send_keys(endRepeatDateCell[:10])
        endRepeat.send_keys(Keys.ENTER)
    
    'Function for the language of an event'
    def Language():
        
        
        languageArray = ['English','French','Hindi','Other','Punjabi','Urdu']
        languageArrayXpath = ['//*[@id="mec_label239"]','//*[@id="mec_label241"]','//*[@id="mec_label243"]','//*[@id="mec_label247"]','//*[@id="mec_label245"]','//*[@id="mec_label313"]']
        
        'xpath indexing within the loop cannot occur here since the index does not follow a pattern in realtion to the languageArray'
        i=0
        for i in range (len(languageArray)):
            if (languageCell) == languageArray[i]:
                lang = driver.find_element_by_xpath(languageArrayXpath[i])
                lang.click()
                break
            
    Language()
    
    'Scroll back to top of the page and publish the event. Create a new event and repeat the process'
    driver.execute_script("window.scrollTo(0, 250)")
    
    Submit = driver.find_element_by_xpath('//*[@id="publish"]')
    Submit.click()
    
    'Test for event existing'
    eventID = driver.find_element_by_xpath('//*[@id="post_ID"]')
    AutomationTest.test_eventExists(eventID)  
    
    newEvent = driver.find_element_by_xpath('//*[@id="wpbody-content"]/div[4]/a')
    newEvent.click()
    
    row = row + 1
    
   





        
    






