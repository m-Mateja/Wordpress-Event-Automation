# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
df = pd.read_excel('')

'Opening Wordpress--------------------------------------------------------------'
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(executable_path='',chrome_options=chrome_options)
driver.get('')

loginbutton = driver.find_element_by_xpath('//*[@id="jetpack-sso-wrap"]/a[1]')
loginbutton.click()

UserIN = input('Enter your Username: ')
username = driver.find_element_by_xpath('//*[@id="user_login"]')
username.send_keys(UserIN)
time.sleep(1)

PassIN = input('Enter your Password: ')
password = driver.find_element_by_xpath('//*[@id="user_pass"]')
password.send_keys(PassIN)

login = driver.find_element_by_xpath('//*[@id="wp-submit"]')
login.click()

'Start Cell integration----------------------------------------------------------'

row = 0
q = 0

while q == 0:

    cell1 = df.iloc[row,0]
    if cell1 == 'END':
        break

    cell1a = df.iloc[row,1]
    cell2 = df.iloc[row,2]
    cell3 = df.iloc[row,3]
    cell4 = df.iloc[row,4]
    cell5 = df.iloc[row,5]
    cell6 = df.iloc[row,6]
    cell7 = df.iloc[row,7]
    cell8 = df.iloc[row,8]
    cell9 = df.iloc[row,9]
    cell10 = df.iloc[row,10]
    cell11 = df.iloc[row,11]
    cell12 = df.iloc[row,12]
    
    'Event Creation'
    time.sleep(0.5)
    title = driver.find_element_by_xpath('//*[@id="title"]')
    title.send_keys(cell1)

    textbutton = driver.find_element_by_xpath('//*[@id="content-html"]')
    textbutton.click()
    
    description = driver.find_element_by_xpath('//*[@id="content"]')
    description.send_keys('<strong>Description:</strong>')
    description.send_keys(Keys.ENTER)
    description.send_keys(Keys.ENTER)
    description.send_keys(cell1a)
    description.send_keys(Keys.ENTER)
    description.send_keys(Keys.ENTER)
    description.send_keys('<strong>Registration:</strong>')
    description.send_keys(Keys.ENTER)
    description.send_keys(Keys.ENTER)
    description.send_keys('To access this program click <a href='+cell12+'>here.</a>')
    
    sdate = driver.find_element_by_xpath('//*[@id="mec_start_date"]')
    sdate.send_keys(cell4[:10])
    
    edate = driver.find_element_by_xpath('//*[@id="mec_end_date"]')
    edate.send_keys(cell4[:10])
    
    def calendarDropdown():
        
        i=0
        time_arr = ['01','02','03','04','05','06','07','08','09','10','11']
        time_arr_2 = ['13','14','15','16','17','18','19','20','21','22','23']
        time_arr_3 = ['00','05','10','15','20','25','30','35','40','45','50','55']
       
        for i in range(len(time_arr)):
            if (cell2[0:2]) == time_arr[i]:
                starttimeH = driver.find_element_by_xpath('//*[@id="mec_start_hour"]/option['+str(i+1)+']')
                starttimeH.click()
                starttimeK = driver.find_element_by_xpath('//*[@id="mec_start_ampm"]/option[1]')
                starttimeK.click()
                break
            
            elif (cell2[0:2]) == time_arr_2[i]:
                starttimeH = driver.find_element_by_xpath('//*[@id="mec_start_hour"]/option['+str(i+1)+']')
                starttimeH.click()
                starttimeK = driver.find_element_by_xpath('//*[@id="mec_start_ampm"]/option[2]')
                starttimeK.click()
                break
            
        if (cell2[0:2]) == '12':
            starttimeH = driver.find_element_by_xpath('//*[@id="mec_start_hour"]/option[12]')
            starttimeH.click()
            starttimeK = driver.find_element_by_xpath('//*[@id="mec_start_ampm"]/option[2]')
            starttimeK.click()   
            
        for i in range(len(time_arr_3)):
            if (cell2[3:5] == time_arr_3[i]):
                 starttimeH = driver.find_element_by_xpath('//*[@id="mec_start_minutes"]/option['+str(i+1)+']')
                 starttimeH.click()
                 break
                
            
        '-------------------------------------------------------------------------------------------------'
        
        for i in range(len(time_arr)):
            if (cell3[0:2]) == time_arr[i]:
                starttimeH = driver.find_element_by_xpath('//*[@id="mec_end_hour"]/option['+str(i+1)+']')
                starttimeH.click()
                starttimeK = driver.find_element_by_xpath('//*[@id="mec_end_ampm"]/option[1]')
                starttimeK.click()
                break
            
            elif (cell3[0:2]) == time_arr_2[i]:
                starttimeH = driver.find_element_by_xpath('//*[@id="mec_end_hour"]/option['+str(i+1)+']')
                starttimeH.click()
                starttimeK = driver.find_element_by_xpath('//*[@id="mec_end_ampm"]/option[2]')
                starttimeK.click()
                break
            
        if (cell3[0:2]) == '12':
            starttimeH = driver.find_element_by_xpath('//*[@id="mec_end_hour"]/option[12]')
            starttimeH.click()
            starttimeK = driver.find_element_by_xpath('//*[@id="mec_end_ampm"]/option[2]')
            starttimeK.click()   
            
        for i in range(len(time_arr_3)):
            if (cell3[3:5] == time_arr_3[i]):
                 starttimeH = driver.find_element_by_xpath('//*[@id="mec_end_minutes"]/option['+str(i+1)+']')
                 starttimeH.click()
                 break
           
            
    calendarDropdown()
    
    tabswitchLocation = driver.find_element_by_xpath('//*[@id="mec_metabox_details"]/div[2]/div/div[1]/a[5]')
    tabswitchLocation.click()
    
    locationdrop = driver.find_element_by_xpath('//*[@id="select2-mec_location_id-container"]')
    locationdrop.click()
    
    locationset = driver.find_element_by_xpath('/html/body/span/span/span[1]/input')
    locationset.send_keys(cell7)
    locationset.send_keys(Keys.ENTER)
    
    tabswitchOrganizer = driver.find_element_by_xpath('//*[@id="mec_metabox_details"]/div[2]/div/div[1]/a[7]')
    tabswitchOrganizer.click()
    
    organizerdrop = driver.find_element_by_xpath('//*[@id="select2-mec_organizer_id-container"]')
    organizerdrop.click()
    
    organizerinput = driver.find_element_by_xpath('/html/body/span/span/span[1]/input')
    organizerinput.send_keys(cell8)
    organizerinput.send_keys(Keys.ENTER)
    
    def ProgramCategories():
        
        i=0
        type_arr = ['Online','In Centre','Outdoor']
        type_arr_2 = ['//*[@id="in-mec_category-166"]','//*[@id="in-mec_category-168"]','//*[@id="in-mec_category-167"]']
        type_ARR = ['EarlyON Outdoor','Exploring Nature','Exploring the Community','Family Time In Centre','Family Time Online','Healthy Eating In Centre','Healthy Eating Online','Indigenous-Led In Centre','Indigenous-Led Online','Mother Goose In Centre','Mother Goose Online','Music and Movement In Centre','Music and Movement Online','Parent Group & Resources In Centre','Parent Group & Resources Online','Story Time In Centre','Story Time Online']
        type_ARR_2 = ['//*[@id="in-mec_category-250"]','//*[@id="in-mec_category-251"]','//*[@id="in-mec_category-249"]','//*[@id="in-mec_category-161"]','//*[@id="in-mec_category-316"]','//*[@id="in-mec_category-162"]','//*[@id="in-mec_category-317"]','//*[@id="in-mec_category-163"]','//*[@id="in-mec_category-318"]','//*[@id="in-mec_category-159"]','//*[@id="in-mec_category-319"]','//*[@id="in-mec_category-160"]','//*[@id="in-mec_category-320"]','//*[@id="in-mec_category-164"]','//*[@id="in-mec_category-321"]','//*[@id="in-mec_category-158"]','//*[@id="in-mec_category-322"]']
        type_ARR_3 = ['#a8507c','#a8507c','#a8507c','#a8507c','#a8507c','#906fce','#906fce','#11637c','#11637c','#8ce580','#8ce580','#44bb99','#44bb99','#b391c9','#b391c9','#77aadd','#77aadd']
        
        for i in range(len(type_arr)):
            if (cell9) == type_arr[i]:
                category1 = driver.find_element_by_xpath(type_arr_2[i])
                category1.click()
                break
            
        for i in range(len(type_ARR)):
            if (cell10) == type_ARR[i]:
                category2 = driver.find_element_by_xpath(type_ARR_2[i])
                category2.click()
                colour = driver.find_element_by_xpath('//*[@id="mec_metabox_color"]/div[2]/div/div[1]/div/button')
                colour.click()
                colourSearch = driver.find_element_by_xpath('//*[@id="mec_event_color"]')
                colourSearch.clear()
                colourSearch.send_keys(type_ARR_3[i])
                break
            
    ProgramCategories()
    
    if (cell5) == 'None':
        pass
        
    else:
        
        tabswitch1 = driver.find_element_by_xpath('//*[@id="mec_metabox_details"]/div[2]/div/div[1]/a[2]')
        tabswitch1.click()
        
        eventrepeat = driver.find_element_by_xpath('//*[@id="mec_repeat"]')
        eventrepeat.click()
        
        certainweekdays = driver.find_element_by_xpath('//*[@id="mec_repeat_type"]/option[4]')
        certainweekdays.click()
        
        def weekdays():
            
            i=0
            day_arr = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
            day_arr_2 = ['//*[@id="mec_repeat_certain_weekdays_container"]/label[2]','//*[@id="mec_repeat_certain_weekdays_container"]/label[3]','//*[@id="mec_repeat_certain_weekdays_container"]/label[4]','//*[@id="mec_repeat_certain_weekdays_container"]/label[5]','//*[@id="mec_repeat_certain_weekdays_container"]/label[6]','//*[@id="mec_repeat_certain_weekdays_container"]/label[7]','//*[@id="mec_repeat_certain_weekdays_container"]/label[8]']
            
            for i in range(len(day_arr)):
                if (cell5) == day_arr[i]:
                    weekday = driver.find_element_by_xpath(day_arr_2[i])
                    weekday.click()
                    break
        
        weekdays()
        
        onButton = driver.find_element_by_xpath('//*[@id="mec_repeat_ends_date"]')
        onButton.click()
        
        endrepeat = driver.find_element_by_xpath('//*[@id="mec_date_repeat_end_at_date"]')
        endrepeat.send_keys(cell6[:10])
        endrepeat.send_keys(Keys.ENTER)
    
    def Language():
        
        i=0
        lang_arr = ['English','French','Hindi','Other','Punjabi','Urdu']
        lang_arr_2 = ['//*[@id="mec_label239"]','//*[@id="mec_label241"]','//*[@id="mec_label243"]','//*[@id="mec_label247"]','//*[@id="mec_label245"]','//*[@id="mec_label313"]']
        
        for i in range (len(lang_arr)):
            if (cell11) == lang_arr[i]:
                lang = driver.find_element_by_xpath(lang_arr_2[i])
                lang.click()
                break
            
    Language()
    
    driver.execute_script("window.scrollTo(0, 250)")
    
    Submit = driver.find_element_by_xpath('//*[@id="publish"]')
    Submit.click()
    
    newEvent = driver.find_element_by_xpath('//*[@id="wpbody-content"]/div[4]/a')
    newEvent.click()
    
    row = row + 1
    
   





        
    







