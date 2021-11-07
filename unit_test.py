# -*- coding: utf-8 -*-
import unittest
    
class AutomationTest(unittest.TestCase):
     
    'If the correct fields contain the password and username provided, test passed'
    def test_login(UserIN,PassIN,username,password): 
        userAssert = username.get_attribute("value") 
        
        try:
            assert UserIN == userAssert
            print('Username  correct')
        except AssertionError as error:
            print(error)
        
        passwordAssert = password.get_attribute("value")
        
        try:   
            assert PassIN == passwordAssert
            print('Password correct')
        except AssertionError as error:
            print(error)
    
    'If the description from the correct cell is inputted into the correct field, test passed'
    'Every event has a pre-written component to the description. This is added with manipulatedDescription'
    def test_description(descriptionCell,totalDescription):
        manipulatedDescription = ('<strong>Description:</strong>\n\n'+descriptionCell)
        
        try:
            assert totalDescription == manipulatedDescription
            print("Description added")
        except AssertionError as error:
            print(error)
    
    'if the start date and end date fields match, test passed'
    'if the date from the cell matches the start date in the correct field, test passed'
    'We can assert only to the start date since we have already asserted that start date and end date are equal'
    def test_startEndDate(sdate,edate,dateCell):
        sdateAssert = sdate.get_attribute("value")
        
        try:
            assert sdate == edate
            print("Event dates match")
        except AssertionError as error:
            print(error)
            
        try:
            assert dateCell[:10] == sdateAssert
            print("Event date correct")
        except AssertionError as error:
            print(error)
     
    'If the eventID exists, test passed'
    'If an event has been posted it must have an ID. We simply ensure that an ID is present after submit is clicked'
    'This test inherently also confirms the function of the submit button, since an ID cannot exist unless the event was submitted'
    def test_eventExists(eventID):
            eventIDAssert = eventID.get_attribute("value")
            
            try:
                assert int(eventIDAssert) > 0
                print(eventIDAssert)
                print("Event upload success")
            except AssertionError as error:
                print(error)
               

if __name__ == "__main__":
    unittest.main()
    