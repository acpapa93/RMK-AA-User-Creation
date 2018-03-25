import os
import random
import string
import csv
import time
import sys
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import win32com.client as win32

#set path if script is frozen or not
if getattr(sys, 'frozen', False):
    # frozen
    dir_path_tmp = os.path.dirname(sys.executable).split("\\")
else:
    # unfrozen
    dir_path_tmp = os.path.dirname(os.path.realpath(__file__)).split("\\")

dir_path_root = "\\".join(dir_path_tmp[0:3])
dir_path=dir_path_root+"\\Documents\\Analytics\\UserCreation"

#load in PS user credentials
dotenv_path = os.path.join(dir_path,".env")
load_dotenv(dotenv_path)
_PsUsr = os.environ.get("PSUSR")
_PsPwd = os.environ.get("PSPWD")
_PsUsrEmail = os.environ.get("PSUSREMAIL")

#LOAD IN DC URLS
_rackspaceLogin = os.environ.get("RACKSPACELOGIN")
_rackspaceMrAdmin = os.environ.get("RACKSPACEMRADMIN")
_rackspaceViewBuilder = os.environ.get("RACKSPACEVIEWBUILDER")
_DC12Login = os.environ.get("DC12LOGIN")
_DC12MrAdmin = os.environ.get("DC12MRADMIN")
_DC12ViewBuilder = os.environ.get("DC12VIEWBUILDER")
_DC17Login = os.environ.get("DC17LOGIN")
_DC17MrAdmin = os.environ.get("DC17MRADMIN")
_DC17ViewBuilder = os.environ.get("DC17VIEWBUILDER")
accepted_dataCenters = ["Rackspace", "rackspace", "DC12", "dc12", "Dc12", "DC17", "dc17", "Dc17"]
rackspace = accepted_dataCenters[0:2]
DC12 = accepted_dataCenters[2:5]
DC17 = accepted_dataCenters[5:8]
accepted_userTypes = ["Standard","standard","Restricted","restricted","Site Manager","site manager","Sitemanager","sitemanager","Internal","internal"]
standard = accepted_userTypes[0:2]
restricted = accepted_userTypes[2:4]
siteManager = accepted_userTypes[4:8]
internal = accepted_userTypes[8:10]

#Set Filenames
users = os.path.join(dir_path,"users.csv")
usersOutput = os.path.join(dir_path,"usersOutput.csv")
userFailures = os.path.join(dir_path,"userFailures.csv")

#chromedriver.exe must be in Analytics > usercreation folder
chromepath = os.path.join(dir_path,"chromedriver.exe")
#open browse
browser = webdriver.Chrome(executable_path=chromepath)
browser.addheaders = [("User-agent","Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.2.13) Gecko/20101206 Ubuntu/10.10 (maverick) Firefox/3.6.13")]

#creates password for each user profile.
def createUserPwd():
    with open(users,'r') as userList:
        reader = csv.reader(userList) 
        # skip the headers
        next(reader, None)
        all = []

        for row in reader:
            #pull out any leading spaces from the csv
            row = [x.strip(' ') for x in row]

            ##Validate user info provided
            #remove trailing commmas for when user info is added via Excel.
            if row[5] =="":
                row = row[:5]

            #ensure each row has required number of arguements
            if len(row)!=5:
                browser.close()
                sys.exit("User " + row[0] + """ has incorrect number of inputs. Please check arguements provided in users.csv and try again.""")

            #ensure all parameters are filled in
            for x in row:
                if x == "":
                    browser.close()
                    sys.exit("User " + row[0] + " has a blank input. Please update and try again.")

            #ensure DC is supported
            if row[2] not in accepted_dataCenters:
                browser.close()
                sys.exit("User " + row[0] + """ has an incorrectly entered or unknown datacenter. Available datacenter options are 'Rackspace' or 'rackspace','DC12' or 'dc12, or 'DC17' or 'dc17'. Please update and try again.""")

            #ensure profile is of the accepted user types
            if row[3] not in accepted_userTypes:
                browser.close()
                sys.exit("User " + row[0] + """ does not have the appropriate user type. Available options are 'Restricted', 'restricted','Standard', 'standard', "site manager", "Site Manager", "sitemanager", "Sitemanager", "Internal", or "Internal". Please update and try again.""")

            #create password
            pwd = ''.join(random.choices(string.ascii_letters + string.digits, k=14))
            row.append(pwd)
            all.append(row)

        with open(usersOutput, 'w', newline="") as userListOutput:
            writer = csv.writer(userListOutput)
            writer.writerows(all) 


    print("Passwords created. Moving to Managed Reoporting Admin.")
    readFileAndGroupDCs()


#group users on DC for managedReportingAdmin function
def readFileAndGroupDCs():
    with open(usersOutput,"r") as userListRead:
        userFileReader = csv.reader(userListRead)
        for row in userFileReader:
            if row[2] in DC17:
                managedReportingAdmin(_DC17MrAdmin, row)

            elif row[2] in DC12:
                managedReportingAdmin(_DC12MrAdmin, row)

            else:
                managedReportingAdmin(_rackspaceMrAdmin, row)

    print("Managed Reporting Admin complete. Let's log in to each profile.")
    login()

#function creates profile in Managed Reporting Admin
def managedReportingAdmin(managedReportingAdminUrl, row):
    try:
        browser.get(managedReportingAdminUrl)
        
        #enter PS username
        PSusernameWait = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.NAME, 'WFA_UserName'))
        )
        PSusername = browser.find_element_by_name('WFA_UserName')
        PSusername.send_keys(_PsUsr)
        
        #enter PS password
        password = browser.find_element_by_name('WFA_Password')
        password.send_keys(_PsPwd)
        #submit login form
        submit = browser.find_element_by_id("sm2")
        submit.click()

        #click users tab
        browser.switch_to.frame("tabFrame")
        usersTabWait = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.ID, 'showUsersTab'))
        )
        usersTab = browser.find_element_by_id('showUsersTab')
        usersTab.click()
        
        #switch out of the nav frame at the top
        browser.switch_to.default_content()
        #click new user button
        browser.switch_to.frame("objectListFrame")

        #check to see if user exists already. Delete user if it does exist already.
        checkForExistingUser(row)

        #continue on with user profile creation
        #click new user button
        newUserWait = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.ID, 'btnNewUser'))
        )
        newUser = browser.find_element_by_id('btnNewUser')
        newUser.click()
        
        #switch out of previous frame
        browser.switch_to.default_content()
        #switch into main box frame
        browser.switch_to.frame("resultSetFrame")

        #enter user info
        newUserWait = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.ID, 'generalInfoBlock_CONTENT'))
        )
        #user id
        enterUserId=browser.find_element_by_name("WFA_UserName")
        enterUserId.send_keys(row[0])
        #name
        enterName=browser.find_element_by_name("WFA_Description")
        enterName.send_keys(row[1])
        #password
        enterPwd=browser.find_element_by_name("WFA_Password")
        enterPwd.send_keys(row[5])
        #password verification
        enterPwd2=browser.find_element_by_name("WFA_ConfirmPassword")
        enterPwd2.send_keys(row[5])
        #email
        enterEmail=browser.find_element_by_name("WFA_Email")
        enterEmail.send_keys(row[0])

        #determine if user is standard or restricted
        if row[3] in standard:
            #select external user
            selectExternal = browser.find_element_by_link_text("Clara External")
        elif row[3] in restricted:
            #select external no detail
            selectExternal = browser.find_element_by_link_text("Clara External No Detail")
        elif row[3] in siteManager:
            #select siteManager
            selectExternal = browser.find_element_by_link_text("Clara Site Manager")
        elif row[3] in internal:
            selectExternal = browser.find_element_by_link_text("Clara Internal")

        else:
            print("Error: user type not specified or poorly specified")
        #double click the selected view
        actionChains = ActionChains(browser)
        actionChains.double_click(selectExternal).perform()

        #save user
        saveUser=browser.find_element_by_id("saveLink")
        saveUser.click()

    except Exception as error:
        print("error in user creation from managedReportingAdmin.")
        print(error)


#this function checks for any existing user and will delete the profile for recreation.
def checkForExistingUser(row):
    waitForUserSearch = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.NAME, 'WFA_PatternType'))
        )
    searchById = browser.find_element_by_css_selector("#userSearchForm > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td:nth-child(2) > select > option:nth-child(2)")
    searchById.click()
    waitForUserSearchBox = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#userSearchForm > table > tbody > tr:nth-child(3) > td:nth-child(1) > input[type='text']"))
        )
    userSearchForm = browser.find_element_by_css_selector("#userSearchForm > table > tbody > tr:nth-child(3) > td:nth-child(1) > input[type='text']")
    userSearchForm.send_keys(row[0])
    submitButtonMR = browser.find_element_by_css_selector("#userSearchForm > table > tbody > tr:nth-child(3) > td:nth-child(2) > a")
    submitButtonMR.click()

    try:
        #need to wait momentarily so it can update the numbers based on filter results.
        time.sleep(1.5)
        messageRetreived = browser.find_element_by_css_selector("#retrievedMessage1")
        messageText = messageRetreived.text

        #if user already exists, delete it and make new profile.
        if messageText == "Results 1-1 of 1":
            #click user
            userFound = browser.find_element_by_css_selector("#userListFieldTable")
            userFound.click()
            #click delete button
            delete = browser.find_element_by_css_selector("#btnDeleteUser")
            delete.click()
            
            #accept the alert that pops up
            browser.switch_to.alert.accept()
            #allow page to refresh briefly before moving on to create profile.
            time.sleep(1)

        else:
            return

    except NoSuchElementerror:
        print("Unable to determine if user already exists.")


#function logs in to profile
def login():
    with open(usersOutput,"r") as userListRead:
        userFileReader = csv.reader(userListRead)
        for row in userFileReader:
            if row[2] in rackspace:
                loginUrl = _rackspaceLogin

            elif row[2] in DC12:
                loginUrl = _DC12Login

            elif row[2] in DC17:
                loginUrl = _DC17Login

            else:
                sys.exit("unable to login as datacenter is not correct.")

            browser.get(loginUrl)
            try:
                element = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located((By.ID, 'enterid'))
                )
                username = browser.find_element_by_id('enterid')
                username.send_keys(row[0])
                password = browser.find_element_by_name('WORP_PASS')
                password.send_keys(row[5])
                form = browser.find_element_by_id('login')
                form.submit()

            except Exception as error:
                print(error)
   
    print("Logging in completed. Starting viewbuilder.")
    #call viewbuilder function
    prepForViewBuilder()


#prep users for viewbuilder
def prepForViewBuilder(): 
    #initiate Arrays for different users
    DC17Arr = []
    DC17RestrictedArr = []
    DC17SMArr = []
    DC17InternalArr = []
    DC12Arr = []
    DC12RestrictedArr = []
    DC12SMArr = []
    DC12InternalArr = []
    rackspaceArr = []
    rackspaceRestrictedArr = []
    rackspaceSMArr = []
    rackspaceInternalArr = []

    with open(usersOutput,"r") as userListRead:
        userFileReader = csv.reader(userListRead)
        for row in userFileReader:

            #Segment users based on DC and profile type
            if row[2] in DC17:                
                
                if row[3] in restricted:
                    DC17RestrictedArr.append(row[0])
                
                elif row[3] in siteManager:
                    DC17SMArr.append(row(0))
                
                elif row[3] in internal:
                    DC17InternalArr.append(row(0))
                
                else:
                    DC17Arr.append(row[0])

            elif row[2] in DC12:                
                
                if row[3] in restricted:
                    DC12RestrictedArr.append(row[0])
                
                elif row[3] in siteManager:
                    DC12SMArr.append(row[0])
                
                elif row[3] in internal:
                    DC12InternalArr.append(row[0])
                
                else:
                    DC12Arr.append(row[0])
            
            else:
                
                if row[3] in restricted:
                    rackspaceRestrictedArr.append(row[0])
                
                elif row[3] in siteManager:
                    rackspaceSMArr.append(row[0])

                elif row[3] in internal:
                    rackspaceInternalArr.append(row[0])

                else:
                    rackspaceArr.append(row[0])
    
    #call viewbuilder for each DC and usertype
    if rackspaceArr:
        viewBuilder(rackspaceArr, _rackspaceViewBuilder, "Standard", "Rackspace")
    
    if rackspaceRestrictedArr:
        viewBuilder(rackspaceRestrictedArr, _rackspaceViewBuilder, "Restricted", "Rackspace")

    if rackspaceSMArr:
        viewBuilder(rackspaceSMArr, _rackspaceViewBuilder, "siteManager", "Rackspace")

    if rackspaceInternalArr:
        viewBuilder(rackspaceInternalArr, _rackspaceViewBuilder, "internal", "Rackspace")

    if DC12Arr:
        viewBuilder(DC12Arr, _DC12ViewBuilder, "Standard", "DC12")

    if DC12RestrictedArr:
        viewBuilder(DC12RestrictedArr, _DC12ViewBuilder, "Restricted", "DC12")

    if DC12SMArr:
        viewBuilder(DC12SMArr, _DC12ViewBuilder, "siteManager", "DC12")

    if DC12InternalArr:
        viewBuilder(DC12InternalArr, _DC12ViewBuilder, "internal", "DC12")

    if DC17Arr:
        viewBuilder(DC17Arr, _DC17ViewBuilder, "Standard", "DC17")

    if DC17RestrictedArr:
        viewBuilder(DC17RestrictedArr, _DC17ViewBuilder, "Restricted", "DC17")

    if DC17SMArr:
        viewBuilder(DC17SMArr, _DC17ViewBuilder, "siteManager", "DC17")

    if DC17InternalArr:
        viewBuilder(DC17InternalArr, _DC17ViewBuilder, "internal", "DC17")

    print("Views set. Login again to test final view.")
    cleanUpFiles()


##function sets user view within viewbuilder
def viewBuilder(DcArray, vbURL, userType, DCFlag):
    try:
        #set Url
        browser.get(vbURL)
        
        #enter username
        waitForPSUser = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.NAME, 'WORP_USER'))
        )
        username = browser.find_element_by_name('WORP_USER')
        username.send_keys(_PsUsr)
        
        #enter password
        password = browser.find_element_by_name('WORP_PASS')
        password.send_keys(_PsPwd)
        
        #submit login form
        vbForm = browser.find_element_by_tag_name("form")
        vbForm.submit()

        #get to userList
        #switch into gentools frame
        browser.switch_to.frame("gentools")
        vbManageUsers=browser.find_element_by_link_text("Manage Users")
        vbManageUsers.click()
        time.sleep(0.5)
        #page only half-renders when you click the vbManageUsers link once. Click it twice and it renders fully.
        vbManageUsers.click()
        print("Letting user list load. Please wait.")
        time.sleep(5)
        browser.switch_to.default_content()
       
        element = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.NAME, "showFrame"))
        )
        browser.switch_to.frame("showFrame")
        browser.switch_to.frame("viewlist")

        #find and select the users
        for user in DcArray:
            #encode @ symbol to %40
            if "@" in user:
                user = user.replace("@","%40")
            #lowercase the username
            user = user.lower()
            
            try:
                element = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located((By.ID, user))
                )
                newRecruiter = browser.find_element_by_id(user)
                newRecruiter.click()
            except Exception as error:
                print("unable to find user: "+ user)

        #click reassign button
        browser.switch_to.default_content()
        browser.switch_to.frame("showFrame")
        browser.switch_to.frame("viewcontrol")

        reassignBtn=browser.find_element_by_css_selector("body > table > tbody > tr > td:nth-child(2) > a")
        reassignBtn.click()
        browser.switch_to.default_content()
        #switch back to the right frame / window
        browser.switch_to.frame("showFrame")
        browser.switch_to.frame("addview")

        #wait for viewlist options to populate
        element = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/form/table[1]/tbody/tr/td/select/option[11]"))
        )

         #alter the selector if restricted otherwise clara external
        if DCFlag == "Rackspace":
            if userType == "Restricted":
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[26]")
            
            elif userType == "siteManager":
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[19]")

            elif userType == "internal":
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[4]")

            else:
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[11]")
        
        elif DCFlag == "DC12":
            if userType == "Restricted":
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[14]")

            elif userType == "siteManager":
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[13]")

            elif userType == "internal":
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[6]")

            else:
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[3]")
        
        elif DCFlag == "DC17":
            if userType == "Restricted":
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[16]")

            elif userType == "siteManager":
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[15]")

            elif userType == "internal":
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[13]")

            else:
                viewList = browser.find_element_by_xpath("/html/body/form/table[1]/tbody/tr/td/select/option[5]")
        else:
            print("Unable to set DCFlag for viewbuilder view selection.")

        viewList.click()

        #click the "replace with user content toggle"
        replaceWithContact = browser.find_element_by_css_selector("body > form > table:nth-child(7)")
        replaceWithContact.click()
        
        #click the submit button
        submitBtnClick = browser.find_element_by_css_selector("body > table > tbody > tr > td:nth-child(2) > a")
        submitBtnClick.click()

        #accept the alert that pops up
        browser.switch_to.alert.accept()

    except Exception as error:
        print(error)


#move all user credentials back to the users.csv file and then delete the usersOutput.csv file.
def cleanUpFiles():
    topRow = ["Email", "Name", "Datacenter", "Standard or Restricted", "Password"]
    allUsers = []
    with open(usersOutput,'r') as userListOutput:
        with open(users, 'w', newline="") as userList:
            #write in header row again
            writer = csv.writer(userList)
            writer.writerow(topRow)

            #read out all user profies and write back to the original users.csv file.
            reader = csv.reader(userListOutput)
            for row in reader:
                allUsers.append(row)
            writer.writerows(allUsers)   

    #delete usersOutput.csv file.
    os.remove(usersOutput)
    testLoginAndEmail()


#tests if login was successful
def testLogin(user, loginUrl):
    browser.get(loginUrl)
    try:
        element = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.ID, 'enterid'))
        )
        username = browser.find_element_by_id('enterid')
        username.send_keys(user[0])
        password = browser.find_element_by_name('WORP_PASS')
        password.send_keys(user[5])
        form = browser.find_element_by_id('login')
        form.submit()

        #switch to the right frame
        browser.switch_to.frame("contentFrame")
        browser.switch_to.frame(browser.find_element_by_tag_name("iframe"))
        
        #swap frame if standard or restricted
        if user[3] in standard or user[3] in restricted:
            browser.switch_to.frame("ae_launch")

        #swap frame if site manager
        elif user[3] in siteManager:
            browser.switch_to.frame("aa_launch")

        #swap frame if internal
        elif user[3] in internal:
            browser.switch_to.frame("ah_launch")

        element = WebDriverWait(browser, 15).until(
            EC.presence_of_element_located((By.ID, 'image3'))
        )

        browser.switch_to.frame(browser.find_element_by_tag_name("iframe"))

        waitForLoading = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.ID, 'text1'))
        )

        #wait for error message saying view could not be set and write to failure file.
        try:
            waitForLoad = WebDriverWait(browser, 15).until(
            EC.presence_of_element_located((By.ID, 'text7'))
            )
            print(user[0] + " has no Command Center assignment. Profile is not complete.")
            with open(userFailures, 'w', newline="") as failures:
                writer = csv.writer(failures)
                writer.writerow(user)

        #if company title is found and no error, go ahead and send email.
        except:
            print("Attempting to verify user creation. Please wait. This may take a moment.")
            verifyCompanyTitle = browser.find_element_by_xpath("/html/body/div/div[1]")
            print("User profile complete. Emailing credentials.")
            
            #email credentials if appropriate.
            if user[4] == "Yes" or user[4] == "yes":
                emailCredentials(user[0], user[5], loginUrl)

    except Exception as e:
        print(e)
        print("Unable to verify profile setup successful.")
        with open(userFailures, 'w', newline="") as failures:
                writer = csv.writer(failures)
                writer.writerow(user)


#email credentials
def emailCredentials(username, password, loginUrl):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    username = username
    password = password
    loginURL = loginUrl
    mail.To = username
    mail.Bcc = _PsUsrEmail
    mail.Subject = 'Advanced Analytics Profile Credentials'
    body = """Greetings,\n 
I created your Recruiting Marketing Advanced Analytics profile. Please see your login credentials below. 

Username: """+username+"""
Password: """+password+"""
URL: """+ loginURL+"""

Please reset your password via the link on the login screen.\n
Thanks,
The Advanced Analytics Team"""
    mail.Body = body

    mail.Send()

#test the profiles
def testLoginAndEmail():
    with open(users, 'r') as userList:
        reader = csv.reader(userList)
        next(userList, None)
        for user in reader:

            if user[2] in rackspace:
                loginUrl = _rackspaceLogin
                testLogin(user, loginUrl)
            
            elif user[2] in DC12:
                loginUrl = _DC12Login
                testLogin(user, loginUrl)
            
            elif user[2] in DC17:
                loginUrl = _DC17Login
                testLogin(user, loginUrl)
            
            else:
                print("Unable to test login as datacenter doesn't match.")

    print("User creation complete. Please review successes and failures.")
    browser.close()

#start the thing!
createUserPwd()
