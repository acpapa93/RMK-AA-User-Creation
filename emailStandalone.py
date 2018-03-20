import os
import sys
import csv
from dotenv import load_dotenv
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
users = os.path.join(dir_path,"users.csv")

#load in .env variables
dotenv_path = os.path.join(dir_path,".env")
load_dotenv(dotenv_path)
_PsUsrEmail = os.environ.get("PSUSREMAIL")

#LOAD IN DC Info
accepted_dataCenters = ["Rackspace", "rackspace", "DC12", "dc12", "Dc12", "DC17", "dc17", "Dc17"]
rackspace = accepted_dataCenters[0:2]
DC12 = accepted_dataCenters[2:5]
DC17 = accepted_dataCenters[5:8]
_rackspaceLogin = os.environ.get("RACKSPACELOGIN")
_DC12Login = os.environ.get("DC12LOGIN")
_DC17Login = os.environ.get("DC17LOGIN")

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

with open(users, 'r') as userList:
        reader = csv.reader(userList)
        next(userList, None)
        for user in reader:
            
            if user[2] in rackspace:
                loginUrl = _rackspaceLogin
                emailCredentials(user[0],user[5],loginUrl)
            
            elif user[2] in DC12:
                loginUrl = _DC12Login
                emailCredentials(user[0],user[5],loginUrl)

            
            elif user[2] in DC17:
                loginUrl = _DC17Login
                emailCredentials(user[0],user[5],loginUrl)
            
            else:
                print("Unable to test login as datacenter doesn't match.")