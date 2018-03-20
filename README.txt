ABOUT
===========================================================
userCreate.py:
Creates user profiles, tests the success of the profile creation, and emails out the successfully created profiles.

emailStandalone.py:
Batch emails profiles from the users.csv file.

Script Requirements
============================================================
1. Must have chromedriver.exe within Analytics > usercreation folder. 
2. Must have .env file with appropriate credentials and urls within Analytics > usercreation folder.
3. Users.csv file must be completed in with the appropriate information.

Directions
============================================================
1A: Double click either exe to run, or:
1B: open command line, enter: "cd documents/analytics/usercreation/build_usercreate_amd64_3.6". Then enter "usercreate.exe". 


Known Behaviors
============================================================
1. Rackspace viewbuilder often doesn't render the entire userlist. Connectivty seems to play a role in this issue. 
If you experience this issue, please retry or just continue on from viewbuilder step on your own.

2. The login testing process is pretty slow. You will see the page load correctly long before the script realizes the profile was created successfully. 
Please wait patiently as it attempts to register if it was successfully created or not. Login will typically fail due to lack of Command Center profile much quicker than the success will be registered. 

3. Chromedriver will often output errors within the console such as:
"[######:#####:####/####.###:ERROR:configuration_policy_handler_list.cc(92)] Unknown policy:" and other errors relating to failures to find by id. 
Please ignore these errors they are coming from chromedriver incorrectly.


FAQ
============================================================
Q: How do I stop a script once I start it?
A: Close out of the command line to stop the script.

Folder Structure
============================================================
|.env (required): enviornment variables. 
|.gitignore (required): excludes some files for upload to git.
|chromedriver.exe (required): automated browser
|README.txt
|emailStandalone.py: emails user profiles only. Must have python 3.6 to run.
|userCreate.py: creates and emails user profiles. Must have python 3.6 to run.
|users.csv (required): initial set of users to create / email.
|cx_freeze_buildScripts
	|setupUserCreate.py: build for .exe file.
	|setupEmailStandalone.py: build for .exe file.
|build_userCreate_amd64_3.6 (required): .exe version of user create script.
	|usercreate.exe (required): this is the actual .exe program to create the users.
	|Other files in this directory are all required.
|build_email_amd64_3.6 (required): .exe version of email standalone script.
	|emailUsers.exe (required): this is the actual .exe program to email credentials.
	|Other files in this directory are all required.