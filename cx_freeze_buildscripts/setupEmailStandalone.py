import os
from cx_Freeze import setup, Executable

os.environ["TCL_LIBRARY"] = r"C:\Users\I856620\AppData\Local\Programs\Python\Python36\tcl\tcl8.6"
os.environ["TK_LIBRARY"] = r"C:\Users\I856620\AppData\Local\Programs\Python\Python36\tcl\tk8.6"

#dependencies
build_exe_options = {"packages": 
					["os", "sys","csv","dotenv","win32com"],
					"excludes": []}

setup(
	name="emailUsers",
	version="0.1",
	description="Email user credentials.",
	executables=[Executable("emailStandalone.py")])