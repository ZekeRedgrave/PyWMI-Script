# PyWMI-Script
A script for Microsoft WMI 

# Basic
- This script will able to get all Hardware and Software Information Automatically without Command Prompt or Powershell
- Easy to Use and no need to type a command on Command Prompt or Powershell
- This is made from scratch for Python Purposes

# Tutorial
```py
import WMIC

x = WMIC.WMIC()
# Call All System Hardware Information and return dict() value
print(x.getBaseboard())

# You can get the Serial Number in Baseboard
print(x.getBaseboard(key="serialnumber"))
print(x.getBaseboard()["serialnumber"])
```

