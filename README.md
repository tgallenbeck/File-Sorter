# Usage
This script is only used for @DillonGrech analysis spreadsheet

You will be asked 

# Requirements
You must have pipenv installed or some other equivalent virtual environment application.

# How to use
Once you have cloned the repo locally, navigate to the parent directory and run the following:
```
touch .env
```
Within the .env file, add the following:
```
DIRECTORY="<your_trading_directory>"
```
*Note: This path needs to be where you will export your back tests from TradingView. For example:*
    
    - Windows: DIRECTORY="C:\\Users\\user_name\\Documents\\Trading"
    - Linux/Mac: DIRECTORY="home/user_name/Documents/Trading"

Once this has been created, run the following to install the dependencies:
```
pipenv install
```

Once the dependencies have been installed, export your back tests to the directory that you created in your .env file, and make a copy of Dillons spreadsheet. 

*Note: I typically make a copy of the spreadsheet to reflect the back test I'm doing and will move it to another directory. For example:*
```
- Windows: "C:\Users\user_name\Documents\Trading\Indicators\C1\Didi_Index_BackTest.xlsm"
- Linux/Mac: "home/user_name/Documents/Trading/Indicators/C1/Didi_Index_Backtest.xlsm"
# Didi_Index_Backtest.xlsm is Dillons spreadsheet, but just copied (to preserve the integrity of the original) and renamed (to reflect your current back test)
```
Now just run the script and enjoy!
```
pipenv run python ./src/file_sorter.py
# When you run the script, you will be asked for the path to the spreadsheet you want to use
# This will be the path of the copied spreadsheet listed above in the example
```
