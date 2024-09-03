# OATS Inator

Automating the entry of OATS in Orion to save you time and preserve your sanity!

[[_TOC_]]

## Setup

Download Python 3.8.6 (or higher - I haven't tested below that version) from [here](https://www.python.org/ftp/python/3.8.6/python-3.8.6-amd64.exe).

Install Python following the on screen instructions. My example commands assume that you install Python under **C:\Python**.

cd into the directory where you downloaded the OATS Inator using **Power Shell**.

Next up we are creating the Python environment to run this application:

```bash
# Creating the venv for Python - assumes you have cd into the directory of the OATS Inator
C:\Python\python.exe -m venv .\OATSinator
# Installing the required packages
.\OATSinator\Scripts\pip.exe install selenium
.\OATSinator\Scripts\pip.exe install pandas
.\OATSinator\Scripts\pip.exe install openpyxl
.\OATSinator\Scripts\pip.exe install webdriver_manager
```

Now the Setup is complete!

## A note on dates

Please be aware that the OATSInator assumes that your browser language is set to English (US) - you can check this by entering the following in Chrome: chrome://settings/languages and ensuring that the display language is English (US):

![Chrome Language](./img/Chrome-Language.png)

This is needed because dates are handled differently across the globe. If you do not want change your browser language then please adjust how you enter dates in the **Enter-OATS.xlsx** - you can test the right format by entering different values into the dates field of an OATS entry manually and see what works. For German for example the format 01July23 works but 01JUL23 doesn't work.

## Entering OATS

Open up the Excel workbook **Enter-OATS.xlsx**

Let me first describe the four sheets first:

1. **Data** - The sheet where you enter new OATS
2. **Metadata** - Contains all the values that need to be selected from dropdowns in Orion (Category, Value Proposition, Duration) - you shouldn't need to modify this sheet - if you find an issue here, please raise the problem
3. **List** - Here you enter all the Opportunities and Accounts on which you want to book OATS
4. **Information** - Here you can find additional links to confirm the OATS you entered are really entered into the system

All dropdowns created in the Excel-Sheets are created using the Data Validation functionality of Excel.

### Data Sheet

Now lets describe the different fields in the Data spreadsheet - as a tip to open up dropdowns in Excel use the keyboard shortcut **alt + arrow down** and then navigate with the **arrow keys**, when you have found your desired package hit the **enter** key.

Here you enter the actual data that gets entered into Orion.

**Note:** Do not change the headers on these sheets!

**Note:** The included Spreadsheet contains an example OATS entry please remove it before running the application for the first time!

- **OATS-Type:** The two options are _Opportunity_ and _Account_ and this is in reference to where you book these OATS in Orion. This is very important as there are many differences between Opportunity and Account based OATS. _This field is required!_
- **ID:** Here you select from the List of provided id in the List spreadsheet. Please insure that the ID you provide matches the **OATS-Type**. If you have an idea on how to filter these IDs depending on the OATS-Type columns please reach out. _This field is required!_
- **Date:** This is the date on which you did the activity. Please follow the format DDMMMYY - Concrete if it is March the 16th in 2022 then you enter 16Mar22. _This field is required!_ - please refer to the note on dates above
- **Duration:** How long did you spend on this activity. _This field is required!_
- **Category:** Select the category that best describes your activity. _This field is required!_
- **Value Proposition:** Select the appropriate value proposition for your activity. This field is only required for Account level OATS and will be ignored for Opportunity level OATS
- **Customer Facing Virtual:** Select (_Yes_) this if the activity was virtually customer facing. If you do not provide a value _No_ is assumed. This field is optional. Please note if you select this field, then Customer Facing Onsite will be ignored.
- **Customer Facing Onsite:** Select (_Yes_) if this activity was onsite with the customer. If you do not provide a value _No_ is assumed. This field is optional.
- **Demonstration Give:** Select (_Yes_) if you gave a demonstration during the activity. If you do not provide a value _No_ is assumed. This field is optional.
- **Partner Present:** Select (_Yes_) if a partner was present during the activity. If you do not provide a value _No_ is assumed. This field is optional.
- **Interregional Collaboration:** This is a temporary solution for you to be able to write down your interregional collaboration but not have it entered into Orion yet. If possible I will provide a migration solution based on this column in the future.
- **Done:** The application will enter an _X_ once it has entered the OATS in Orion. Do not edit this field yourself.

Conditional Formatting rules that are applied to the sheet:

- **Date:** *=AND($C1="";OR($A1="Opportunity";$A1="Account"))*
- **Duration:** *=AND($D1="";OR($A1="Opportunity";$A1="Account"))*
- **Category:** *=AND($E1="";OR($A1="Opportunity";$A1="Account"))*
- **Value Proposition:** *=AND($F1="";$A1="Account")*

### List Sheet

Now lets describe the different fields in the List spreadsheet. Here you enter the information about the Opportunities and Accounts you work on.

**Note:** The included Spreadsheet contains an example Account entry please remove it before running the application for the first time!

- **OATS-Type:** The two options are *Opportunity* and *Account* and this is in reference to where you book these OATS in Orion. This is very important as there are many differences between Opportunity and Account based OATS. _This field is required!_
- **Name:** Here you can enter any name that helps you to identify the Accounts and Opportunities, it is not used by the application itself. You are not allowed enter _double pipes (||)_ in this field.
- **Orion-ID:** Please enter the actual Orion-ID of the Account or Opportunity here. This ID has to match with the ID in Orion. _This field is required!_
- **Combined-Data:** This field combines the **Name** and **Orion-ID** column, separating them with _double pipes (||)_. Do not change this field, an Excel-Formula is applied (_=IF(A2="";" ";CONCAT(B2; " || ("; C2; ")"))_).

## Running the OATS Inator

Before running please insure that you have followed the Setup instructions and have thoroughly read the chapter on Entering OATS.

While the OATS Inator is running please do not click or type on your PC, while it shouldn't matter in theory I haven't checked it in practice. So either run the command below and grab yourself a coffee or watch the OATS Inator do its magic. The output in Power Shell informs about what the application is doing and reports the URLs as well. If you receive any errors please reach out.

The application always runs through all Account level OATS first and then starts entering the Opportunity level OATS afterwards. After everything is entered the browser is closed. The application will fail to run if you have the Excel sheet still open, so please ensure that you have closed it.

```bash
# Assumes you have cd into the directory of the OATS Inator
.\OATSinator\Scripts\python.exe .\OATSInator.py
```

Happy OATSing!

Please regularly check the OATS Data Quality Report and fix the issues directly in Orion: https://cisviya.sas.com/SASVisualAnalytics/?appSwitcherMessage=%7B%22type%22%3A%22App.Open%22%2C%22data%22%3A%7B%22directive%22%3A%22SASVisualAnalytics%22%2C%22targetUri%22%3A%22%2Freports%2Freports%2F7191b1fb-e0e5-430b-8c8d-b83f84afd218%22%2C%22targetPublicType%22%3A%22report%22%7D%7D

## Reporting

Within this repo you will also find the OATSInator-Stats-VA-Report.json which is an exported report to give you first quick insights into your OATSing.

The report was created on Viya 2021.1.5, the report is saved in the Public folder and the data is assumed to be saved in your own casuser. Please import the data using the Data Explorer > Local files > Microsoft Excel (multiple worksheets) > Select only the Data sheet and import the data.

## Common Error Messages (and how to fix them)

Here I describe common error messages that can occur while using the OATS Inator.

### Message: session not created

Sometimes, especially if you leave your PC running for some time you might receive an error message that looks like this:

![Session not created](img/session-not-created.png)

The OATS Inator assumes that you are always on the current version of Chrome (this enables it to not require you to install an additional Chrome driver). But because of this assumption this error message can occur. To fix this enter the following as a URL in Chrome *chrome://settings/help* and update Chrome to the latest version (maybe you will have to relaunch Chrome). Now rerun the OATS Inator and this should be fixed.

