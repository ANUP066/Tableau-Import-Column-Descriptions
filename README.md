# Tableau-Import-Column-Descriptions
A Python Application that helps in importing the column comments of a table in a Oracle database to the Tableau workbook. This application thereby automates this task of getting column comments in the workbook which if done manually for a large datasource can be time consuming.

## How to run :
* Install all the required libraries mentioned in the import statements of the Python code.
* Connect your Tableau workbook to the required datasource.
* Run the python code using python importDescriptions.py
* A GUI will open up asking for your details to connect to the Oracle database.
* At the end of it, browse and select your workbook to which you want to import the column comments of the table.
* Click on submit and the application runs and notifies when finished through a pop-up.
* Reopen the Tableau workbook to the see the changes in effect.
* Works for Single Table, Multiple Table joins, Custom SQL Tableau datasources
* Tested for Tableau Desktop versions - 2018 to 2020.4.4
* Also uploaded a Python file that helps in converting it to an executable on Windows OS platforms.
