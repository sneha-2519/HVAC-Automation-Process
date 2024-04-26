# HVAC-Automated Data Scraping, Mapping, and Comparison
Objective:
The objective of this process is to automate the extraction of data from the ENERGY STAR Certified Product Data Sets and APIs webpage, map the raw data according to customer requirements, create formatted UP files, convert them into JSON files, and compare them with the previous month's UP files to identify any new columns.

Steps:
Step 1: Data Scraping: Utilize Python and relevant libraries (e.g., BeautifulSoup, requests) to scrape data from the ENERGY STAR webpage. Identify and scrape data under the "Heating & Cooling" header and its 12 subcategories. Handle cases where direct scraping is not possible by using HTML code. Store the scraped data in Excel files (Raw files).

Step 2: Data Mapping: Based on customer requirements, map the raw Excel files to create formatted UP files. Perform necessary data manipulation, cleaning, and formatting during the mapping process.

Step 3: JSON File Creation Convert the formatted UP files into JSON files using Python. Ensure the JSON files adhere to the specified structure and contain all necessary data fields.

Step 4: Comparison with Previous Month: Retrieve the UP files from the previous month. Automatically compare the current formatted UP files with the previous month's UP files. Identify any new columns added in the current files compared to the previous month.

Step 5: Automated Execution: Develop a Python script to automate the entire process. Use Anaconda Prompt to execute the script, providing all required file paths as inputs. Implement error handling and logging mechanisms to ensure smooth execution and troubleshooting.

Tools and Resources: Python programming language BeautifulSoup, requests libraries for web scraping Pandas library for data manipulation and Excel handling JSON conversion tools/libraries Anaconda environment for script execution and package management

Continuous Improvement: Encourage feedback from stakeholders and team members to identify areas for process improvement. Implement changes based on feedback and lessons learned to enhance efficiency and effectiveness.

Documentation: Maintain comprehensive documentation covering all aspects of the automated process, including setup instructions, code documentation, and user guides. Document any updates, changes, or issues encountered during the process for future reference.

# To run the Script:
Open all the python Script on Pythhon IDE (Spyder), open the Anaconda Prompt to run the code

on the prompt mention five arguments like

python file name with extension
output directory
Location of brands mapping excel file
Location of Map_api excel file
Location of map file
Location last month files folder
for exapmle : Keeping all the python files in "D:\HVAC\COMBINING TEST\SCRIPT" and restof files and folder are in "D:\HVAC\Automate HVAC"

so on the commad prompt: python Automate__API.py "D:\HVAC\Automate HVAC" "D:\HVAC\Automate HVAC\brands-mapping.xlsx" "D:\HVAC\Automate HVAC\Map-Api.xlsx" "D:\HVAC\Automate HVAC\map.xlsx" "D:\HVAC\Automate HVAC\LAST MONTH\EXCEL"
