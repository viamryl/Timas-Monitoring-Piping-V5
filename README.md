# PIPING MONITORING BETA VERSION
This program is developed to monitor PIPING projects at TIMAS SUPLINDO. It will synchronize the databases of the Engineering, QAQC, and PPC departments, while providing visualizations through a dashboard.

## About This Program
The program consists of 4 sub-products :
- Initiator App <br />
  The initiator app is located in the root directory [`./installment.py.`](./installment.py) This initiator will set up the working directory for each administrator and data entry personnel from the Engineering, QAQC, and PPC departments, as well as provide pre-filled templates ready for use. Additionally, it will generate an automated script for synchronization to support the operation of the dashboard.
- Microsoft Excel Template <br />
  The Microsoft Excel template is a workbook ready to be used and directly filled in accordance with the predefined columns and sheets.
- Script (Python and .exe) <br />
  The Python script is a data pipeline program designed to perform ETL (Extract, Transform, Load) processes for data across departments and prepare the data for visualization on the dashboard. Additionally, it also performs periodic data retrieval to update the data on the dashboard.
- Dashboard <br />
  A Streamlit-based dashboard that will display semi-real-time reports.

## Program Logic & Workflow
- PIPING PROGRESS data syncronization workflow
![presentasi 1 timas (17)](https://github.com/user-attachments/assets/4101ab7d-64e3-43f8-a148-9f350da32d58)
- PIPING PROGRESS data puller workflow
![presentasi 1 timas (18)](https://github.com/user-attachments/assets/1ff162e2-5595-45b2-b265-6138f1c7f8af)

## Prerequisites & Minimum Requirements
- Server Computer (Mandatory)
  - The server computer is running 24 hours a day, continuously
  - Intel Core i3 9th gen or above
  - At least 8 GB RAM (Synchronization.exe requires a large amount of RAM resources)
  - Windows 10 Pro/Enterprise/Server or above
  - SSD (Recommended)
- [Python 3.11.3](https://www.python.org/ftp/python/3.11.3/python-3.11.3-amd64.exe) or above (Mandatory)
- [Git](https://github.com/git-for-windows/git/releases/download/v2.48.1.windows.1/Git-2.48.1-64-bit.exe) (Optional)
- [Text Editor](https://code.visualstudio.com/download) (Optional)
- [Windows terminal](https://apps.microsoft.com/detail/9n0dx20hk701?hl=en-US&gl=US) (Optional)
- Github Account (Optional)

## How to Run
Manual Guide Video : [Click Here](https://drive.google.com/file/d/1KkjFeD_uhZxxhkOovBEi7q2EUmgueQ0R/view?usp=sharing)
- Install Python 3 and ensure that the <strong>Pyton path and PIP path are added to the environment variables.</strong>
- Clone or Pull this repository
- Install dependency on [`./requirements.txt`](./requirements.txt) by running the command `pip install -r requirements.txt`
- Run [`./installment.py`](./installment.py) by running the command `python.exe ./installment.py` on terminal.
- Choose work directories
  <p align='left'><img src = "https://github.com/user-attachments/assets/f9dbb3cd-5739-402b-91a1-959aa6881f70", width = 550> </br>
  this action will generate .exe script
  <p align='left'><img src = "https://github.com/user-attachments/assets/63e6b33c-371e-4395-a158-03be75648f07", width = 700> </br>
  and the Microsoft Excel template is already available in the selected directory and ready to be filled out. 
  <p align='left'><img src = "https://github.com/user-attachments/assets/93384150-66c8-492f-8af8-c10f748db6b1", width = 650> </br>
  <strong>Important! You are required to fill in the monitoring data in the provided Excel file. It is strictly prohibited to rename the file, move the file to another folder, or work on data input in a different Excel file, as this may cause synchronization issues.</strong>
- Run .exe script for the first time manually
  <p align='left'><img src = "https://github.com/user-attachments/assets/1c66e2a7-e22d-4a64-ab37-d0c90d611d4b", width = 700> </br>
- Set up task sheduler both for puller.exe and syncronization.exe for each plant
  <p align='left'><img src = "https://github.com/user-attachments/assets/b1071e66-fba7-4f4b-a88d-3e7dde5ff6d9", width = 700> </br>
  Set the task scheduler to run daily. It is recommended to schedule syncronization.exe to run once every day at 1:00 AM and puller.exe every 10 minutes. Then, select the executable file that was generated earlier.</br>
  <strong>Repeat steps 1-5 for each piping plant / area in the project.</strong>
- Edit dashboard path on [`./dashboard.py`](./dashboard.py)
  ![presentasi 1 timas (19)](https://github.com/user-attachments/assets/2d3cd2b5-48cb-4196-80fc-e5f497bae5e6)
  ![presentasi 1 timas (20)](https://github.com/user-attachments/assets/c94ddb88-4f48-4ac1-9f18-f9777755327f)
  ![presentasi 1 timas (21)](https://github.com/user-attachments/assets/93f710f2-1e00-46c3-8b53-565a24202cf9)
- Edit startdashboard.bat and startdashboard.vbs
  ![presentasi 1 timas (22)](https://github.com/user-attachments/assets/c7e1eea0-3443-4b36-9ed0-105ed64efd99)
- Run startdashboard.vbs manually. You can make the dashboard runs automatically after the server computer is powered on using `Win + r`, `shell:startup`, and make a shortcut to `startdashboard.vbs`.
- You can now access the dashboard by navigating to `your_ip:8501` (Streamlit runs on port 8501 by default).
### IMPORTANT NOTES
- Don't forget to open port 8501 in the firewall to ensure the dashboard can be accessed locally within the office network.
- If you want the dashboard to be accessible online, you can use [Port Forwarding](https://www.noip.com/support/knowledgebase/general-port-forwarding-guide) or a hosting service.
- Done, Awesome, now you have a dashboard 🎉🎉🎉

## Do & Don't
### Do
- Fill the data consistently
### Don't (COMPLETELY HAROM)
- Renaming the file
- Renaming the sheet
- Renaming the folder (including full path)
- Renaming columns
- Adding row above the columns
- Moving file directory

## Limitation, Restriction, and Disclaimer
- Data Structure </br>
  This program still uses Microsoft Excel for data input. Any changes to the data structure (especially in the columns), such as adding columns, removing columns, renaming columns, renaming sheet, renaming file name, or adding headers to the Excel sheet, will cause errors during synchronization.
- Data Consistency </br>
  Consistency is required in data entry (such as naming materials and lines) when inputting data. Differences and inconsistencies in the data will cause the dashboard to be unrepresentative and misleading.
- Bugs and errors </br>
  Please note that this program is still in the Beta version. Errors and bugs are highly likely to occur during operation. Feedback on errors and bugs from users is crucial for the development of this program.

## Miscellaneous
- Maintenance </br>
  It is possible that each project requires different components and values to be monitored. <strong>Since this program is open-source, each project’s IT team can modify the source code according to the management's needs.</strong>
- Log Monitoring </br>
  ![image](https://github.com/user-attachments/assets/d69f795f-ac69-4572-a947-b27818d637e8)</br>
  It is highly recommended to check the program logs daily to monitor if there are any errors when the program runs during the early hours.
- Rollback </br>
  The rollback feature (to restore Excel data to a few days back) is currently under development.

## Further Enhancements
- Potential use of Apache Airflow. </br>
  Currently, the pipeline is packaged in a .exe file, which makes data pipeline monitoring and maintenance difficult. The use of Apache Airflow is expected to simplify the development and maintenance of scripts and the pipeline.
- Potential use of Nginx (or other webserver). </br>

## Ownership
TIMAS SUPLINDO  
Developed by [Auvi. A.](https://www.linkedin.com/in/auviamril/) <br/>
Assisted by Github Copilot and OpenAI (Big Thanks to Sam Altman)
