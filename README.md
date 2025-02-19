This program is developed to monitor PIPING projects at TIMAS SUPLINDO. It will synchronize the databases of the Engineering, QAQC, and PPC departments, while providing visualizations through a dashboard.

## About This Program
The program consists of 4 sub-products :
- Initiator App <br />
  The initiator app is located in the root directory ./installment.py. This initiator will set up the working directory for each administrator and data entry personnel from the Engineering, QAQC, and PPC departments, as well as provide pre-filled templates ready for use. Additionally, it will generate an automated script for synchronization to support the operation of the dashboard.
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
  1. The server computer is running 24 hours a day, continuously
  2. Intel Core i3 9th gen or above
  3. At least 8 GB RAM (Synchronization.exe requires a large amount of RAM resources)
  4. Windows 10 or above
  5. SSD (Recommended)
- [Python 3.11.3](https://www.python.org/ftp/python/3.11.3/python-3.11.3-amd64.exe) or above (Mandatory)
- [Git](https://github.com/git-for-windows/git/releases/download/v2.48.1.windows.1/Git-2.48.1-64-bit.exe) (Optional)
- [Text Editor](https://code.visualstudio.com/download) (Optional)
- Github Account (Optional)

## How to Run
- Install Python 3 and ensure that the Pyton path and PIP path are added to the environment variables.
- Clone / Pull this repository
- Install dependency on [./requirements.txt](./requirements.txt)


## Limitation, Restriction, and Disclaimer
- Data Structure </br>
  This program still uses Microsoft Excel for data input. Any changes to the data structure (especially in the columns), such as adding columns, removing columns, renaming columns, or adding headers to the Excel sheet, will cause errors during synchronization.
- Data Consistency </br>
  Consistency is required in data entry (such as naming materials and lines) when inputting data. Differences and inconsistencies in the data will cause the dashboard to be unrepresentative and misleading.
- Bugs and errors </br>
  Please note that this program is still in the Alpha Version (even not yet in Beta Version). Errors and bugs are highly likely to occur during operation. Feedback on errors and bugs from users is crucial for the development of this program.
