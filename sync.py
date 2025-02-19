import pandas as pd
from openpyxl import load_workbook
import datetime
import time
import shutil
import os
import logging
import warnings
import numpy as np
import traceback

def load_data(file_path, sheetname,skip_row_n, used_cols_list, add_pk = False):
    df = pd.read_excel(file_path,
                       sheet_name=sheetname,
                       skiprows=skip_row_n,
                       usecols=used_cols_list)
    
    if add_pk == True:
        df["PK"] = df["AREA"].astype(str) + df["LINE NO"].astype(str) + " " + df["MAT'L LINE"].astype(str) + " " + df["JOINT NO"].astype(str)
    
    return df

def update_data(df1, df2, pk):
    diff = pd.merge(df1,
                    df2,
                    on = pk,
                    how = 'outer',
                    indicator=True)
    
    missing_data = diff[diff['_merge'] != 'both'][[pk, 'LINE NO_x', 'JOINT NO_x','LINE NO_y', "JOINT NO_y", '_merge']].reset_index(drop=True)
    missing_data['Missing From'] = missing_data['_merge'].map({'left_only': 'Added', 'right_only': 'Deleted'})
    missing_data = missing_data.drop(columns=['_merge'])
    missing_data["DATE"]=datetime.datetime.now().strftime('%Y/%m/%d')
    missing_data.reset_index(drop=True,inplace=True)
    missing_data.rename(columns={
                            "LINE NO_x" : "ADDED LINE NO",
                            "JOINT NO_x" : "ADDED JOINT NO",
                            "LINE NO_y" : "DELETED LINE NO",
                            "JOINT NO_y" : "DELETED JOINT NO",
                            "Missing From" : "STATUS"
                        },inplace = True)

    return missing_data

def write_data(df, destination_path, sheetname, startrow, startcol, pk = "PK"):
    wb = load_workbook(destination_path)
    ws = wb[sheetname]

    if sheetname == "Monitoring Piping" :
        try:
            # ws.delete_rows(startrow, ws.max_row - startrow + 1)
            df.dropna(subset = pk, inplace = True)
            df = df[df[pk] != ""]
            df = df[df[pk] != "nan"]
            df = df[~df[pk].astype(str).str.startswith('nan')]
            df.reset_index(drop=True, inplace=True)
        except:
            pass
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                ws.cell(row=startrow + i, column=startcol + j, value=value)

        total_rows_in_ws = ws.max_row
        rows_to_keep = startrow + len(df) - 1
        if total_rows_in_ws > rows_to_keep:
            ws.delete_rows(rows_to_keep + 1, total_rows_in_ws - rows_to_keep)

    elif sheetname == "Updates from Engineer": 
        total_rows_in_ws = ws.max_row
        df.reset_index(drop = True, inplace = True)
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                ws.cell(row=1 + i + total_rows_in_ws, column=startcol + j, value=value)

    elif sheetname == "Conflict":
        total_rows_in_ws = ws.max_row
        df.reset_index(drop = True, inplace = True)
        for i, row, in df.iterrows():
            for j, value in enumerate(row):
                ws.cell(row = total_rows_in_ws + 1 + i, column=startcol+j, value=value)

    else :
        print("ngaco lu")
        print("yamaap")

    wb.save(destination_path)

def syncronized(df1, df2, pk1, pk2, how = 'left'):
    df1.drop_duplicates(subset=[pk1], inplace = True, keep = 'first')
    df1.dropna(subset=[pk1], inplace =True)

    df2.drop_duplicates(subset=[pk2], inplace = True, keep = 'first')
    df2.dropna(subset=[pk2], inplace =True)
    
    merged_data = pd.merge(df1,
                            df2,
                            left_on=pk1,
                            right_on=pk2,
                            how = how,
                            sort = False)
    
    merged_data[pk2] = merged_data[pk2].fillna(merged_data[pk1])

    return merged_data 

def production(df,dashboardpath,filename,dropdup = None,dropn = None, instance = False):
    if instance == False:
        df.drop_duplicates(subset=dropdup,keep = "first", inplace = True)
        df.dropna(subset = dropn, inplace=True)
        df = df[~df[dropdup].astype(str).str.startswith('nan')]
        df.to_csv(f"{dashboardpath}/{filename}.csv", index=False)
    else : 
        df.to_csv(f"{dashboardpath}/{filename}.csv", index=False)

def backup(destination_path, backup_path,filename):
    shutil.copy(destination_path,backup_path)
    os.rename(f"{backup_path}/{filename}.xlsx",
            f"{backup_path}/{filename} {datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S-%f')[:-3]}.xlsx")
    
    status = "Data Has Been Writen Succesfully"
    return status

def slicing(path):
    slashindex = path.rfind('/')
    flname_ext = path[slashindex + 1:]
    last_dot_index = flname_ext.rfind('.')
    flname = flname_ext[:last_dot_index]

    return flname

if __name__ == '__main__':
    allcols = [
            'LOI Date',
            'KOM Date',
            'Welding Map Date',
            "Receive Dwg",
            'Asbuilt Approved',
            'Start Fabric',
            'Schedule',
            'NS',
            'AREA',
            'System',
            'Priority',
            'Sub Area',
            'Drawing No.',
            'LINE NO',
            'LINE CODE',
            "MAT'L LINE",
            "MAT'L CODE",
            'PAGE',
            'REV',
            'FW / SW',
            'JOINT TYPE',
            'JOINT NO',
            'PIPE SPOOL',
            'DN',
            'DIA-INCH PLAN SW',
            'DIA-INCH PLAN FW',
            'TOTAL',
            'Comment',
            'FIT-UP RECORD SW',
            'FIT-UP RECORD FW',
            'FIT-UP RECORD DATE',
            'FU SPV',
            'WELDING RECORD SW',
            'WELDING RECORD FW',
            'WELDING RECORD DATE',
            'W SPV',
            'W STAMP',
            'QAQC AFI F/U DATE',
            'QAQC AFI F/U NO',
            'QAQC AFI F/U RES',
            'QAQC VISUAL DATE',
            'QAQC VISUAL NO',
            'QAQC VISUAL RES',
            'PPC CLAIM REPORT FABS 30%',
            'PPC CLAIM REPORT INSTALL 60%',
            'PPC CLAIM REPORT PUNCHLIST 10%',
            'PPC CLAIM REPORT TOTAL',
            'SUMMARY AFI',
            'LAST PERIOD',
            'THIS PERIOD',
            'CUMM',
            'Remarks'
            ]
    
    matlcols = ['No.',
                'LINE NO',
                'Dwg No.',
                'Dwg Rev',
                'PID No',
                'PID Rev',
                'Page',
                'Detail',
                'Code',
                'Description',
                'TYPE',
                'TYPE MATERIAL',
                'GROUP MATERIAL',
                'BASE SIZE 1',
                'BASE SIZE 2',
                'CLASS / SCH',
                'THK 1',
                'THK 2',
                'Standard Matl Name',
                'matcode',
                'UOM',
                'QTY',
                'Remark'
                ]

    mircols  = ['Date',
                'Status',
                'MIR NO',
                'TYPE',
                'TYPE MATERIAL',
                'GROUP MATERIAL',
                'BASE SIZE 1',
                'BASE SIZE 2',
                'CLASS / SCH',
                'THK 1',
                'THK 2',
                'Standard Matl Name',
                'matcode',
                'UOM ',
                'Hauling',
                'QTY',
                'Status In',
                'Area']

    pworkscols = ['NO',
                 'DWG NO',
                 'LINE NO',
                 'JOINT NO',
                 'PART 1',
                 'PART 2',
                 'SPEC 1',
                 'SPEC 2',
                ]

    engcols = allcols[:allcols.index("Comment")+1]
    engallcols = allcols[:allcols.index("W STAMP")+1]
    constructioncols = allcols[allcols.index("FIT-UP RECORD SW"):allcols.index("W STAMP")+1]
    ppcols = allcols[allcols.index("PPC CLAIM REPORT FABS 30%"):]
    qccols = allcols[allcols.index("QAQC AFI F/U DATE") : allcols.index("QAQC VISUAL RES")+1]
    qcallcols = allcols[:allcols.index("QAQC VISUAL RES")+1]
    
    engineerdata_path = "ENGINEER_PLACEHOLDER"
    ppc_path = "PPC_PLACEHOLDER"
    qaqc_path = "QAQC_PLACEHOLDER"

    mto_path = "MTO_PLACEHOLDER"
    mir_path = "MIR_PLACEHOLDER"
    pworks_path = "PWORKS_PLACEHOLDER"

    masterdata_path = "MASTERDATA_PLACEHOLDER/masterdata.xlsx"
    dashboard_path = "DASHBOARD_PLACEHOLDER"

    backupmaster_path = "MASTERDATA_PLACEHOLDER/BACKUP"
    backupprod_path = "PRODUCTION_PLACEHOLDER/BACKUP"
    backupstaging_path = "STAGINGDATA_PLACEHOLDER/BACKUP"

    sheet_name = "Monitoring Piping"
    sheet_name2 = "Updates from Engineer"
    sheet_name3 = "Conflict"

    try:
        logging.basicConfig(
        level=logging.INFO,  # Level logging hanya untuk info utama
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("pipeline_log_main.txt"),  # Log ke file dengan ekstensi .txt
            logging.StreamHandler()  # Log ke konsol
            ]
        )
        warnings.filterwarnings("ignore")

        logging.info("============ Starting the Program ============")

        logging.info("Load Data")
        engdata_all = load_data(engineerdata_path, sheet_name, 3, engallcols, True)
        engdata_joint = engdata_all[engcols + ["PK"]]
        engdata_conflict = load_data(engineerdata_path, sheet_name3, 3, engallcols, True)
        ppc_all = load_data(ppc_path, sheet_name, 3, allcols, True)
        ppc_cons = ppc_all[constructioncols + ["PK"]]
        ppc_claim = ppc_all[ppcols + ["PK"]]
        ppc_conflict = ppc_all[engallcols + ["PK"]]
        ppc_without_afi = ppc_all[engallcols + ppcols + ["PK"]]
        qc_all = load_data(qaqc_path, sheet_name, 3, qcallcols, True)
        qc_afi = qc_all[qccols + ["PK"]]
        masterdata = load_data(masterdata_path, sheet_name, 0, None)
        masterdata_updates = load_data(masterdata_path, sheet_name2, 0, None)
        mto_data = load_data(mto_path, "B.O.M", 3, matlcols)
        mir_data = load_data(mir_path, "MATL IN", 3, mircols)
        pwroks_data = load_data(pworks_path, "DATA BASE", 5, pworkscols)

        logging.info("Standardization Data Types")
        engdata_all["PK"] = engdata_all["PK"].astype(str)
        engdata_joint["PK"] = engdata_joint["PK"].astype(str)
        engdata_conflict["PK"] = engdata_conflict["PK"].astype(str)
        ppc_all["PK"] = ppc_all["PK"].astype(str)
        ppc_cons["PK"] = ppc_cons["PK"].astype(str)
        ppc_claim["PK"] = ppc_claim["PK"].astype(str)
        qc_all["PK"] = qc_all["PK"].astype(str)
        qc_afi["PK"] = qc_afi["PK"].astype(str)
        masterdata["PK"] = masterdata["PK"].astype(str)

        logging.info("Backup Data")
        backup(engineerdata_path, backupmaster_path, slicing(engineerdata_path))
        backup(ppc_path, backupprod_path, slicing(ppc_path))
        backup(qaqc_path, backupstaging_path, slicing(qaqc_path))
        backup(mto_path, backupmaster_path, slicing(mto_path))
        backup(mir_path, backupmaster_path, slicing(mir_path))
        backup(pworks_path, backupstaging_path, slicing(pworks_path))

        logging.info("Process Data")
        updates = update_data(engdata_joint, masterdata, "PK")
        md_cons_data = syncronized(engdata_joint, ppc_cons, "PK", "PK", "left")
        md_afi_data = syncronized(md_cons_data,qc_afi, "PK", "PK", "left")
        md_claim_data = syncronized(md_afi_data,ppc_claim, "PK", "PK", "left")
        md = md_claim_data[allcols + ["PK"]]
        md_for_ppc = syncronized(ppc_without_afi, md[qccols+["PK"]], "PK", "PK", 'left')
        md_for_ppc = md_for_ppc[allcols]
        md_for_qc = md_afi_data[qcallcols]
        conflict = pd.merge(ppc_conflict, pd.DataFrame(engdata_joint["PK"]), on = "PK", how = 'outer', indicator = True)
        conflict = conflict[conflict["_merge"] == 'left_only']
        conflict = conflict[engallcols + ["PK"]]
        still_conflict = pd.merge(pd.DataFrame(engdata_conflict["PK"]), conflict, on = "PK",how = "outer", indicator=True)
        still_conflict = still_conflict[still_conflict["_merge"] == "right_only"]
        still_conflict = still_conflict[engallcols]
        engqc = syncronized(engdata_all, md_afi_data[qccols+["PK"]], "PK", "PK", 'left')

        logging.info("Writing Data")
        write_data(md, masterdata_path, sheet_name, 2, 1, "PK")
        write_data(md_for_qc, qaqc_path, sheet_name, 5, 1)
        write_data(md_for_ppc, ppc_path, sheet_name, 5, 1)
        write_data(still_conflict, engineerdata_path, sheet_name3, 5, 1)
        write_data(updates, masterdata_path, sheet_name2, 2,1)
        write_data(updates, ppc_path, sheet_name2, 2,1)
        write_data(updates, qaqc_path, sheet_name2, 2,1)

        logging.info("Produce Data")
        production(md, dashboard_path, "Ready for dashboard", "PK", "PK", False)
        production(engqc, dashboard_path, "Engineer for dashboard", "PK", "JOINT NO")
        production(conflict, dashboard_path, "Conflict for dashboard", "PK", "PK",False)
        production(mto_data, dashboard_path, "MTO for dashboard", instance=True)
        production(mir_data, dashboard_path, "MIR for dashboard", instance=True)
        production(pwroks_data, dashboard_path, "Pworks for dashboard", instance=True)

        logging.info("Aman bro no error")
        logging.info("TIMAS SUPLINDO - Developed by Auvi A")

    except PermissionError as epe:
        error_message = "Close destination file and try again"
        logging.error(f"An error occurred: {epe}")
        logging.error(error_message)

        for i in range(120, 0, -1):
            print(f"exit in {i}s...", end="\r")
            time.sleep(1)
        pass

    except Exception as e:
        logging.critical("An critical occurred: {e}")
        logging.critical(traceback.format_exc())

        for i in range(120, 0, -1):
            print(f"exit in {i}s...", end="\r")
            time.sleep(1)

    else:
        for i in range(5, 0, -1):
            print(f"exit in {i}s...", end="\r")
            time.sleep(1)

#EoL