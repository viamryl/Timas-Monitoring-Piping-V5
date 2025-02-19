import pandas as pd
import logging
import warnings
import time
import traceback

def load_data(file_path, sheetname,skip_row_n, used_cols_list, add_pk = False):
    df = pd.read_excel(file_path,
                       sheet_name=sheetname,
                       skiprows=skip_row_n,
                       usecols=used_cols_list)
    if add_pk == True:
        df["PK"] = df["AREA"].astype(str) + df["LINE NO"].astype(str) + " " + df["MAT'L LINE"].astype(str) + " " + df["JOINT NO"].astype(str)
    
    return df

def clean(df, pk):
    df.drop_duplicates(subset = [pk], inplace = True, keep = 'first')
    df.dropna(subset = [pk], inplace = True)
    df[pk] = df[pk].astype(str)

    return df

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

    dashboard_path = "DASHBOARD_PLACEHOLDER"

    sheet_name = "Monitoring Piping"

    logging.basicConfig(
        level=logging.INFO,  # Level logging hanya untuk info utama
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("puller_log_main.txt"),  # Log ke file dengan ekstensi .txt
            logging.StreamHandler()  # Log ke konsol
            ]
        )
    warnings.filterwarnings("ignore")

    try:
        logging.info("============ Starting the Program ============")

        logging.info("Load Data")
        engalldata = load_data(engineerdata_path, sheet_name, 3, engallcols, True)
        engdata = engalldata[engcols + ["PK"]]
        ppcdata = load_data(ppc_path, sheet_name, 3, allcols, True)
        ppc_conflict = ppcdata[engallcols + ["PK"]]
        ppcdata = ppcdata[constructioncols + ["PK"]]
        qcdata = load_data(qaqc_path, sheet_name, 3, qcallcols,True)
        qcdata = qcdata[qccols + ["PK"]]
        mto_data = load_data(mto_path, "B.O.M", 3, matlcols)
        mir_data = load_data(mir_path, "MATL IN", 3, mircols)
        pworks_data = load_data(pworks_path, "DATA BASE", 5, pworkscols)

        logging.info("Standardization Data")
        engalldata = clean(engalldata, "PK")
        engdata = clean(engdata, "PK")
        ppcdata = clean(ppcdata, "PK")
        qcdata = clean(qcdata, "PK")
        
        logging.info("Merge Data")
        progress_merge = pd.merge(engdata, ppcdata, on = "PK", how = "left", sort = False)
        progress_merge = pd.merge(progress_merge, qcdata, on = "PK", how = "left", sort = False)
        progress_merged = progress_merge[qcallcols + ["PK"]]
        eng_merged = pd.merge(engalldata, qcdata, on = "PK", how = "left", sort = False)
        conflict = pd.merge(ppc_conflict, pd.DataFrame(engdata["PK"]), on = "PK", how = 'outer', indicator = True)
        conflict = conflict[conflict["_merge"] == 'left_only']
        conflict = conflict[engallcols + ["PK"]]

        logging.info("Produce Data")
        progress_merged.to_csv(f"{dashboard_path}/Ready for dashboard RT.csv", index = False)
        eng_merged.to_csv(f"{dashboard_path}/Engineer for dashboard RT.csv", index = False)
        mto_data.to_csv(f"{dashboard_path}/MTO for dashboard RT.csv", index = False)
        mir_data.to_csv(f"{dashboard_path}/MIR for dashboard RT.csv", index = False)
        pworks_data.to_csv(f"{dashboard_path}/Pworks for dashboard RT.csv", index = False)
        conflict.to_csv(f"{dashboard_path}/Conflict for dashboard RT.csv", index = False)

        logging.info("Aman bro no error")
        logging.info("TIMAS SUPLINDO - Developed by Auvi A")

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