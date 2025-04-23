import pandas as pd
import logging
import warnings
import time
import traceback
import numpy as np

def load_data(file_path, sheetname,skip_row_n, used_cols_list, add_pk = False):
    df = pd.read_excel(file_path,
                       sheet_name=sheetname,
                       skiprows=skip_row_n,
                       usecols=used_cols_list)
    if add_pk == True:
        df["PK"] =df["AREA"].astype(str) + df["LINE NO"].astype(str) + " " + df["MAT'L LINE"].astype(str) + " " + df["JOINT NO"].astype(str)
    
    return df

def clean(df, pk):
    df.drop_duplicates(subset = [pk], inplace = True, keep = 'first')
    df.dropna(subset = [pk], inplace = True)
    df = df[~df['PK'].astype(str).str.startswith('nan')]
    df[pk] = df[pk].astype(str)

    return df

def calculate_dn(P, T, type):
    if type == "SW":
        factor1 = 1 if P == "SW" else 0 if P == "FW" else None
    elif type == "FW":
        factor1 = 0 if P == "SW" else 1 if P == "FW" else None

    factor2 = {
        15: 0.5, 20: 0.75, 25: 1, 32: 1.25, 
        40: 1.5, 50: 2, 65: 2.5, 80: 3
    }.get(T, T / 25 if T >= 100 else None)

    return factor1 * factor2 if factor1 is not None and factor2 is not None else None


if __name__ == '__main__':
    allcols  =  ['LOI Date',
                'KOM Date',
                'Schedule',
                'Welding Map Date',
                'Receive Dwg',
                'Start Fabric',
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
                'P LEGEND',
                'P MATL CODE',
                'P TYPE',
                'P MATL TYPE',
                'P GROUP MATL',
                'P CLASS / SCH',
                'P BASE SIZE 1',
                'P BASE SIZE 2',
                'P THK 1',
                'P Standard Matl Name',
                'P QTY MATL',
                'P UOM',
                'S MATL CODE',
                'S TYPE',
                'S MATL TYPE',
                'S GROUP MATL',
                'S CLASS / SCH',
                'S BASE SIZE 1',
                'S BASE SIZE 2',
                'S THK 1',
                'S Standard Matl Name',
                'S QTY MATL',
                'S UOM',
                'T MATL CODE',
                'T TYPE',
                'T MATL TYPE',
                'T GROUP MATL',
                'T CLASS / SCH',
                'T BASE SIZE 1',
                'T BASE SIZE 2',
                'T THK 1',
                'T Standard Matl Name',
                'T QTY MATL',
                'T UOM',
                'MATL REMARKS',
                'P SR / MIR NO',
                'P SR DATE',
                'P PIC',
                'P QTY TOTAL',
                'P QTY HAULING',
                'P QTY OUTSTANDING',
                'P QTY MAPPING',
                'P MAPPING BALANCE',
                'S SR / MIR NO',
                'S SR DATE',
                'S PIC',
                'S QTY TOTAL',
                'S QTY HAULING',
                'S QTY OUTSTANDING',
                'S QTY MAPPING',
                'S MAPPING BALANCE',
                'T SR / MIR NO',
                'T SR DATE',
                'T PIC',
                'T QTY TOTAL',
                'T QTY HAULING',
                'T QTY OUTSTANDING',
                'T QTY MAPPING',
                'T MAPPING BALANCE',
                "MIR REMARKS",
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
                'Remarks']

    mtocols =  ["No.", 
                "LINE NO", 
                "Dwg No.",
                "Dwg Rev", 
                "PID No", 
                "PID Rev", 
                "Page", 
                "Detail", 
                "Code", 
                "Description", 
                "TYPE", 
                "MATL TYPE", 
                "GROUP MATL", 
                "BASE SIZE 1", 
                "BASE SIZE 2", 
                "CLASS / SCH", 
                "THK 1", 
                "Standard Matl Name", 
                "MATL CODE", 
                "UOM", 
                "QTY", 
                "Remark"
                ]

    mirmastercols = ["No Item",
                    "Project/Vendor Code",
                     "NS",
                     "SR / MIR NO",
                     "SR DATE",
                     "Date Action",
                     "PIC",
                     "TYPE",
                     "MATL TYPE",
                     "GROUP MATL",
                     "BASE SIZE 1",
                     "BASE SIZE 2",
                     "CLASS / SCH",
                     "THK 1",
                     "Standard Matl Name",
                     "MATL CODE",
                     "Ident Code",
                     "Tag Number / PO Number",
                     "QTY TOTAL",
                     "QTY HAULING",
                     "QTY OUTSTANDING",
                     "UOM",
                     "Status Hauling",
                     "Status SR",
                     "Remarks 1",
                     "Remarks 2",
                     "SML",
                     "SMOA"
                     ]  
    
    dbexcols = ['MATL CODE',
                'TYPE',
                'MATL TYPE',
                'GROUP MATL',
                'CLASS / SCH',
                'BASE SIZE 1',
                'BASE SIZE 2',
                'THK 1',
                'Standard Matl Name',
                'UOM',]
    
    datecols = ['LOI Date', 
                'KOM Date', 
                'Welding Map Date', 
                "Receive Dwg", 
                'Start Fabric', 
                'Schedule', 
                'FIT-UP RECORD DATE', 
                'WELDING RECORD DATE',
                'QAQC AFI F/U DATE',
                'QAQC VISUAL DATE',
                'P SR DATE',
                'S SR DATE',
                'T SR DATE',
                ]
    
    for i in range (1, 11):
        mirmastercols.append(f"HAULING QTY {i}")
        mirmastercols.append(f"HAULING DATE {i}")
        mirmastercols.append(f"HAULING PIC {i}")
    engcols = allcols[:allcols.index("Comment")+1]
    constructioncols = allcols[allcols.index("FIT-UP RECORD SW"):allcols.index("W STAMP")+1]
    engallcols = engcols + constructioncols
    qccols = allcols[allcols.index("QAQC AFI F/U DATE") : allcols.index("QAQC VISUAL RES")+1]
    qcallcols = engallcols + qccols
    ppcols = allcols[allcols.index("PPC CLAIM REPORT FABS 30%"):]
    ppcallcols = qcallcols + ppcols
    progressallcols = engallcols+qccols+ppcols
    matlallcols = allcols[:allcols.index("MATL REMARKS")+1]
    matlcols = allcols[allcols.index('P LEGEND'):allcols.index('MATL REMARKS')+1]
    srallcols = allcols[:allcols.index("MIR REMARKS")+1]
    srcols = allcols[allcols.index("P SR / MIR NO"):allcols.index("MIR REMARKS")+1]
    fullmatlcols = engcols + matlcols + srcols
    mh_progress = allcols[:allcols.index("P UOM")+1] + \
                  allcols[allcols.index('P SR / MIR NO'):allcols.index('P MAPPING BALANCE')+1] + \
                  allcols[allcols.index('S MATL CODE'):allcols.index('S UOM')+1] + \
                  allcols[allcols.index('S SR / MIR NO'):allcols.index('S MAPPING BALANCE')+1] + \
                  allcols[allcols.index('T MATL CODE'):allcols.index('MATL REMARKS')+1] + \
                  allcols[allcols.index('T SR / MIR NO'):allcols.index('MIR REMARKS')+1]

    normal_datecols = datecols[:datecols.index('QAQC VISUAL DATE')+1]
    matl_datecols = datecols[datecols.index("P SR DATE"):datecols.index("T SR DATE")+1]
    first_datecols = datecols[:datecols.index('Schedule')+1]
    
    engineerdata_path = "ENGINEER_PLACEHOLDER"
    ppc_path = "PPC_PLACEHOLDER"
    qaqc_path = "QAQC_PLACEHOLDER"

    mto_path = "MTO_PLACEHOLDER"
    mir_path = "MIR_PLACEHOLDER"

    dashboard_path = "DASHBOARD_PLACEHOLDER"

    sheet_name = "Monitoring Piping"
    sheet_name4 = "DB MATL"
    sheet_name5 = "DB Matl"

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
        engalldata['DIA-INCH PLAN SW'] = engalldata.apply(lambda row: calculate_dn(row['FW / SW'], row['DN'], "SW"), axis=1)
        engalldata['DIA-INCH PLAN FW'] = engalldata.apply(lambda row: calculate_dn(row['FW / SW'], row['DN'], "FW"), axis=1)
        engalldata["TOTAL"] = engalldata['DIA-INCH PLAN SW'] + engalldata['DIA-INCH PLAN FW']
        engalldata["FIT-UP RECORD SW"] = engalldata.apply(lambda x: x["DIA-INCH PLAN SW"] if pd.notna(x["FIT-UP RECORD DATE"]) else np.nan, axis=1)
        engalldata["FIT-UP RECORD FW"] = engalldata.apply(lambda x: x["DIA-INCH PLAN FW"] if pd.notna(x["FIT-UP RECORD DATE"]) else np.nan, axis=1)
        engalldata["WELDING RECORD SW"] = engalldata.apply(lambda x: x["DIA-INCH PLAN SW"] if pd.notna(x["WELDING RECORD DATE"]) else np.nan, axis=1)
        engalldata["WELDING RECORD FW"] = engalldata.apply(lambda x: x["DIA-INCH PLAN FW"] if pd.notna(x["WELDING RECORD DATE"]) else np.nan, axis=1)
        engdata = engalldata[engcols + ["PK"]]
        ppcdata = load_data(ppc_path, sheet_name, 3, ppcallcols, True)
        ppcdata['DIA-INCH PLAN SW'] = ppcdata.apply(lambda row: calculate_dn(row['FW / SW'], row['DN'], "SW"), axis=1)
        ppcdata['DIA-INCH PLAN FW'] = ppcdata.apply(lambda row: calculate_dn(row['FW / SW'], row['DN'], "FW"), axis=1)
        ppcdata["TOTAL"] = ppcdata['DIA-INCH PLAN SW'] + ppcdata['DIA-INCH PLAN FW']
        ppcdata["FIT-UP RECORD SW"] = ppcdata.apply(lambda x: x["DIA-INCH PLAN SW"] if pd.notna(x["FIT-UP RECORD DATE"]) else np.nan, axis=1)
        ppcdata["FIT-UP RECORD FW"] = ppcdata.apply(lambda x: x["DIA-INCH PLAN FW"] if pd.notna(x["FIT-UP RECORD DATE"]) else np.nan, axis=1)
        ppcdata["WELDING RECORD SW"] = ppcdata.apply(lambda x: x["DIA-INCH PLAN SW"] if pd.notna(x["WELDING RECORD DATE"]) else np.nan, axis=1)
        ppcdata["WELDING RECORD FW"] = ppcdata.apply(lambda x: x["DIA-INCH PLAN FW"] if pd.notna(x["WELDING RECORD DATE"]) else np.nan, axis=1)
        ppcdata["PPC CLAIM REPORT FABS 30%"] = ppcdata.apply(lambda x: x["TOTAL"]*0.3 if pd.notna(x["QAQC VISUAL DATE"]) else np.nan, axis=1)
        ppcdata["PPC CLAIM REPORT INSTALL 60%"] = ppcdata.apply(lambda x: x["TOTAL"]*0.6 if pd.notna(x["QAQC VISUAL DATE"]) else np.nan, axis=1)
        ppcdata["PPC CLAIM REPORT PUNCHLIST 10%"] = ppcdata.apply(lambda x: x["TOTAL"]*0.1 if pd.notna(x["QAQC VISUAL DATE"]) else np.nan, axis=1)
        ppcdata["PPC CLAIM REPORT TOTAL"] = ppcdata["PPC CLAIM REPORT FABS 30%"] + ppcdata["PPC CLAIM REPORT INSTALL 60%"] + ppcdata["PPC CLAIM REPORT PUNCHLIST 10%"]
        ppcdata["CUMM"] = ppcdata["PPC CLAIM REPORT TOTAL"]
        ppcclaimdata = ppcdata[ppcols + ["PK"]]
        ppcdata = ppcdata[constructioncols + ["PK"]]
        matlalldata = load_data(mto_path,sheet_name,3, matlallcols, True)
        matldata = matlalldata[matlcols + ["PK"]]
        sralldata = load_data(mir_path, sheet_name, 3, srallcols, True)
        srdata = sralldata[srcols + ["PK"]]
        mtodata = load_data(mto_path, "MTO", 3, mtocols)
        mirdata = load_data(mir_path, "MASTER", 2, mirmastercols)
        qcdata = load_data(qaqc_path, sheet_name, 3, qcallcols,True)
        qcdata = qcdata[qccols + ["PK"]]
        dbmatl_mto = load_data(mto_path, sheet_name4,0, dbexcols)
        dbmatl_mir = load_data(mir_path, sheet_name5, 0,None)
        dbmatl_mto['MATL CODE'] = dbmatl_mto[["TYPE","MATL TYPE","BASE SIZE 1","BASE SIZE 2","CLASS / SCH","GROUP MATL","THK 1"]].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)
        dbmatl_mir['MATL CODE'] = dbmatl_mir[["TYPE","MATL TYPE","BASE SIZE 1","BASE SIZE 2","CLASS / SCH","GROUP MATL","THK 1"]].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)
        matlalldata = pd.merge(matlalldata, dbmatl_mto, left_on=matlalldata["P MATL CODE"].astype(str), right_on="MATL CODE", how="left")
        matlalldata = pd.merge(matlalldata, dbmatl_mto, left_on=matlalldata["S MATL CODE"].astype(str), right_on="MATL CODE", how="left", suffixes=("","_S"))
        matlalldata = pd.merge(matlalldata, dbmatl_mto, left_on=matlalldata["T MATL CODE"].astype(str), right_on="MATL CODE", how="left", suffixes=("","_T"))
        mirdata["TEMP_PK"] = mirdata[["SR / MIR NO", "MATL CODE"]].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)
        mirdata["QTY HAULING"] = mirdata[[f"HAULING QTY {i}" for i in range(1, 11)]].sum(axis=1)
        mirdata["QTY OUTSTANDING"] = mirdata["QTY TOTAL"] - mirdata["QTY HAULING"]
        mirdata = pd.merge(mirdata, dbmatl_mir, left_on="MATL CODE", right_on="MATL CODE", how = 'left', suffixes=("", "_dup"))
        mtodata = pd.merge(mtodata, dbmatl_mto, left_on="MATL CODE", right_on="MATL CODE", how = 'left', suffixes=("", "_dup")) 
        for i in dbexcols :
            matlalldata["P "+ i] = matlalldata["P "+ i].fillna(matlalldata[i])
            matlalldata["S "+ i] = matlalldata["S "+ i].fillna(matlalldata[i+"_S"])
            matlalldata["T "+ i] = matlalldata["T "+ i].fillna(matlalldata[i+"_T"])
            matlalldata.drop(columns = i, inplace=True)
            matlalldata.drop(columns = i+"_S", inplace=True)
            matlalldata.drop(columns = i+"_T", inplace=True)
            if i != "MATL CODE":
                mirdata[i] = mirdata[i].fillna(mirdata[i+"_dup"])
                mirdata.drop(columns = i+"_dup", inplace = True)
                mtodata[i] = mtodata[i].fillna(mtodata[i+"_dup"])
                mtodata.drop(columns = i+"_dup", inplace = True)
        matldata = matlalldata[matlcols+["PK"]]
        matldata["PK"] = matldata["PK"].astype(str)
        mirdata["Status Hauling"] = mirdata["QTY OUTSTANDING"].fillna(0).apply(lambda x: "OPEN" if x != 0 else "CLOSE")
        srdata["PK"] = srdata["PK"].astype(str)
        sralldata["P TEMP_PK"] = sralldata[["P SR / MIR NO", "P MATL CODE"]].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)
        sralldata["S TEMP_PK"] = sralldata[["S SR / MIR NO", "S MATL CODE"]].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)
        sralldata["T TEMP_PK"] = sralldata[["T SR / MIR NO", "T MATL CODE"]].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)

        sralldata_merge = pd.merge(sralldata, mirdata, left_on= "P TEMP_PK", right_on="TEMP_PK",  how = "left", suffixes=("","_P"))
        sralldata_merge = pd.merge(sralldata_merge, mirdata, left_on= "S TEMP_PK", right_on="TEMP_PK", how = "left", suffixes=("","_S"))
        sralldata_merged = pd.merge(sralldata_merge, mirdata, left_on= "T TEMP_PK", right_on="TEMP_PK", how = "left", suffixes=("","_T"))

        for i in ["SR DATE", "PIC", "QTY TOTAL", "QTY HAULING"]:
            sralldata_merged["P " + i] = sralldata_merged["P " + i].fillna(sralldata_merged[i])
            sralldata_merged["S " + i] = sralldata_merged["S " + i].fillna(sralldata_merged[i+"_S"])
            sralldata_merged["T " + i] = sralldata_merged["T " + i].fillna(sralldata_merged[i+"_T"])
        sralldata_merged.drop(columns=sralldata_merged.columns[sralldata_merged.columns.get_loc("TEMP_PK"):], inplace=True)
        srdata = sralldata_merged[srcols +["PK"]]
        mirdata.drop(columns = "TEMP_PK", inplace=True)

        logging.info("Standardization Data")
        engalldata = clean(engalldata, "PK")
        engdata = clean(engdata, "PK")
        ppcdata = clean(ppcdata, "PK")
        qcdata = clean(qcdata, "PK")
        matldata = clean(matldata,"PK")
        srdata = clean(srdata, "PK")
        
        logging.info("Merge Data")
        progress_merge = pd.merge(engdata, matldata, on = "PK", how = "left", sort = False)
        progress_merge = pd.merge(progress_merge, srdata, on = "PK", how = "left", sort = False)
        progress_merge = pd.merge(progress_merge, ppcdata, on = "PK", how = "left", sort = False)
        progress_merged = pd.merge(progress_merge, qcdata, on = "PK", how = "left", sort = False)
        progress_merged = pd.merge(progress_merged, ppcclaimdata, on = "PK", how = "left", sort = False)
        eng_merged = pd.merge(engalldata, matldata, on = "PK", how = "left", sort = False)
        eng_merged = pd.merge(eng_merged, srdata, on = "PK", how = "left", sort = False)
        eng_merged = pd.merge(eng_merged, qcdata, on = "PK", how = "left", sort = False)
        eng_merged = pd.merge(eng_merged, ppcclaimdata, on = "PK", how = "left", sort = False)
        eng_merged = eng_merged[progress_merged.columns.tolist()]

        logging.info("Produce Data")
        progress_merged.to_csv(f"{dashboard_path}/Ready for dashboard RT.csv", index = False)
        eng_merged.to_csv(f"{dashboard_path}/Engineer for dashboard RT.csv", index = False)
        mtodata.to_csv(f"{dashboard_path}/MTO for dashboard RT.csv", index = False)
        mirdata.to_csv(f"{dashboard_path}/MIR for dashboard RT.csv", index = False)

        logging.info("Aman bro no error")
        logging.info("TIMAS SUPLINDO - Developed by Auvi A")

    except Exception as e:
        logging.critical(f"An critical occurred: {e}")
        logging.critical(traceback.format_exc())

        for i in range(120, 0, -1):
            print(f"exit in {i}s...", end="\r")
            time.sleep(1)

    else:
        for i in range(5, 0, -1):
            print(f"exit in {i}s...", end="\r")
            time.sleep(1)

#EoL