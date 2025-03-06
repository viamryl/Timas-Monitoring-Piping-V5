import streamlit as st
import pandas as pd
import locale
import hmac
import time
import os
import datetime
from streamlit_theme import st_theme
import numpy as np

locale.setlocale(locale.LC_ALL, '')

st.set_page_config(layout="wide",
                   page_title="PIPING MONITORING",
                   initial_sidebar_state="expanded",
                   page_icon = "assets/timaslogo.ico",
                   menu_items =  {
                       "about" : "### TIMAS SUPLINDO\n\n*Developed by Auvi Amril*\n\n*https://linkedin.com/in/auviamril*\n\n\n\n"
                   })

st.markdown("""
            <style>
            [data-testid="stMetric"] {
                display: flex;
                flex-direction: column;
                align-items: center;
                text-align: center;
            }

            /* Label metric (judul atas) */
            [data-testid="stMetricLabel"] {
                text-align: center;
                justify-content: center;
            }
            button[data-baseweb="tab"] {
                font-size: 24px;
                margin: 0;
                width: 100%;
                }

            [data-testid="stMarkdownContainer"] {
                font-size: 12px;
            }
            
            [data-testid="stMetricValue"] {
                font-size: 18px;
                text-align: center;
                justify-content: center;
            }
                    
            [data-testid="stMetricDelta"]{
                font-size : 13px;
                text-align: center;
                justify-content: center;
            }
                    
            [data-testid="stMetricDelta"] svg {
                display: none;
            }

            [data-testid="stHeading"]{
                text-align: center;
                justify-content: center;
                font-size: 18px;
            }

            [data-testid="stMarkdown"]{
                text-align: center;
                justify-content: center;
            }
            [data-testid="stSidebarHeader"]{
                padding-bottom: 0.5rem;
            }

            [data-testid="stPopoverBody"] {
                width: 680px !important;
                max-width: 1500px !important;
                height: 600px !important;
                overflow: auto;
                border-radius: 10px;
                box-shadow: 0px 4px 10px rgba(0,0,0,0.2);
            }

            .footer {
                position: fixed;
                bottom: 0;
                left: 0;
                width: 100%;
                background-color: rgba(10,10,10,0) !important;
                text-align: center;
                padding: 10px;
                font-size: 11px;
                color: #6e706e;;
            }

            .css-hi6a2p {padding-top: 0rem;}
            .block-container {
                    padding-top: 1rem;
                    padding-bottom: 3rem;
                    padding-left: 5rem;
                    padding-right: 5rem;
                }
            </style>
            <div class="footer">
                TIMAS SUPLINDO&nbsp;&nbsp; •&nbsp;&nbsp; Developed by Auvi.A.
            </div>
            """, unsafe_allow_html=True)


if "password_correct" not in st.session_state:
    st.session_state.password_correct = False
    st.session_state.login_time = 0

def check_password():
    """Returns `True` if the user has the correct password and session is still valid."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if hmac.compare_digest(st.session_state["password"], st.secrets["password"]):
            st.session_state.password_correct = True
            st.session_state.login_time = time.time() 
            st.session_state["password"] = "" 
        else:
            st.session_state.password_correct = False

    if st.session_state.password_correct:
        if time.time() - st.session_state.login_time < st.secrets["expiration_time"]:
            return True
        else:
            st.session_state.password_correct = False 

    st.text_input("Password", type="password", on_change=password_entered, key="password")
    if not st.session_state.password_correct:
        st.error("😕 Password incorrect or session expired")

    return False

if not check_password():
    st.stop()

def rtdata(source_plant):
    plant = ["NaOH", "C6H6", "Crude Oil"]
    mto_path = ["D://Dashboard/Pertamina/Monitoring Piping NaOH/MTO for dashboard RT.csv", 
                "D://Dashboard/Pertamina/Monitoring Piping C6H6/MTO for dashboard RT.csv",
                "D://Dashboard/Pertamina/Monitoring Piping Crude Oil/MTO for dashboard RT.csv"]
    mir_path = ["D://Dashboard/Pertamina/Monitoring Piping NaOH/MIR for dashboard RT.csv",
                "D://Dashboard/Pertamina/Monitoring Piping C6H6/MIR for dashboard RT.csv",
                "D://Dashboard/Pertamina/Monitoring Piping Crude Oil/MIR for dashboard RT.csv"]
    pworks_path = ["D://Dashboard/Pertamina/Monitoring Piping NaOH/Pworks for dashboard RT.csv",
                   "D://Dashboard/Pertamina/Monitoring Piping C6H6/Pworks for dashboard RT.csv",
                   "D://Dashboard/Pertamina/Monitoring Piping Crude Oil/Pworks for dashboard RT.csv"]

    plant_index = plant.index(source_plant)

    mto = pd.read_csv(mto_path[plant_index])
    mir = pd.read_csv(mir_path[plant_index])
    pworks = pd.read_csv(pworks_path[plant_index])

    path = mto_path[plant_index]

    return mto, mir, pworks, path

def lastdaydata(source_plant):
    plant = ["NaOH", "C6H6", "Crude Oil"]
    mto_path = ["D://Dashboard/Pertamina/Monitoring Piping NaOH/MTO for dashboard.csv",
                "D://Dashboard/Pertamina/Monitoring Piping C6H6/MTO for dashboard.csv",
                "D://Dashboard/Pertamina/Monitoring Piping Crude Oil/MTO for dashboard.csv"]
    mir_path = ["D://Dashboard/Pertamina/Monitoring Piping NaOH/MIR for dashboard.csv",
                "D://Dashboard/Pertamina/Monitoring Piping C6H6/MIR for dashboard.csv",
                "D://Dashboard/Pertamina/Monitoring Piping Crude Oil/MIR for dashboard.csv"]
    pworks_path = ["D://Dashboard/Pertamina/Monitoring Piping NaOH/PWorks for dashboard.csv",
                   "D://Dashboard/Pertamina/Monitoring Piping C6H6/PWorks for dashboard.csv",
                   "D://Dashboard/Pertamina/Monitoring Piping Crude Oil/PWorks for dashboard.csv"]

    plant_index = plant.index(source_plant)

    mto = pd.read_csv(mto_path[plant_index])
    mir = pd.read_csv(mir_path[plant_index])
    pworks = pd.read_csv(pworks_path[plant_index])

    path = mto_path[plant_index]

    return mto, mir, pworks, path

def piprog_lastday(source_plant, dashboard = "main"):
    plant = ["NaOH", "C6H6", "Crude Oil"]
    if dashboard == "main":
        piping_progress_path = ["D://Dashboard/Pertamina/Monitoring Piping NaOH/Ready for dashboard.csv",
                                "D://Dashboard/Pertamina/Monitoring Piping C6H6/Ready for dashboard.csv",
                                "D://Dashboard/Pertamina/Monitoring Piping Crude Oil/Ready for dashboard.csv"]
    elif dashboard == "eng":
        piping_progress_path = ["D://Dashboard/Pertamina/Monitoring Piping NaOH/Engineer for dashboard.csv",
                                "D://Dashboard/Pertamina/Monitoring Piping C6H6/Engineer for dashboard.csv",
                                "D://Dashboard/Pertamina/Monitoring Piping Crude Oil/Engineer for dashboard.csv"]
    
    plant_index = plant.index(source_plant)
    piping_progress = pd.read_csv(piping_progress_path[plant_index])

    return piping_progress

def piprog_newest(source_plant, dashboard = "main"):
    plant = ["NaOH", "C6H6", "Crude Oil"]
    if dashboard == "main":
        piping_progress_path = ["D://Dashboard/Pertamina/Monitoring Piping NaOH/Ready for dashboard RT.csv",
                                "D://Dashboard/Pertamina/Monitoring Piping C6H6/Ready for dashboard RT.csv",
                                "D://Dashboard/Pertamina/Monitoring Piping Crude Oil/Ready for dashboard RT.csv"]
    elif dashboard == "eng":
        piping_progress_path = ["D://Dashboard/Pertamina/Monitoring Piping NaOH/Engineer for dashboard RT.csv",
                                "D://Dashboard/Pertamina/Monitoring Piping C6H6/Engineer for dashboard RT.csv",
                                "D://Dashboard/Pertamina/Monitoring Piping Crude Oil/Engineer for dashboard RT.csv"]
    
    plant_index = plant.index(source_plant)
    piping_progress = pd.read_csv(piping_progress_path[plant_index])

    return piping_progress

st.title("PIPING PROGRESS", anchor = False)

if "selected_line" not in st.session_state:
    st.session_state.selected_line = "All"
if "selected_joint" not in st.session_state:
    st.session_state.selected_joint = "All"
if "selected_TYPE" not in st.session_state:
    st.session_state.selected_TYPE = "All"
if "selected_matl" not in st.session_state:
    st.session_state.selected_matl = "All"
if "selected_group" not in st.session_state:
    st.session_state.selected_group = "All"
if "selected_class" not in st.session_state:
    st.session_state.selected_class = "All"
if "selected_size" not in st.session_state:
    st.session_state.selected_size = "All"
if "selected_item" not in st.session_state:
    st.session_state.selected_item = "All"
if "selected_mir" not in st.session_state:
    st.session_state.selected_mir = "All"

cons_tab, matl_tab = st.tabs(["Construction", "Material Handling"])

with cons_tab:
    with st.sidebar:
        theme = st_theme()["base"]
        if theme == "dark":
            st.image("assets/timasputih.png")
        else : 
            st.image("assets/timaspanjang.png")
        st.title("Options")
        
        plant = st.selectbox("Choose Plant", options = ["NaOH", "C6H6", "Crude Oil"])

        datasource = st.selectbox("Choose Data Source", options = ["Last Day Data", "Newest Data"])
        if datasource == "Last Day Data":
            mto, mir, pworks, path = lastdaydata(plant)
        elif datasource == "Newest Data":
            mto, mir, pworks, path = rtdata(plant)

        dashboard = st.radio(label = "Choose Dashboard",options = ["Main Dashboard", "Engineer Dashboard"], label_visibility="visible", horizontal=False)

        if dashboard == "Main Dashboard":
            if datasource == "Last Day Data":
                piping_progress = piprog_lastday(plant, dashboard="main") 
            elif datasource == "Newest Data":
                piping_progress = piprog_newest(plant, dashboard="main")
        elif dashboard == "Engineer Dashboard":
            if datasource == "Last Day Data":
                piping_progress = piprog_lastday(plant, dashboard="eng")
            elif datasource == "Newest Data":
                piping_progress = piprog_newest(plant, dashboard="eng")
        st.container(height = 130, border = False)
        if st.button("Reset Filter"):
            st.session_state.selected_line = "All"
            st.session_state.selected_joint = "All"
            st.session_state.selected_TYPE = "All"
            st.session_state.selected_matl = "All"
            st.session_state.selected_group = "All"
            st.session_state.selected_class = "All"
            st.session_state.selected_size = "All"
            st.session_state.selected_item = "All"
            st.session_state.selected_mir = "All"
            st.rerun()

        last_mondified = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%b %d, %Y %H:%M:%S")
        st.write(f"Last Pull ⇥ {last_mondified}")

    mto_pure = mto
    piping_progress["TOTAL"] = piping_progress["DIA-INCH PLAN SW"] + piping_progress["DIA-INCH PLAN FW"]
    piping_progress["TOTAL"] = pd.to_numeric(piping_progress["TOTAL"])
    piping_progress["FIT-UP RECORD DATE"] = pd.to_datetime(piping_progress["FIT-UP RECORD DATE"], format = "mixed")
    piping_progress["FIT-UP RECORD DATE"] = piping_progress["FIT-UP RECORD DATE"].dt.strftime("%B %d, %Y")
    piping_progress["WELDING RECORD DATE"] = pd.to_datetime(piping_progress["WELDING RECORD DATE"], format = "mixed")
    piping_progress["WELDING RECORD DATE"] = piping_progress["WELDING RECORD DATE"].dt.strftime("%B %d, %Y")
    piping_progress["QAQC AFI F/U DATE"] = pd.to_datetime(piping_progress["QAQC AFI F/U DATE"], format = "mixed")
    piping_progress["QAQC AFI F/U DATE"] = piping_progress["QAQC AFI F/U DATE"].dt.strftime("%B %d, %Y")
    piping_progress["QAQC VISUAL DATE"] = pd.to_datetime(piping_progress["QAQC VISUAL DATE"], format = "mixed")
    piping_progress["QAQC VISUAL DATE"] = piping_progress["QAQC VISUAL DATE"].dt.strftime("%B %d, %Y")

    cons_col, afi_col = st.columns([3, 1])

    selected_line = cons_col.selectbox("Choose Line", ["All"] + piping_progress["LINE NO"].unique().tolist(), key="selected_line")
    if selected_line != "All":
        piping_progress = piping_progress[piping_progress["LINE NO"] == selected_line]
        selected_joint = afi_col.selectbox("Choose Joint", ["All"] + piping_progress["JOINT NO"].unique().tolist(),key="selected_joint")
    else:
        piping_progress = piping_progress
        st.session_state.selected_joint = "All"
        selected_joint = afi_col.selectbox("Choose Joint", ["All"] + piping_progress["JOINT NO"].unique().tolist(), key="selected_joint", disabled=True)

    if selected_joint != "All": 
        piping_progress = piping_progress[piping_progress["JOINT NO"] == selected_joint]
    else:
        piping_progress = piping_progress

    if selected_line != "All":
        mto = mto[mto["LINE NO"] == selected_line]
        pworks = pworks[pworks["LINE NO"] == selected_line]
    else:
        mto = mto  # Tampilkan semua jika "All"
        pworks = pworks

    #CONSTRUCTION
    left_date=piping_progress[["LOI Date", 'KOM Date', "Welding Map Date",'Receive Dwg']].head(1).T
    left_date.columns = ["Date"]
    right_date=piping_progress[["Asbuilt Approved","Start Fabric", 'Schedule']].head(1).T
    right_date.columns = ["Date"]

    unfitted_df = piping_progress[piping_progress['FIT-UP RECORD DATE'].isnull()][["LINE NO", "JOINT NO", "Sub Area", "MAT'L LINE", "TOTAL"]]
    unwelded_df = piping_progress[piping_progress['WELDING RECORD DATE'].isnull()][["LINE NO", "JOINT NO", "Sub Area", "MAT'L LINE", "TOTAL"]]
    ssdone_df = piping_progress[piping_progress["MAT'L LINE"] != "CS LINE"].loc[piping_progress["WELDING RECORD DATE"].notna()][["LINE NO", "Sub Area", "JOINT NO", "TOTAL"]]
    csdone_df = piping_progress[piping_progress["MAT'L LINE"] == "CS LINE"].loc[piping_progress["WELDING RECORD DATE"].notna()][["LINE NO", "Sub Area", "JOINT NO", "TOTAL"]]
    ssnotdone_df = piping_progress[piping_progress["MAT'L LINE"] != "CS LINE"].loc[piping_progress["WELDING RECORD DATE"].isna()][["LINE NO", "Sub Area", "JOINT NO", "TOTAL"]]
    csnotdone_df = piping_progress[piping_progress["MAT'L LINE"] == "CS LINE"].loc[piping_progress["WELDING RECORD DATE"].isna()][["LINE NO", "Sub Area", "JOINT NO", "TOTAL"]]
    backlog_df = piping_progress.loc[piping_progress["FIT-UP RECORD DATE"].notna() & piping_progress["WELDING RECORD DATE"].isna(), ["LINE NO", "JOINT NO", "Sub Area", "MAT'L LINE", "TOTAL"]]

    total_joint = piping_progress.shape[0]
    total_id = piping_progress["TOTAL"].sum()
    total_line = piping_progress["LINE NO"].nunique()

    unfitted_joint = unfitted_df.shape[0]
    unfitted_id = unfitted_df["TOTAL"].sum()
    fitted_joint = total_joint - unfitted_joint
    fitted_id = total_id - unfitted_id

    unwelded_joint = unwelded_df.shape[0]
    unwelded_id = unwelded_df["TOTAL"].sum()
    welded_joint = total_joint - unwelded_joint
    welded_id = total_id - unwelded_id

    ssline = piping_progress.loc[piping_progress["MAT'L LINE"] != "CS LINE", "LINE NO"].nunique()
    ssdone = (piping_progress[piping_progress["MAT'L LINE"] != "CS LINE"].groupby("LINE NO")["WELDING RECORD DATE"].apply(lambda x: x.notna().all())).sum()
    ssnotdone = ssline - ssdone
    csline = piping_progress.loc[piping_progress["MAT'L LINE"] == "CS LINE", "LINE NO"].nunique()
    csdone = (piping_progress[piping_progress["MAT'L LINE"] == "CS LINE"].groupby("LINE NO")["WELDING RECORD DATE"].apply(lambda x: x.notna().all())).sum()
    csnotdone = csline - csdone

    joint_backlog = fitted_joint - welded_joint
    joint_id = fitted_id - welded_id

    #AFI
    afifujoint = piping_progress["QAQC AFI F/U DATE"].count()
    afifujoint_notyet = total_joint - afifujoint
    afifuid = sum(piping_progress[piping_progress["QAQC AFI F/U DATE"].apply(pd.notna)]["TOTAL"].tolist())
    afifuid_notyet = total_id - afifuid

    afivisjoint = piping_progress["QAQC VISUAL DATE"].count()
    afivisjoint_notyet = total_joint - afivisjoint
    afivisid = sum(piping_progress[piping_progress["QAQC VISUAL DATE"].apply(pd.notna)]["TOTAL"].tolist())
    afivisid_notyet = total_id - afivisid

    #Material
    summary_matl = mto_pure.groupby("Standard Matl Name").agg(Total = ("QTY", "sum"), TYPE = ("TYPE", "first"), matl_code = ("TYPE MATERIAL", "first"), sch = ("CLASS / SCH", "first"), size = ("BASE SIZE 1", "first"),item_criteria = ("Standard Matl Name", "first"))
    summary_matl = summary_matl[["TYPE", "matl_code", "item_criteria", "sch", "size", "Total"]]
    summary_matl = summary_matl.rename(columns = {"matl_code" : "Matl Code",
                                                "item_criteria": "Standard Matl Name",
                                                "sch" : "CLASS / SCH",
                                                "size" : "Size",
                                                "Total" : "Total"})
    summary_matl = summary_matl.sort_values(by = "Total", ascending = False)
    summary_matl = summary_matl.reset_index(drop=True)

    summary_mir = mir.groupby("Standard Matl Name").agg(QTY = ("QTY", "sum"),Hauling = ("Hauling", "sum"), item_criteria = ("Standard Matl Name", "first"))
    summary_mir = summary_mir.rename(columns = {"item_criteria" : "STANDARD MATL NAME"})

    summary_merged = pd.merge(summary_matl,
                        summary_mir,
                        left_on= "Standard Matl Name",
                        right_on = "STANDARD MATL NAME",
                        how = "left")
    summary_merged["Balance MTO"] =  summary_merged["QTY"] - summary_merged["Total"]
    summary_merged["BALANCE MIR"] = summary_merged["Hauling"] - summary_merged["QTY"]
    summary_merged = summary_merged.rename(columns = {"TYPE_x" : "TYPE",
                                                    "CLASS / SCH_x" : "CLASS / SCH",
                                                    "Total" : "MTO Total",
                                                    "QTY" : "MIR Total",
                                                    "Hauling" : "MIR Hauling"})
    summary_merged.drop(columns = "STANDARD MATL NAME", inplace=True)

    mir = mir.rename(columns = {"QTY" : "QTY"})


    ###########################################################################
    ############################### L A Y O U T ###############################
    ###########################################################################


    #Construction Container
    with cons_col :
        cons_container = st.container(border = True)
        cons_container.subheader("CONSTRUCTION MONITORING", anchor = False)

        date_container = cons_container.container()
        left_col, right_col = date_container.columns(2)
        left_col.dataframe(left_date, width=2000, column_config={"Date" : st.column_config.DatetimeColumn(format="MMMM DD, YYYY")})
        right_col.dataframe(right_date, width=2000, column_config={"Date" : st.column_config.DatetimeColumn(format="MMMM DD, YYYY")})

        joint_col, id_col, line_col = cons_container.columns(3)
        joint_col.metric(label="Total Joint", value="{:,.0f}".format(total_joint), border=True)
        joint_col.metric(label="Fitted Joint", value="{:,.0f}".format(fitted_joint), delta="{:,.0f} Joint Not Fitted Yet".format(unfitted_joint), delta_color="inverse", border=True)
        joint_col.metric(label="Welded Joint", value="{:,.0f}".format(welded_joint), delta= "{:,.0f} Joint Not Welded Yet".format(unwelded_joint), delta_color="inverse", border=True)
        
        id_col.metric(label="Total ID", value="{:,.2f}".format(total_id), border=True)
        id_col.metric(label="Fitted ID", value=f"{fitted_id:,.2f}", delta= f"{unfitted_id:,.2f} ID Not Fitted Yet", delta_color="inverse", border=True)
        id_col.metric(label="Welded ID", value= f"{welded_id:,.2f}",  delta=f"{unwelded_id:,.2f} ID Not Welded Yet", delta_color="inverse", border=True)
        
        line_col.metric(label="Total Line", value="{:,.0f}".format(total_line), border=True)
        line_col.metric(label="Stainless Steel Line", value=f"{ssline:,.0f}", delta = f"{ssnotdone:,.0f} Lines Not Done Yet",delta_color="inverse", border=True)
        line_col.metric(label="Carbon Steel Line", value=f"{csline:,.0f}", delta = f"{csnotdone:,.0f} Lines Not Done Yet", delta_color="inverse", border=True)

        backlog_container = cons_container.container()
        joint_backlog_col, id_backlog_col = backlog_container.columns(2)
        joint_backlog_col.metric(label="Joint Backlog", value="{:,.0f} Joint".format(fitted_joint - welded_joint), border=True)
        id_backlog_col.metric(label="ID Backlog", value="{:,.2f} ID".format(fitted_id - welded_id), border=True)

        button_container = cons_container.container()
        ufj,uwj,ssd,csd,ssnd,csnd,bl = button_container.columns(7)

        with ufj.popover("Unfitted Joint", use_container_width = True):
            st.markdown("Unfitted Line & Joint")
            dfcont = st.container(border=True)
            dfcont.dataframe(unfitted_df)

        with uwj.popover("Unwelded Joint", use_container_width = True):
            st.markdown("Unwelded Line & Joint")
            dfcont = st.container(border=True)
            dfcont.dataframe(unwelded_df)

        with ssd.popover("SS Done", use_container_width = True):
            st.markdown("Stainless Steel Line Done")
            dfcont = st.container(border=True)
            dfcont.dataframe(ssdone_df)

        with csd.popover("CS Done", use_container_width = True):
            st.markdown("Carbon Steel Line Done")
            dfcont = st.container(border=True)
            dfcont.dataframe(csdone_df)
        
        with ssnd.popover("SS Not Done Yet", use_container_width = True):
            st.markdown("Stainless Steel Line Not Done Yet")
            dfcont = st.container(border=True)
            dfcont.dataframe(ssnotdone_df)

        with csnd.popover("CS Not Done Yet", use_container_width = True):
            st.markdown("Carbon Steel Line Not Done Yet")
            dfcont = st.container(border=True)
            dfcont.dataframe(csnotdone_df)
        
        with bl.popover("Backlog", use_container_width = True):
            st.markdown("Unfitted Line & Joint")
            dfcont = st.container(border=True)
            dfcont.dataframe(backlog_df)

    with afi_col:
        afi_container = st.container(border=True)
        afi_container.subheader("AFI MONITORING", anchor = False)

        afi_container.metric(label = "AFI F/U Approved (Joint)", value = f"{afifujoint:,.0f}",delta = "Unapproved : {:,.0f} from {:,.0f}".format(afifujoint_notyet, total_joint), delta_color="inverse", border = True)
        afi_container.metric(label = "AFI F/U Approved (ID)", value = f"{afifuid:,.2f}",delta = "Unapproved : {:,.2f} from {:,.2f}".format(afifuid_notyet, total_id), delta_color="inverse", border = True)
        afi_container.metric(label = "AFI Visual Approved (Joint)", value = f"{afivisjoint:,.0f}",delta = "Unapproved : {:,.0f} from {:,.0f}".format(afivisjoint_notyet, total_joint), delta_color="inverse", border = True)
        afi_container.metric(label = "AFI Visual Approved (ID)", value = f"{afivisid:,.2f}",delta = "Unapproved : {:,.2f} from {:,.2f}".format(afivisid_notyet, total_id), delta_color="inverse", border = True)
        
        detail_container = st.container(border = True)
        detail_container.markdown("###### Details")

        if selected_joint != "All" : 
            with detail_container.popover("Joint Detail", use_container_width = True, ):
                st.markdown("Joint Detail")
                dfcont = st.container(border = True)
                joint_detail = piping_progress[["Sub Area", "LINE NO", "Drawing No.", "REV", "PAGE", "JOINT NO", "JOINT TYPE", "PIPE SPOOL", "DIA-INCH PLAN SW", "DIA-INCH PLAN FW", "Comment"]]
                joint_detail = joint_detail.T
                joint_detail.columns = ["Detail"]
                dfcont.dataframe(joint_detail, width = 2000) 

            with detail_container.popover("Construction Detail", use_container_width = True):
                st.markdown("Construction Detail")
                dfcont = st.container(border = True)
                cons_detail = piping_progress[["FIT-UP RECORD DATE", "FIT-UP RECORD SW", "FIT-UP RECORD FW", "FU SPV", "WELDING RECORD DATE", "WELDING RECORD SW", "WELDING RECORD FW", "W SPV", "W STAMP"]]
                cons_detail = cons_detail.T
                cons_detail.columns = ["Detail"]
                dfcont.dataframe(cons_detail, width = 2000)

            with detail_container.popover("AFI Detail", use_container_width = True):
                st.markdown("AFI Detail")
                dfcont = st.container(border = True)
                afidetail = piping_progress[["QAQC AFI F/U DATE", "QAQC AFI F/U NO", "QAQC AFI F/U RES", "QAQC VISUAL DATE", "QAQC VISUAL NO", "QAQC VISUAL RES"]]

                afidetail = afidetail.T
                afidetail.columns = ["Detail"]
                dfcont.dataframe(afidetail, width = 2000)
        
        else:
            with detail_container.popover("Joint Detail", use_container_width = True, disabled=True):
                st.markdown("Joint Detail")

            with detail_container.popover("Construction Detail", use_container_width = True, disabled=True):
                st.markdown("Construction Detail")

            with detail_container.popover("AFI Detail", use_container_width = True, disabled=True):
                st.markdown("AFI Detail")

    matl_container = st.container(border = True)
    with matl_container:
        matl_container.subheader("MATERIAL MONITORING", anchor = False)
        matl_filter_container = st.container()
        left_col, middleleft_col, middle_col,middleright_col, right_col = matl_filter_container.columns(5)

        selected_TYPE = left_col.selectbox("Choose TYPE", ["All"] + mto["TYPE"].unique().tolist(),key="selected_TYPE")
        if selected_TYPE != "All":
            mto = mto[mto["TYPE"] == selected_TYPE]
        else:
            mto = mto

        selected_matl= middleleft_col.selectbox("Choose Material Code", ["All"] + mto["TYPE MATERIAL"].unique().tolist(),key="selected_matl")
        if selected_matl != "All":
            mto = mto[mto["TYPE MATERIAL"] == selected_matl]
        else:
            mto = mto

        selected_group= middle_col.selectbox("Choose Material Group", ["All"] + mto["GROUP MATERIAL"].unique().tolist(),key="selected_group")
        if selected_group != "All":
            mto = mto[mto["GROUP MATERIAL"] == selected_group]
        else:
            mto = mto

        selected_class= middleright_col.selectbox("Choose CLASS / SCH", ["All"] + mto["CLASS / SCH"].unique().tolist(),key="selected_class")
        if selected_class != "All":
            mto = mto[mto["CLASS / SCH"] == selected_class]
        else:
            mto = mto

        selected_size= right_col.selectbox("Choose Size", ["All"] + mto["BASE SIZE 1"].unique().tolist(),key="selected_size")
        if selected_size != "All":
            mto = mto[mto["BASE SIZE 1"] == selected_size]
        else:
            mto = mto

        itemsfilter_col, mirfilter_cols = matl_container.columns([3, 1])
        selected_item = itemsfilter_col.selectbox("Choose Standard Matl Name", ["All"] + mto["Standard Matl Name"].unique().tolist(),key="selected_item")
        if selected_item != "All":
            mto = mto[mto["Standard Matl Name"] == selected_item]
        else:
            mto = mto

        selected_mir = mirfilter_cols.selectbox("Choose MIR Number", ["All"] + mir["MIR NO"].unique().tolist(),key="selected_mir")
        if selected_mir != "All":
            mir = mir[mir["MIR NO"] == selected_mir]
        else:
            mir = mir

        item_unique = mto["Standard Matl Name"].unique()
        mir = mir[mir["Standard Matl Name"].isin(item_unique)]
        summary_merged = summary_merged[summary_merged["Standard Matl Name"].isin(item_unique)]

        matl_container.markdown("##### B.O.M")
        mto1 = mto[["LINE NO", "TYPE","TYPE MATERIAL",  "Standard Matl Name", "CLASS / SCH", "BASE SIZE 1","BASE SIZE 2", "UOM", "QTY" ]].reset_index(drop = True)
        event = st.dataframe(mto1, selection_mode="single-row", key = "data", on_select="rerun", width=2000, height = 400, hide_index=True)

        try:
            vardewa = mto1.loc[event.selection["rows"][0]]["Standard Matl Name"]
        except:
            vardewa = None

        if vardewa is not None :
            mir = mir[mir["Standard Matl Name"] == vardewa]
            mto = mto[mto["Standard Matl Name"] == vardewa]
            summary_merged = summary_merged[summary_merged["Standard Matl Name"] == vardewa]
        else:
            mir = mir
            mto = mto
            summary_merged = summary_merged

        left_col, middleleft_col, middle_col, middleright_col, right_col = matl_container.columns(5)

        left_col.metric(label = "MTO QTY", value = "{:,.2f}".format(mto["QTY"].sum()), border = True)
        middleleft_col.metric(label = "MIR QTY", value = "{:,.2f}".format(mir["QTY"].sum()), border = True)
        middle_col.metric(label = "Hauling QTY", value = "{:,.2f}".format(mir["Hauling"].sum()), border = True)
        middleright_col.metric(label = "Balance MTO", value = "{:,.2f}".format(mir["QTY"].sum()-mto["QTY"].sum()), border = True)
        right_col.metric(label = "Balance MIR", value = "{:,.2f}".format(mir["Hauling"].sum()-mir["QTY"].sum()), border = True)
        
        matl_container.markdown("##### MIR")
        st.dataframe(mir[["MIR NO", "Standard Matl Name", "QTY", "Hauling"]], width=2000, height = 360)    

        matl_container.markdown("##### Piping Works")
        st.dataframe(pworks[["LINE NO", "JOINT NO", "PART 1", "PART 2"]], width=2000, height = 280)    

        matl_container.markdown("##### Summary Material")
        st.dataframe(summary_merged, width=2000, height = 280, hide_index = True)

with matl_tab:
    matl_tab.markdown("# ^\_^ THIS PAGE IS UNDER DEVELOPMENT ^_^")
    mirexist = pd.read_csv("D:/Dashboard/Pertamina/MIR TEMPLATE - Copy.csv")

    if "mirnum" not in st.session_state:
        st.session_state.mirnum = None
    if "mir_confirmed" not in st.session_state:
        st.session_state.mir_confirmed = False
    if "dumpdf" not in st.session_state:
        st.session_state.dumpdf = pd.DataFrame(columns = ["TYPE", "TYPE Material","GRUP material", "Class / Sch", "Base size 1", "Base size 2", "Thk", "UOM","Standard Matl Name"])

    def clear_mirnum():
        st.session_state.mirnum = None
        st.rerun()

    @st.dialog("⚠ WARNING: MIR NO IS EXISTED!", width = "large")
    def check():
        st.write("Are you sure to continue?")
        if st.button("Agree"):
            st.session_state.mir_confirmed = True 
            st.rerun()
        if st.button("Cancel Input"):
            clear_mirnum()

    @st.dialog("!INFO!", width = "large")
    def hasbeensaved(what):
        st.write(f"{what} HAS BEEN SAVERD")
        if st.button("Close"):
            st.rerun()
    

    def callback():
        edited_rows = st.session_state["data_editor"]["edited_rows"]
        rows_to_delete = []

        for idx, value in edited_rows.items():
            if value["Del"] is True:
                rows_to_delete.append(idx)

        st.session_state["dumpdf"] = (
            st.session_state["dumpdf"].drop(rows_to_delete, axis=0).reset_index(drop=True)
        )

    def addmir():
        pass


    if "selected_rows" not in st.session_state:
        st.session_state.selected_rows = []
    matl_container = st.container(border=True)
    with matl_container:
        st.subheader("EXISTING MIR", anchor = False)
        mirexist["TYPE"]=mirexist["TYPE"].astype(str)
        # column_config_mirexist = {col: st.column_config.TextColumn(label=col) for col in mirexist.columns}
        mirexist = st.data_editor(mirexist, column_config={
            "TYPE" : st.column_config.TextColumn("TYPE")
        }, num_rows="dynamic")
        st.divider()
        matl_container.subheader("INPUT NEW MIR", anchor=False)

        col1, col2, col3, col4, col5, col6 = matl_container.columns(6)
        mirno = col1.text_input("MIR NO", key="mirnum") 

        if mirno and mirno.upper() in mirexist["PL/SR"].astype(str).tolist() and not st.session_state.mir_confirmed:
            check() 
        else : 
            st.session_state.mir_confirmed = False

        vendor = col2.text_input("Vendor Code")
        mirdate = col3.date_input("MIR DATE")
        actiondate = col4.date_input("Action Date")
        statussr = col5.text_input("Status SR")
        area = col6.text_input("Area")

        matl_container.subheader("MATERIAL DB", anchor = False)
        db_matl = pd.read_csv("D://Dashboard/DB MATL.csv")
        db_matl["keycode"] = (
            db_matl[["TYPE", "TYPE Material","GRUP material", "Class / Sch", "Base size 1", "Base size 2", "Thk"]]
            .fillna('')
            .astype(str)
            .apply(lambda row: ' '.join(filter(None, row)), axis=1)
        )

        db_matl = db_matl[["MatCode dev", "TYPE", "TYPE Material","GRUP material","Standard Matl Name", "Class / Sch", "Base size 1", "Base size 2", "Thk", "UOM", "keycode"]]
        
        col1, col2, col3, col4, col5, col6 = matl_container.columns(6)

        # Menampilkan editor data
        event = st.dataframe(db_matl, selection_mode="multi-row", on_select="rerun", use_container_width=True)

        matl_container.subheader("MIR PREVIEW", anchor = False)
        selected_item = db_matl.loc[event.selection["rows"]]
        selected_item = selected_item[["TYPE", "TYPE Material","GRUP material", "Class / Sch", "Base size 1", "Base size 2", "Thk", "UOM","Standard Matl Name", "MatCode dev"]]
        selected_item["No Item"] = [i+1 for i in range(len(selected_item))]
        selected_item["Qty Total"] = None
        selected_item["NS"] = mirexist["NS"].max() +1
        selected_item["Ident Code"] = None
        selected_item["Tag Number / PO Number"] = None
        selected_item["PL/SR"] = mirno
        selected_item["Project/Vendor Code"] = vendor
        selected_item["SR Date"] = mirdate
        selected_item["Date Action"] = actiondate
        selected_item["Status SR"] = statussr
        selected_item["Area"] = area
        selected_item["THK 2"] = None
        selected_item["Qty Hauling"] = None
        selected_item["Qty Outstanding"] = None
        selected_item["Status Hauling"] = None
        selected_item["Remarks 1"] = area
        selected_item["Remarks 1.1"] = None
        selected_item["SML"] = None
        selected_item["SMOA"] = None
        selected_item =selected_item[["No Item", "Project/Vendor Code", "NS", "PL/SR","SR Date", "Date Action", "TYPE", "TYPE Material", 
                                      "GRUP material",  "Base size 1", "Base size 2", "Class / Sch", "Thk", "THK 2", "Standard Matl Name", 
                                      "MatCode dev", "Ident Code", "Tag Number / PO Number", "Qty Total","Qty Hauling", "Qty Outstanding","UOM", 
                                      "Status Hauling", "Status SR", "Remarks 1", "Remarks 1.1", "SML", "SMOA"]]
        
        selected_item = selected_item.rename(columns = {
            "GRUP material" : "GROUP MATERIAL",
            "Base size 1" : "BASE SIZE 1",
            "Base size 2" : "BASE SIZE 2",
            "Class / Sch" : "CLASS / SCH",
            "Thk" : "THK 1",
            "MatCode dev" : "Matl Code"
        })
        selected_item_df= st.data_editor(selected_item, num_rows="dynamic",column_config={
            "MIR QTY" : st.column_config.NumberColumn(width="medium", required=True),
            "IDENT CODE" : st.column_config.TextColumn(),
            "PO NUMBER" : st.column_config.TextColumn()
        }, disabled = ["No Item", "Project/Vendor Code", "NS", "PL/SR","SR Date", "Date Action", "TYPE", "TYPE Material", 
                                      "GRUP material",  "Base size 1", "Base size 2", "Class / Sch", "Thk", "THK 2", "Standard Matl Name", 
                                      "MatCode dev","Qty Hauling", "Qty Outstanding","UOM", 
                                      "Status Hauling", "Status SR", "Remarks 1", "Remarks 1.1", "SML", "SMOA"])

        if st.button("Save MIR"):
            mirexist = pd.concat([mirexist, selected_item_df], ignore_index=True)
            mirexist.to_csv("D:/Dashboard/Pertamina/MIR TEMPLATE - Copy.csv", index=False)
            del st.session_state.mirnum
            hasbeensaved("MIR")

        st.divider()
        st.subheader("HAULING INPUT", anchor = False)

        col1, col2 = st.columns(2)
        
        st.session_state.haulingdf = mirexist
        mirnumber = col1.selectbox("Select MIR Number", options = ["All"] + st.session_state.haulingdf["PL/SR"].unique().tolist())

        st.session_state.haulingdf = mirexist[["No Item","NS","PL/SR", "Standard Matl Name", "Qty Total","Qty Outstanding", "Qty Hauling","UoM"]]
        if mirnumber != "All":
            st.session_state.haulingdf = st.session_state.haulingdf[st.session_state.haulingdf["PL/SR"] == mirnumber]
        else:
            st.session_state.haulingdf = st.session_state.haulingdf

        item_selected = col2.selectbox("Select Matl Name", options = ["All"] + st.session_state.haulingdf["Standard Matl Name"].unique().tolist())
        if item_selected != "All":
            st.session_state.haulingdf = st.session_state.haulingdf[st.session_state.haulingdf["Standard Matl Name"] == item_selected]
        else:
            st.session_state.haulingdf = st.session_state.haulingdf

        nantiberguna = ["No Item","NS","PL/SR", "Standard Matl Name", "Qty Total","Qty Outstanding","UoM"]
        hauling_edited = st.dataframe(st.session_state.haulingdf, use_container_width=True, selection_mode="multi-row", on_select="rerun")

        selected_forhauling = st.session_state.haulingdf.loc[hauling_edited.selection["rows"]]
        selected_forhauling = st.data_editor(selected_forhauling, use_container_width=True, disabled = nantiberguna)

        if st.button("Save Hauling"):
            hauling_edited.to_csv("D://Dashboard/Pertamina/Hauling.csv", index = False)
            mirexist["Qty Hauling"] = hauling_edited["Qty Hauling"]
            mirexist["Qty Outstanding"] = mirexist["Qty Total"] - mirexist["Qty Hauling"]
            mirexist.to_csv("D:/Dashboard/Pertamina/MIR TEMPLATE - Copy.csv", index=False)
            hasbeensaved("HAULING")

# End of Line
