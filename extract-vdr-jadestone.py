import pandas as pd
import os
import shutil
from openpyxl import load_workbook
from datetime import date, timedelta

vdrDirectory = 'VDR'


for vdrFile in os.listdir(vdrDirectory):
    if vdrFile.endswith('.xlsx'):
        date_date = date.today()
        yesterday = date_date - timedelta(days=1)
        todayDate = yesterday.strftime("%d/%m/%Y")
        fileTemplate = 'Template.xlsx'
        vdrFileDir = os.path.join(vdrDirectory,vdrFile)
        wb = load_workbook(vdrFileDir, data_only=True)
        sheetVDR = wb['VDR']
        vesselName = sheetVDR['E12'].value
        fileName = vesselName + '.xlsx'
        #DUPLICATE TEMPLATE
        template_path = os.path.dirname(fileTemplate)
        file_path = os.path.join(template_path,fileName)
        shutil.copy(fileTemplate,file_path)
        sheetCrew = wb['CORE-CREW']
        shipMMSI = 533180042
        shipNationality = 'MALAYSIA'
        shipType = 'UV'
        captainName = sheetVDR['J12'].value
        location = sheetVDR['L12'].value
        engineHour = sheetVDR['J23:J44']
        engineFuel = sheetVDR['T23:T45']
        activity = sheetVDR['I76':'I87']
        weather_data = []
        activityLog_data = []
        crew_data = []
        rob_data = []

        #GET DATA
        engineHour_value = [cell[0].value for cell in engineHour]
        engineFuel_value = [cell[0].value for cell in engineFuel]
        activity_value = [cell[0].value for cell in activity]
        for row in sheetVDR.iter_rows(min_row=14,max_row=103,min_col=23,max_col=35):
            activityLog_data.append([cell.value for cell in row])
        for row in sheetVDR.iter_rows(min_row=16,max_row=19,min_col=9,max_col=15):
            weather_data.append([cell.value for cell in row])
        for row in sheetCrew.iter_rows(min_row=6,max_row=30,min_col=2,max_col=9):
            crew_data.append([cell.value for cell in row])
        for row in sheetVDR.iter_rows(min_row=52,max_row=72,min_col=2,max_col=21):
            rob_data.append([cell.value for cell in row])

        #COLUMN NAME
        engine_hour_column = ['PORT ME (hrs)','PORT ME OUTLET FLOWMETER (hrs)','CENTER ME (P) (hrs)','CENTER ME (P) OUTLET FLOWMETER (hrs)','CENTER ME (STBD) (hrs)','CENTER ME (STBD) OUTLET FLOWMETER (hrs)','STBD ME (hrs)','STBD ME OUTLET FLOWMETER (hrs)','BOW THRUSTER 1 (hrs)','BOW THRUSTER 2 (hrs)','BOW THRUSTER 3 (hrs)','STERN THRUSTER 1 (hrs)','STERN THRUSTER 2 (hrs)','SHAFT GENERATOR 1 (hrs)','SHAFT GENERATOR 2 (hrs)','GENERATOR 1 (hrs)','GENERATOR 2 (hrs)','GENERATOR 3 (hrs)','GENERATOR 4 (hrs)','GENERATOR 5 (hrs)','GENERATOR 6 (hrs)','EMERGENCY GENERATOR (hrs)']
        engine_fuel_column = ['PORT ME (ltrs)','PORT ME OUTLET FLOWMETER (ltrs)','CENTER ME (P) (ltrs)','CENTER ME (P) OUTLET FLOWMETER (ltrs)','CENTER ME (STBD) (ltrs)','CENTER ME (STBD) OUTLET FLOWMETER (ltrs)','STBD ME (ltrs)','STBD ME OUTLET FLOWMETER (ltrs)','BOW THRUSTER 1 (ltrs)','BOW THRUSTER 2 (ltrs)','BOW THRUSTER 3 (ltrs)','STERN THRUSTER 1 (ltrs)','STERN THRUSTER 2 (ltrs)','SHAFT GENERATOR 1 (ltrs)','SHAFT GENERATOR 2 (ltrs)','GENERATOR 1 (ltrs)','GENERATOR 2 (ltrs)','GENERATOR 3 (ltrs)','GENERATOR 4 (ltrs)','GENERATOR 5 (ltrs)','GENERATOR 6 (ltrs)','EMERGENCY GENERATOR (ltrs)','fuel']
        activity_column = ['Maneuvering at Supply Base','Alongside berth at Supply Base','Anchorage at Supply Base','','En-route: full speed','En-route: economical speed','','Inter-rig/ Maneuvering offshore','Standby steaming offshore','Cargo works within 500m zone','Towing/ Static Towing/ Rigmove/ Hose Handling','Mooring to buoy/platform offshore']
        weather_column = ['TIME','WIND','SWELL','SEA','SKY','VISIBILITY','TEMP ( °C )']
        activityLog_column = ['FROM (0000)','TO (2400)','DURATION','ACTIVITY & LOCATION (Consecutive entries - no gaps allowed)','','','','','VESSEL MOVEMENT CATEGORY','','LOCATION','DP MODE (Y/N)','NON-CREW POB']
        crew_column = ['No.','Name','Rank','Age','Nationality','IC or Passport No.','Date Sign-on(DD/MM/YYYY)','No. of working days (max. 90 days)']
        rob_column = ['ROB','OPENING @ 0000H','','','','','','Consumed','','','','','','','','','','','CLOSING @ 2400H','Remarks']

        #CREATE DATAFRAME
        df = pd.DataFrame({'MMSI': shipMMSI,
                           'Vessel Name': [vesselName],
                           'Nationality of Ship': shipNationality,
                           'Type of Vessel': shipType,
                           'Captain on Duty': [captainName],
                           'Location @ 2400H': [location],
                           'date': todayDate})
        df_engine_hour = pd.DataFrame([engineHour_value], columns=engine_hour_column)
        df_engine_fuel = pd.DataFrame([engineFuel_value], columns=engine_fuel_column)
        df_activity= pd.DataFrame([activity_value],columns=activity_column)
        df_activityLog = pd.DataFrame(activityLog_data,columns=activityLog_column)
        df_weather = pd.DataFrame(weather_data,columns=weather_column)
        df_crew = pd.DataFrame(crew_data,columns=crew_column)
        df_rob = pd.DataFrame(rob_data,columns=rob_column)

        df_summary = [df,df_engine_hour,df_engine_fuel]
        combined_1 = pd.concat(df_summary, axis=1)

        #RESET INDEX
        concat_both = combined_1.reset_index(drop=True)
        df_weather = df_weather.reset_index(drop=True)
        df_activity = df_activity.reset_index(drop=True)
        df_activityLog = df_activityLog.reset_index(drop=True)
        df_crew = df_crew.reset_index(drop=True)
        df_rob = df_rob.reset_index(drop=True)

        #EXPORT TO_EXCEL
        with pd.ExcelWriter(fileName, engine = 'openpyxl', mode='a') as writer:
            concat_both.to_excel(writer, sheet_name='Summary', index=False)
            df_weather.to_excel(writer, sheet_name='Weather', index=False)
            df_activity.to_excel(writer, sheet_name='Activity', index=False)
            df_activityLog.to_excel(writer, sheet_name='Activity log', index=False)
            df_crew.to_excel(writer, sheet_name='Crew list', index=False)
            df_rob.to_excel(writer, sheet_name='ROB', index=False)

#combined_1.to_excel(r'AISHAH AIMS 3.xlsx', sheet_name='Summary', index=False)
#weather_table.to_excel(r'AISHAH AIMS 3.xlsx', sheet_name='Weather', index=False)
memory_usage_bytes = df_activityLog.memory_usage(deep=True).sum()
print(memory_usage_bytes)
