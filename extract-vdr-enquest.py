import pandas as pd
import os
import shutil
from openpyxl import load_workbook
from datetime import date, timedelta

#VDR DIRECTORY
vdrClientDir = 'C:/Users/DOUBLE33/PycharmProjects/extract-vdr/vdr/enquest'

for vdrFile in os.listdir(vdrClientDir):
    if vdrFile.endswith('.xlsx'):
        date_date = date.today()
        yesterday = date_date - timedelta(days=2)
        day_yesterday = yesterday.day
        todayDate = yesterday.strftime("%d/%m/%Y")
        fileTemplate = 'Template.xlsx'
        vdrFileDir = os.path.join(vdrClientDir,vdrFile)
        wb = load_workbook(vdrFileDir, data_only=True)
        sheetDailyReport = wb['Daily Report']
        vesselName = sheetDailyReport['C4'].value
        fileName = vesselName + '.xlsx'
        #DUPLICATE TEMPLATE
        template_path = os.path.dirname(fileTemplate)
        file_path = os.path.join(template_path,fileName)
        shutil.copy(fileTemplate,file_path)

        sheetActivities = wb['Boat Movements']
        sheetFuel = wb['Fuel Monitoring']
        sheetTOD = wb['TOD Report']
        location = sheetDailyReport['U4'].value
        weather_data_1 = []
        weather_data_2 = []
        rob_data = []
        fuel_data = []
        activities_data = []
        TOD_data = []
        cargo_descriptions = ['FUEL OIL (Ltrs)', 'FRESH WATER (Ltrs)', 'LUB OIL (Ltrs)', 'DISPERSANT (Ltrs)']

        #GET DATA
        for row in sheetActivities.iter_rows(min_row=5,max_row=35,min_col=6,max_col=17):
            activities_data.append([cell.value for cell in row])
        for row in sheetFuel.iter_rows(min_row=8,max_row=38,min_col=4,max_col=22):
            fuel_data.append([cell.value for cell in row])
        for row in sheetDailyReport.iter_rows(min_row=7,max_row=9,min_col=3,max_col=4):
            weather_data_1.append([cell.value for cell in row])
        for row in sheetDailyReport.iter_rows(min_row=7,max_row=9,min_col=14,max_col=15):
            weather_data_2.append([cell.value for cell in row])
        for row in sheetDailyReport.iter_rows(min_row=13,max_row=16,min_col=14,max_col=28):
            rob_data.append([cell.value for cell in row])
        for row in sheetTOD.iter_rows(min_row=18,max_row=25,min_col=3,max_col=19):
            TOD_data.append([cell.value for cell in row])

        #COLUMN NAME
        summary_column = ['MMSI','Vessel Name','Nationality of Ship','Type of Vessel','Location','Date']
        weather_column = ['column-1', 'column-2']
        fuel_column = ['Port ME (HRS)','Port ME (LTRS)','Centre ME (HRS)','Centre ME (LTRS)','Stbd ME (HRS)','Stbd ME (LTRS)','Genset 1 (HRS)','Genset 1 (LTRS)','Genset 2 (HRS)','Genset 2 (LTRS)','Genset 3 (HRS)','Genset 3 (LTRS)','Genset 4 (HRS)','Genset 4 (LTRS)','Others (HRS)','Others (LTRS)','Main Eng. Cons. (LTRS)','Aux Eng. Cons. (LTRS)','Total Daily Cons. (LTRS)']
        activities_column = ['Anchorage','In Port - Shifting','Alongside Jetty','Enroute Econ Speed (85% MCR)','Enroute Econ Speed (100% MCR)','Inter Rig','Cargo Work / Passenger Transfer','Standby Close - Active in/outside 500m Zone','Standby Normal - Nil Activity','Towing @ Barge / Rig Move','Tied-up to Mooring Standby Buoy','Anchor Handling']
        TOD_column = ['No.','Name','','','','','Vessel Joining Date (DD-MM-YY)','','Length of Stay Onboard (days)','','Rank','','Nationality','','International Passport Number / NRIC','Valid Thru (Medical)','BOSIET Validity']
        rob_column = ['Open','','','Consumption','','','Loaded','','','','Discharged','','','','ROB']

        #COMBINE VERTICALLY
        combined_weather_data = weather_data_1 + weather_data_2

        #CREATE DATAFRAME
        df = pd.DataFrame({'Vessel Name': [vesselName],
                           'Location': [location],
                           'date': todayDate})
        df_weather_combine = pd.DataFrame(combined_weather_data,columns=weather_column)
        df_TOD = pd.DataFrame(TOD_data,columns=TOD_column)
        df_cargo_des = pd.DataFrame({'CARGO DESCRIPTION': cargo_descriptions})
        df_rob = pd.DataFrame(rob_data,columns=rob_column)
        df_activities = pd.DataFrame(activities_data,columns=activities_column)
        df_fuel = pd.DataFrame(fuel_data,columns=fuel_column)
        df_rob_combine = pd.concat((df_cargo_des,df_rob),axis=1)
        df_activities_fuel_combine = pd.concat((df_fuel,df_activities),axis=1)
        selected_row_transposed = df_activities_fuel_combine.iloc[day_yesterday - 1:day_yesterday].T

        #EXPORT TO EXCEL
        with pd.ExcelWriter(fileName, engine = 'openpyxl', mode='a') as writer:
            df_weather_combine.to_excel(writer, sheet_name='Weather', index=False)
            selected_row_transposed.to_excel(writer, sheet_name='Summary', index=False)
            df_TOD.to_excel(writer, sheet_name='TOD', index=False)
            df_rob_combine.to_excel(writer, sheet_name='ROB', index=False)