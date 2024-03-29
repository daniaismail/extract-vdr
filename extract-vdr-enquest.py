import sys
import imaplib
import email
import sys
import pandas as pd
import os
import shutil
from openpyxl import load_workbook
from datetime import date, timedelta
import logging

#VDR DIRECTORY
vdrClientDir = 'C:/Users/DOUBLE33/PycharmProjects/extract-vdr/vdr/enquest'

#FUNCTIONS
#download vdr from email
def dwl_vdr(email_add, password, server, vesselEmail):
    imap = imaplib.IMAP4_SSL(server, 993)
    imap.login(email_add, password)
    imap.select('INBOX')

    index = 0

    while index < len(vesselEmail):
        try:
            logging.debug(f'Download from email {vesselEmail[index]}')
            typ, data = imap.search(None, '(SINCE %s)' % (dt,),'(FROM %s)' % (vesselEmail[index],))

            print(vesselEmail[index])
            for num in data[0].split():
                typ, data = imap.fetch(num, '(RFC822)')
                raw_email = data[0][1]
                raw_email_string = raw_email.decode('ISO-8859–1')
                email_message = email.message_from_string(raw_email_string)
                subject_name = email_message['subject']
                print(subject_name)

                # att_path = "No attachment found from email " + subject_name
                logging.debug(f'Email subject: {subject_name}')
                for part in email_message.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue

                    if part.get_filename() and any(keyword in part.get_filename() for keyword in ['VDR','VDMR']) and part.get_filename().endswith('.xlsx'):
                        filename = os.path.join(vdrClientDir, part.get_filename())
                        logging.debug(f'file path: {filename}')
                        print(filename)
                        with open(filename, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        logging.debug(f'File downloaded: {filename}')
            index += 1

        except (OSError, TypeError) as k:
            print(k)
            continue

    imap.close()
    imap.logout()
    logging.debug('Job finished..')

def normalize_sheet_name(sheet_name):
    return sheet_name.lower().strip()

#EMAIL INFO
email_vdr = "mvmcc@meridiansurveys.com.my"
pwd_vdr = "dc)in]}Xzk&%"
server_mssb = "meridian-svr.meridiansurveys.com.my"

#READ FROM TXT FILE AND APPEND INTO LIST
vesselemail_list = []

date_date = date.today()
dt = date_date.strftime('%d-%b-%Y')
yesterday = date_date - timedelta(days=1)
day_yesterday = yesterday.day
todayDate = yesterday.strftime("%d/%m/%Y")
fileTemplate = 'Template.xlsx'

logging.debug('Logging started..')
logging.debug('Open enquest-email-list.txt')
try:
    with open(r'C:\Users\DOUBLE33\PycharmProjects\extract-vdr\enquest-email-list.txt') as f:
        try:
            for line in f:
                vesselemail_list.append(line.replace("\n", ""))
        except:
            logging.error('Error info in text file. Check for typo')
except:
    logging.exception('Error info in text file or check txt file name')
    sys.exit()

try:
    logging.debug('Downloading from mvmcc@meridiansurveys.com.my..')
    dwl_vdr(email_vdr, pwd_vdr, server_mssb, vesselemail_list)

except Exception as e:
    logging.error("Error occurred", exc_info=True)

for vdrFile in os.listdir(vdrClientDir):
    if vdrFile.endswith('.xlsx'):
        print(f'Accessing: {vdrFile}')
        vdrFileDir = os.path.join(vdrClientDir, vdrFile)
        wb = load_workbook(vdrFileDir, data_only=True)

        # Normalize sheet names
        normalized_sheet_names = [normalize_sheet_name(sheet_name) for sheet_name in wb.sheetnames]

        sheetDailyReport = None
        sheetActivities = None
        sheetFuel = None
        sheetTOD = None

        # Find the sheets dynamically
        for normalized_name, sheet_name in zip(normalized_sheet_names, wb.sheetnames):
            if normalized_name == 'daily report':
                sheetDailyReport = wb[sheet_name]
            elif normalized_name == 'boat movements':
                sheetActivities = wb[sheet_name]
            elif normalized_name == 'fuel monitoring':
                sheetFuel = wb[sheet_name]
            elif normalized_name == 'tod report':
                sheetTOD = wb[sheet_name]

        if sheetDailyReport is None or sheetActivities is None or sheetFuel is None or sheetTOD is None:
            print(f'Error: One or more required sheets not found in {vdrFile}')
            continue

        vesselName = sheetDailyReport['C4'].value
        fileName = vesselName + '.xlsx'
        # DUPLICATE TEMPLATE
        template_path = os.path.dirname(fileTemplate)
        file_path = os.path.join(template_path, fileName)
        shutil.copy(fileTemplate, file_path)
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
        for row in sheetTOD.iter_rows(min_row=18,max_row=40,min_col=3,max_col=19):
            TOD_data.append([cell.value for cell in row])

        #COLUMN NAME
        summary_column = ['MMSI','Vessel Name','Nationality of Ship','Type of Vessel','Location','Date']
        weather_column = ['column-1', 'column-2']
        fuel_column = ['Port ME (HRS)','Port ME (LTRS)','Centre ME (HRS)','Centre ME (LTRS)','Stbd ME (HRS)','Stbd ME (LTRS)','Genset 1 (HRS)','Genset 1 (LTRS)','Genset 2 (HRS)','Genset 2 (LTRS)','Genset 3 (HRS)','Genset 3 (LTRS)','Bow Thruster (HRS)','Bow Thruster (LTRS)','Others (HRS)','Others (LTRS)','Main Eng. Cons. (LTRS)','Aux Eng. Cons. (LTRS)','Total Daily Cons. (LTRS)']
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
        selected_row = df_activities_fuel_combine.iloc[day_yesterday - 1]
        selected_row_transposed = selected_row.to_frame().T
        #selected_row_transposed = df_activities_fuel_combine.iloc[day_yesterday - 1:day_yesterday].T

        #EXPORT TO EXCEL
        with pd.ExcelWriter(fileName, engine = 'openpyxl', mode='a') as writer:
            df_weather_combine.to_excel(writer, sheet_name='Weather', index=False)
            selected_row_transposed.to_excel(writer, sheet_name='Summary', index=False)
            df_TOD.to_excel(writer, sheet_name='TOD', index=False)
            df_rob_combine.to_excel(writer, sheet_name='ROB', index=False)
            print(f'{vdrFile} is DONE')