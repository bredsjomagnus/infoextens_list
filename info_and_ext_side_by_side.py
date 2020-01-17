from __future__ import print_function
import pandas as pd
import datetime
import os
import sys
import getopt
import pickle
import os.path
import pprint
from termcolor import colored, cprint
from env import *
from sheet_config import *
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Config PrettyPrinter
pp = pprint.PrettyPrinter(indent=2)

# SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

time_stamp = str(datetime.datetime.now())[:10] # 2019-09-04
extenslist_file_name = time_stamp + '_extenslista.csv'
infomentorlist_file_name = time_stamp + '_elevlista.csv'

def authenticate():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('sheet_token.pickle'):
        with open('sheet_token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'sheet_credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('sheet_token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    return service

def get_extens(service, EXTENS_ID):
    """
    Get Extens file in Drive and return it as a Dataframe
    """
    RANGE = 'extens!A1:D'
    FILENAME = extenslist_file_name

     # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=EXTENS_ID,
                                range=RANGE).execute()
    values = result.get('values', [])

    if not values:
        print('No data found in extens.')
    else:
        # print('Läser in extensfil från Driven till Dataframe:')
        klasser = []
        personnummers = []
        names = []
        
        # emails = []
        # print(values)
        for i, row in enumerate(values):
            if i > 0:
                # Print columns A and E, which correspond to indices 0 and 4.
                try: # this will take care of eventually empty cells.
                    # name = "\""+row[0]+"\""
                    klass = row[0]
                    personnummer = str(row[1])
                    lastname = row[2]
                    firstname = row[3]
                    name = lastname + ", " + firstname
                    # print(personnummer)

                    klasser.append(klass)
                    personnummers.append(personnummer)
                    names.append(name)
                    # emails.append(email)
                except Exception as e:
                    print()
                    print("While reading rows from file error ->", e)
                    print("Null at ", row[1])
                    print()
        # print("Skapar DataFrame och sparar som %s" % (FILENAME))
        elevlista_dict = {
            'Klass': klasser,
            'Namn': names,
            'Personnummer': personnummers,
        }
        df_extens = pd.DataFrame.from_dict(elevlista_dict)
        # df_extens.to_csv(FILENAME, sep=",", index=False)

        return df_extens

def get_infomentor(service, INFOMENTOR_ID):
    """
    Get Infometor file in Drive and return it as a Dataframe
    """
    RANGE = 'elevlista!A1:D'
    FILENAME = infomentorlist_file_name

     # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=INFOMENTOR_ID,
                                range=RANGE).execute()
    values = result.get('values', [])

    if not values:
        print('No data found in extens.')
    else:
        # print('Läser in infomentorfil från Driven till Dataframe:')
        klasser = []
        personnummers = []
        names = []
        
        # emails = []
        # print(values)
        for i, row in enumerate(values):
            if i > 0:
                # Print columns A and E, which correspond to indices 0 and 4.
                try: # this will take care of eventually empty cells.
                    # name = "\""+row[0]+"\""
                    klass = row[0]
                    personnummer = str(row[3])
                    name = row[1]

                    if len(personnummer) < 10:
                        personnummer = "0" + personnummer
                    personnummer = personnummer.strip()
                    personnummer = personnummer[:6] + '-' + personnummer[-4:]

                    personnummer = personnummer.strip()
                    
                    # print(personnummer)
                    # if len(personnummer) < 11:
                    #     print("personnummer_pre: %s, personnummer_first: %s, personnummer: %s" % (personnummer_pre, personnummer_first, personnummer))
                    #     print("personnummer_first[:6]: %s, personnummer_first[-4:]: %s" % (personnummer_first[:6],personnummer_first[-4:]))
                    klasser.append(klass)
                    personnummers.append(personnummer)
                    names.append(name)
                    # emails.append(email)
                except Exception as e:
                    print()
                    print("While reading rows from file error ->", e)
                    print("Null at ", row[1])
                    print()
        # print("Skapar DataFrame och sparar som %s" % (FILENAME))
        infomentor_dict = {
            'Klass': klasser,
            'Namn': names,
            'Personnummer': personnummers,
        }
        df_infomentor = pd.DataFrame.from_dict(infomentor_dict)
        # df_infomentor.to_csv(FILENAME, sep=",", index=False)

    return df_infomentor

def create_spreadsheet(service):
    """
    Create new spreadsheet
    """
    # print("Beginning process...")

    time_stamp = str(datetime.datetime.now())[:10] # 2019-09-04
    spreadsheet_name = time_stamp + '_' + SPREADSHEET_TITLE
    SPREADSHEET_ID = ""
    spreadsheet_body = {
        "properties": {
            "title": spreadsheet_name
        }
    }
    try:
        request = service.spreadsheets().create(body=spreadsheet_body)
        spreadsheet = request.execute()
        SPREADSHEET_ID = spreadsheet['spreadsheetId']
        # print("spreadsheet id: ", SPREADSHEET_ID)
    except Exception as e:
        print("While trying to create new spreadsheet error: ", e)
        sys.exit()
    
    return SPREADSHEET_ID

def update_spreadsheet(service, SPREADSHEET_ID, body, message="No message"):
    """
    Update batch of requests in body
    """

    response = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()

    # print(message)

    return response

def create_sheets(service, SPREADSHEET_ID, sheet_objects):
    requests = []
    # print("Creating sheets...")
    # CREATE SHEETS
    for sheet in sheet_objects:
        requests.append(sheet_objects[sheet])

    # Trying to update spreadsheet with assigned requests
    try:
        body = {
            "requests": requests
        }
        response = update_spreadsheet(service, SPREADSHEET_ID, body, "Sheets created")
    except Exception as e:
        print("While trying to batchUpdate error: ", e)
        sys.exit()

    return response

def get_sheet_ids(service, SPREADSHEET_ID):
    """
    Getting the proporites of the created sheets.
    """
    # print("Getting sheet proporties...")
    sheet_dict = {}
    # Trying to get sheetIds
    try:
        fields="sheets.properties"
        request = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields=fields)
        response = request.execute()

        # How to iterate the response and get title of a sheet
        for prop in response['sheets']:
            # print(prop['properties']['title'])
            sheet_dict[prop['properties']['title']] = prop['properties']['sheetId']
    except Exception as e:
        print("While trying spreadsheets().get() error: ", e)
        sys.exit()
    # pp.pprint(sheet_dict)
    return sheet_dict

def customize_columns(service, SPREADSHEET_ID, columns):
    """
    Set column width in sheets. columns is generated in sheet_config.py
    based on template and sheetId
    """
    # print("Setting column and row widths in the sheets")
    requests = []
    for key in columns:
        requests.append(columns[key])
    
    # Trying to update spreadsheet with assigned requests
    try:
        body = {
            "requests": requests
        }
        response = update_spreadsheet(service, SPREADSHEET_ID, body, "Columns set")
        return response
    except Exception as e:
        print("While trying to batchUpdate error: ", e)
        sys.exit()
    
def check_name(klass, info_name, ext_firstname, ext_lastname):
    ext_name = ext_lastname +", " + ext_firstname
    if info_name != ext_name:
        # print("Klass %s;   INFOMENTOR NAMN; %s --> EXTENS NAMN: %s" % (klass, info_name, ext_name))
        pass

def find_corresponding_name(name, personnummer, df):

    res = {
        "namn": "NOT FOUND",
        "personnummer": "NOT FOUND",
        "matching": 'NOT FOUND'
    }
    try:
        student_personnummer_series = df[df['Personnummer'].str.contains(personnummer)]
        res = {
            "namn": str(student_personnummer_series['Namn'].values[0]),
            "personnummer": str(student_personnummer_series['Personnummer'].values[0]),
            "matching": "Personnummer"
        }
    except:
        try:
            student_name_series = df[df['Namn'].str.contains(name)]
            res = {
                "namn": str(student_name_series['Namn'].values[0]),
                "personnummer": str(student_name_series['Personnummer'].values[0]),
                "matching": "Namn"
            }
        except:
            pass
    
    return res

def is_lists_equal(row):
    equal = True
    missing = False
    reason = []
    if row[1] == "NOT FOUND":
        reason.append("Saknas i Infomentor")
        missing = True
        equal = False
    elif row[4] == "NOT FOUND":
        reason.append("Saknas i Extens")
        missing = True
        equal = False
    
    if not missing:
        if row[1].lower() != row[4].lower():
            equal = False
            reason.append("Namnet")
        if row[2].lower() != row[5].lower():
            equal = False
            reason.append("Personnummret")

    if len(reason) > 1:
        print(reason)
    return equal, reason


def add_content(service, SPREADSHEET_ID, df_infomentor, df_extens):
    sheet_name = sheet_names[0]

    # ADDING HEADER
    content = []
    reported_content = []
    content = header
    reported_content = reported_header
    # print("content:",content)
    sheet_range = sheet_name + "!A1"
    reported_range = "Avvikelser!A1"
    try:
        range = sheet_range
        values = content
        resource = {
            "values": values
        }
        # use append to add rows and update to overwrite
        response = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=range, body=resource, valueInputOption="USER_ENTERED").execute()

        range = reported_range
        values = reported_content
        resource = {
            "values": values
        }
        # use append to add rows and update to overwrite
        response = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=range, body=resource, valueInputOption="USER_ENTERED").execute()
    except Exception as e:
        print("While trying to append values error: ", e)
        sys.exit()

    # CONTENT
    content = []
    reported_content = []


    for klass in klasser:
        """
        Tar ut klass för klass. 
        Kollar om de är lika långa. 
        Utgår från den lista som eventuellt är längst, annars Infomentors. 

        Letar efter korresponderande elev i andra listan och placerar dessa
        båda på samma rad i nya spreadsheetet.
        """
        info_klass_df = df_infomentor[df_infomentor['Klass'].str.contains(klass)]   # alla elever ur klassen 'klass' i elevlistan
        ext_klass_df = df_extens[df_extens['Klass'].str.contains(klass)]            # alla elever ur klassen 'klass' i extenslistan

        info_klass = info_klass_df.loc[:,'Klass'].tolist()                  # klasstillhörighet i en lista; ["7A", "7A", "7A",...]
        info_namn = info_klass_df.loc[:,'Namn'].tolist()                    # elevernas namn i en lista
        info_personnummer = info_klass_df.loc[:,'Personnummer'].tolist()    # elevernas personnummer i en lista

        ext_klass = ext_klass_df.loc[:,'Klass'].tolist()                    # klasstillhörighet i en lista; ["7A", "7A", "7A",...]
        ext_namn = ext_klass_df.loc[:,'Namn'].tolist()                      # elevernas namn i en lista
        ext_personnummer = ext_klass_df.loc[:,'Personnummer'].tolist()      # elevernas personnummer i en lista

        ext_longer = len(info_klass_df.index) < len(ext_klass_df)           # is class in extenslista longer than in elevlista?

        not_found_in_other = 0
        not_found_students = []
        found_in_other = 0  # counter to check so that no students in smaller list is overlooked
        found_in_class = ""

        class_reasons = []

        row = []
        if ext_longer:
            for i, klass in enumerate(ext_klass):
                student = {}
                student = find_corresponding_name(ext_namn[i], ext_personnummer[i], info_klass_df)

                row = [ext_klass[i], student['namn'], student['personnummer'], ext_klass[i], ext_namn[i], ext_personnummer[i]]
                
                equal, reasons = is_lists_equal(row)
                if not equal:
                    reason_string =', '.join(map(str, reasons))
                    row.append(reason_string)
                    reported_content.append(row)

                content.append(row)

                # checking if all students in infolist for this class is found.
                if student['matching'] == "NOT FOUND":
                    not_found_in_other += 1
                    found_in_class = ext_klass[i]
                    not_found_students.append(ext_namn[i])
                else:
                    found_in_other += 1
                    found_in_class = ext_klass[i]
                
                if len(reasons) > 0:
                    class_reasons.append([ext_klass[i], ext_namn[i], ext_personnummer[i], reason_string])
            
            if found_in_other == len(info_klass_df.index):
                cprint("Class %s    Compare agains Extlist: %d    Matched students in Infomentorlist: %s, Number of students in Infomentorlist: %d" % (found_in_class, len(ext_klass_df.index), found_in_other, len(info_klass_df.index)), 'green')
            elif len(info_klass_df.index) > found_in_other:
                cprint("Class %s    Matched students in Infomentorlist: %s, Number of students in Infomentorlist: %d STUDENT IN INFOMENTORLIST OVERLOOKED" % (found_in_class, found_in_other, len(info_klass_df.index)), 'red')
            
            if len(class_reasons) > 0:
                cprint("Class %s    Students that differ in Infomentorlist compared to Extenslist: " % (found_in_class), 'yellow')
                for cr in class_reasons:
                    cprint("\t\t\t\t\t\t\t\t\t\t%s; %s; '%s'" %
                           (cr[1], cr[2], cr[3].upper()), 'cyan')
                print()

        else:
            for i, klass in enumerate(info_klass):
                student = {}
                student = find_corresponding_name(info_namn[i], info_personnummer[i], ext_klass_df)
                
                row = [info_klass[i], info_namn[i], info_personnummer[i], info_klass[i], student['namn'], student['personnummer']]

                equal, reasons = is_lists_equal(row)
                if not equal:
                    reason_string =', '.join(map(str, reasons))
                    row.append(reason_string)
                    reported_content.append(row)

                content.append(row)

                # checking if all students in infolist for this class is found.
                if student['matching'] == "NOT FOUND":
                    not_found_in_other += 1
                    found_in_class = info_klass[i]
                    not_found_students.append(info_namn[i])
                else: # if a match exists add to counter for matched students
                    found_in_other += 1
                    found_in_class = info_klass[i]
                
                if len(reasons) > 0: # if deviation reason exist append this to class_reasons list
                    class_reasons.append([info_klass[i], info_namn[i], info_personnummer[i], reason_string])
                
            
            # after for loop check if found matches corresponds to list length
            if found_in_other == len(ext_klass_df.index):
                cprint("Class %s    Compare agains Infolist: %d   Matched students in Extenslist: %s,     Number of students in Extenslist: %d" % (found_in_class, len(info_klass_df.index), found_in_other, len(ext_klass_df.index)), 'green')
            elif len(ext_klass_df.index) > found_in_other:
                cprint("Class %s    Matched students in Extenslist: %s,     Number of students in Extenslist: %d STUDENT IN EXTENSLIST OVERLOOKED" % (found_in_class, found_in_other, len(ext_klass_df.index)), 'red')
            
            # after for loop display eventual devaition reasons
            if len(class_reasons) > 0:
                cprint("Class %s    Students that differ in Extenslist compared to Infomentorlist: " % (found_in_class), 'yellow')
                for cr in class_reasons:
                    cprint("\t\t\t\t\t\t\t\t\t\t%s; %s; '%s'" %
                           (cr[1], cr[2], cr[3].upper()), 'cyan')
                print()
            
        empty_row = ["", "", "", "", "", ""]
        content.append(empty_row)
        # reported_content.append(empty_row)
        
    print()
    print("FOUND %s DEVIATIONS!" % (len(reported_content)), end=" ")
    sheet_range = sheet_name + "!A2"
    reported_range = "Avvikelser!A2"
    try:
        range = sheet_range
        values = content
        resource = {
            "values": values
        }
        # use append to add rows and update to overwrite
        response = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=range, body=resource, valueInputOption="USER_ENTERED").execute()

        range = reported_range
        values = reported_content
        resource = {
            "values": values
        }
        # use append to add rows and update to overwrite
        response = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=range, body=resource, valueInputOption="USER_ENTERED").execute()
    except Exception as e:
        print("While trying to append values error: ", e)
    
    print("REPORT CREATED!")

service = authenticate()
df_extens = get_extens(service, EXTENS_ID)
df_infomentor = get_infomentor(service, INFOMENTOR_ID)

SPREADSHEET_ID = create_spreadsheet(service)                    # Skapar spreadsheetet:service.spreadsheets().create
create_sheets(service, SPREADSHEET_ID, sheet_objects)           # Skapar sheets:service.spreadsheets().batchUpdate
sheet_dict = get_sheet_ids(service, SPREADSHEET_ID)             # sheet_dict:{"7A": sheet id nummer,...}
columns_object = generate_columns_update_object(sheet_dict)            # Skapar requestobjekt för justering av kolumnvidder
customize_columns(service, SPREADSHEET_ID, columns_object)             # Ändrar kolumnvidd i sheets:service.spreadsheets().batchUpdate
print()
add_content(service, SPREADSHEET_ID, df_infomentor, df_extens)                # Lägger till innehåll i sheets:service.spreadsheets().values().update
