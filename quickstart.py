from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime
import time

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'mykeys.json'

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# performer_count = [1 for i in range(8)]

Input_ID = '1l3t6XPk-mC7-z2HejG3eD2f_BpX6luCdb_TkHxYWKfM'
Output_ID = '1k0YuJUFq52xtb19ZBrnn411rBQEG0XKrXSeUXw62uE4'
Webflow_ID = '19xtDhqs3H4_3iBGgfKwllwW2AxusEpu1F67cIWY37i4'
Zoom_links_ID = '1MV-4ZjOjoa--HYHSvZrOraP4lvmWyxmjPaM4v0TeacM'


def getinfo(lis):

    #0.performer email
    info = [lis[1]]

    #1.name of performer
    first_name, last_name_letter = lis[2], lis[3][0]
    info.append(first_name + ' ' + last_name_letter + '.')

    #2.performance type
    performance_type = lis[9]
    if performance_type.lower() == 'instrument':
        performance_type = lis[10]
    if lis[10].lower() == 'other': performance_type = 'Instrument/Other'
    info.append(performance_type)

    #3.performance piece
    title = lis[11]
    #Opus info
    if 'op' not in title.lower():
        if lis[13] != '':
            if 'op' not in lis[13].lower(): title += ' Op.' + lis[13]
            else: title += ' ' + lis[13]
    #Movement info
    if 'no.' in title.lower() or 'mov' in title.lower() or 'mvt' in title.lower():
        title += ''
    else:
        if lis[14] != '':
            if 'no' in lis[14].lower() or 'mov' in lis[14].lower() \
                    or 'mvt' in lis[14].lower(): title += ' ' + lis[14]
            else: title += ' Mvt.' + lis[14]
    #Key info
    if lis[12] != '': title += ' in ' + lis[12]
    info.append(title)

    #4.Composer Name
    info.append(lis[15])

    #session number
    ind = lis[8].index('Session')
    ind_res = int(lis[8][ind+8])

    #############################################
    #not printed information: for comparison ONLY
    #5.performer age
    info.append(int(lis[6]))

    #6.submission date/time
    info.append(lis[0])
    #############################################

    print(info)

    return info, ind_res


def zoom_info(lis):
    #session number
    lis[1], lis[0] = lis[0], lis[1]
    lis[0] = 'Session ' + lis[0]

    # date, time
    date = lis.pop(1)
    lis[1] += ', ' + date

    # remove zoom meeting title
    lis.pop(-2)

    #add extra words
    lis[2] = 'Meeting ID: ' + lis[2]
    lis[3] = 'Password: ' + lis[3]

    return lis


def sort(lis):
    #used bubble sort for simplicity
    for i in range(len(lis)-1):
        for j in range(i+1, len(lis)):
            #compare age
            if lis[i][5] > lis[j][5]:
                lis[i], lis[j] = lis[j], lis[i]
            #if same age: compare submission time
            elif lis[i][5] == lis[j][5]:
                date_time1, date_time2 = lis[i][6], lis[j][6]
                timestamp1 = time.mktime(datetime.strptime(date_time1, "%Y-%m-%d %H:%M:%S").timetuple())
                timestamp2 = time.mktime(datetime.strptime(date_time2, "%Y-%m-%d %H:%M:%S").timetuple())
                if timestamp1 > timestamp2: lis[i], lis[j] = lis[j], lis[i]


def main():

    # Call the Sheets API and get input
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    #get input from web-flow sheet
    raw_input1 = sheet.values().get(spreadsheetId=Webflow_ID,
                                    range="Sheet1!A3:W12").execute()
    performer_info = raw_input1.get('values', [])

    #get input from zoom link sheet
    raw_input2 = sheet.values().get(spreadsheetId=Zoom_links_ID,
                                    range="May 2021!A2:G9").execute()
    zoom_links = raw_input2.get('values', [])

    """0.Submission Time	1.Email Address   2.Performer's First Name  3.Performer's Last Name	
    4.Parent's First Name   5.Parent's Last Name	6.Performer's Age	7.Describe your experience	
    8.When would you like to perform?	9.What will your performance consist of?	10.What instrument?	
    11.Title of piece	12.Key	13.Opus	14.Movement	15.Composer Full Name   16.Length of Performance
    17.Number of Performers	    18.Day 1 changing piece acknowledgement	    19.Performance Method
    20.Answer host questions?	21.How did you hear about us?   23.Receive email updates?
    
    required info: 1, 2, 3, 9, 10, 11, 15
    preferred info: 8, 12, 13, 14, 19, 20
    """

    #initialize
    title = ['Email', 'Performer First Name', 'Performance Type', 'Performance Piece Name', 'Composer']
    master_list = {i+1: [] for i in range(8)}

    #organize zoom links
    for i in range(len(zoom_links)):
        zoom_info(zoom_links[i])


    #gather information for each performer & append to master list -> dictionary
    for i in range(len(performer_info)):
        performer, ind = getinfo(performer_info[i])
        master_list[ind].append(performer)

    #sort performers within each session
    for i in range(1, 9):
        sort(master_list[i])

    #delete extra info, add zoom links and titles
    for i in range(1, 9):
        for j in range(len(master_list[i])):
            temp = master_list[i][j]
            del temp[5:]
        master_list[i].insert(0, zoom_links[i-1])
        master_list[i].insert(1, title)

    #delete extra info(not needed currently)
    # service.spreadsheets().values().batchClear(spreadsheetId=Output_ID,
    #                                            body={'ranges': "Sheet1!A1:F8"}).execute()

    #output to file
    for i in range(1, 9):
        service.spreadsheets().values().update(spreadsheetId=Output_ID, range="Session{}!A1".format(i),
                                               valueInputOption="USER_ENTERED",
                                               body={"values": master_list[i]}).execute()
    print("done")

    #end of program


if __name__ == '__main__':
    main()

#(in the future) automatically create zoom meetings
