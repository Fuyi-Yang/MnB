from googleapiclient.discovery import build
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'mykeys.json'

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# performer_count = [1 for i in range(8)]

Input_ID = '1l3t6XPk-mC7-z2HejG3eD2f_BpX6luCdb_TkHxYWKfM'
Output_ID = '1k0YuJUFq52xtb19ZBrnn411rBQEG0XKrXSeUXw62uE4'
Webflow_ID = '19xtDhqs3H4_3iBGgfKwllwW2AxusEpu1F67cIWY37i4'
Zoom_links_ID = '1MV-4ZjOjoa--HYHSvZrOraP4lvmWyxmjPaM4v0TeacM'


def getinfo(lis, order):

    #order of performer
    info = [str(order)]

    #name of performer
    first_name, last_name_letter = lis[2], lis[3][0]
    info.append(first_name + ' ' + last_name_letter + '.')
    info.append(lis[10])

    # performance piece
    title = lis[11]
    if lis[13] != '': title += ' ' + lis[13]
    if lis[14] != '': title += ' ' + lis[14]
    if lis[12] != '': title += ' in ' + lis[12]
    info.append(title)

    #Composer Name
    info.append(lis[15])

    return info


def zoominfo(lis):
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


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API and get input
    sheet = service.spreadsheets()
    raw_input1 = sheet.values().get(spreadsheetId=Webflow_ID,
                                    range="Sheet1!A2:X4").execute()
    performer_info = raw_input1.get('values', [])

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
    order = 1
    title_switch = True
    res = []

    if title_switch:
        res.append(['Order', 'Performer First Name', 'Performance Type', 'Performance Piece Name', 'Composer'])\


    # gather information for each performer & append to output 2d list
    for i in range(len(performer_info)):
        res.append(getinfo(performer_info[i], order))
        order += 1

    #organize zoom links
    for i in range(len(zoom_links)):
        zoominfo(zoom_links[i])

    # res.insert(1, [])
    # print(res)

    #delete extra info
    # service.spreadsheets().values().batchClear(spreadsheetId=Output_ID,
    #                                            body={'ranges': "Sheet1!A1:F8"}).execute()

    #output to file
    # service.spreadsheets().values().update(spreadsheetId=Output_ID, range="Sheet1!A1",
    #                                        valueInputOption="USER_ENTERED", body={"values": zoom_links}).execute()



    """todo:
    sort people by session
    organize all into
    """

if __name__ == '__main__':
    main()
