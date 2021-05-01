from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime
import time
import string

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'mykeys.json'

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)


Input_ID = '1l3t6XPk-mC7-z2HejG3eD2f_BpX6luCdb_TkHxYWKfM'
Output_ID = '1k0YuJUFq52xtb19ZBrnn411rBQEG0XKrXSeUXw62uE4'
Webflow_ID = '19xtDhqs3H4_3iBGgfKwllwW2AxusEpu1F67cIWY37i4'
Zoom_links_ID = '1MV-4ZjOjoa--HYHSvZrOraP4lvmWyxmjPaM4v0TeacM'


def getinfo(lis):
    # A, B, C, D, J, L, M, N

    # 0.Session Number
    # 1.performer email
    ind = lis[8].index('Session')
    ind_res = int(lis[8][ind + 8])
    info = ['Session ' + lis[8][ind + 8], lis[1]]

    # 2.performer full name
    info.append(lis[2] + ' ' + lis[3])

    # 3.performer first name only
    info.append(lis[2])

    # 4.performer age
    if lis[6] == '18+':
        info.append(20)
    elif len(lis[6]) > 3:
        info.append(30)
    else:
        info.append(int(lis[6]))

    # 5.interview decision
    info.append(lis[20])

    # 6.performance type
    performance_type = lis[9]
    if performance_type.lower() == 'instrument':
        performance_type = lis[10]
    if lis[10].lower() == 'other': performance_type = 'Instrument/Other'
    info.append(performance_type)

    # 7.performance piece
    title = lis[11]
    # Opus info
    arr = ['op', 'bwv', 'k', 'nr']
    if not any(map(title.lower().__contains__, arr)) and lis[13] != '':
        if any(map(lis[13].lower().__contains__, arr)):
            title += ' ' + lis[13]
        else:
            title += ' Op.' + lis[13]

    # Movement info
    arr = ['mov', 'mvt', 'mvmt', 'no.']
    if not any(map(title.lower().__contains__, arr)) and lis[14] != '':
        if any(map(lis[14].lower().__contains__, arr)):
            title += ' ' + lis[14]
        else:
            title += ' Mvt.' + lis[14]

    # Key info
    if not any(map(title.lower().__contains__, ['major', 'minor'])) and lis[12] != '':
        title += ' in ' + lis[12]

    info.append(title)

    # 8.Composer Name
    info.append(lis[15])

    #############################################
    # not printed information: for comparison ONLY
    # 9.submission date/time
    info.append(lis[0])
    #############################################

    print(info)

    return info, ind_res


#   not needed currently, for final program
def getinfo_final(lis):
    pass


def zoom_info(lis):
    # session number
    lis[1], lis[0] = lis[0], lis[1]
    lis[0] = ''

    # date, time
    date = lis.pop(1)
    lis[1] += ', ' + date

    # remove zoom meeting title
    lis.pop(-2)

    # add extra words
    lis[2] = 'Meeting ID: ' + lis[2]
    lis[3] = 'Password: ' + lis[3]

    return lis


def order(lis):
    for i in range(len(lis) - 1):
        for j in range(i+1, len(lis)):
            name1, name2 = lis[i][2] + lis[i][3], lis[j][2] + lis[j][3]
            remove = string.punctuation + string.whitespace
            if name1.translate(remove) > name2.translate(remove):
                lis[i], lis[j] = lis[j], lis[i]
    # print(lis)


def sort(lis):
    # used bubble sort for simplicity
    for i in range(len(lis) - 1):
        for j in range(i + 1, len(lis)):
            # compare age
            if lis[i][4] > lis[j][4]:
                lis[i], lis[j] = lis[j], lis[i]
            # if same age: compare submission time
            elif lis[i][4] == lis[j][4]:
                date_time1, date_time2 = lis[i][-1], lis[j][-1]
                time1, time2 = timestamp(date_time1, date_time2)
                if time1 > time2: lis[i], lis[j] = lis[j], lis[i]


def timestamp(date1, date2):
    str1 = time.mktime(datetime.strptime(date1, "%Y-%m-%d %H:%M:%S").timetuple())
    str2 = time.mktime(datetime.strptime(date2, "%Y-%m-%d %H:%M:%S").timetuple())
    return str1, str2


def output_info(service, master_list, same_tab, flag):
    if same_tab:
        # METHOD 1
        # print all in same tab
        for i in range(1, 9):
            master_list[i].extend([[], []])
        curr_row = 1
        for i in range(1, 9):
            service.spreadsheets().values().update(spreadsheetId=Output_ID, range="All Sessions!A{}".format(curr_row),
                                                   valueInputOption="USER_ENTERED",
                                                   body={"values": master_list[i]}).execute()
            curr_row += len(master_list[i])

        return curr_row
    else:
        # METHOD 2
        # print each session in different tab
        for i in range(1, 9):
            service.spreadsheets().values().update(spreadsheetId=Output_ID, range="Session{}!A1".format(i),
                                                   valueInputOption="USER_ENTERED",
                                                   body={"values": master_list[i]}).execute()
            num = 15 - len(master_list[i]) + 1
            if flag: num += 1
            info = str(num) + " spots left"
            service.spreadsheets().values().update(spreadsheetId=Output_ID,
                                                   range="Session{}!A{}".format(i, len(master_list[i]) + 2),
                                                   valueInputOption="USER_ENTERED",
                                                   body={"values": [[info]]}).execute()
        return -1


def output_info2(service, master_list, row, flag):
    total = 0
    res = []
    for i in range(1, 9):
        temp = master_list[i]
        num = 15 - len(temp) + 3
        if flag:
            num += 1
            total -= 1
        res.append([f"Session {i}:", str(num) + " spots left"])
        total += len(temp) - 3

    res.extend([[], ["Total:", str(total) + ' registrations']])

    service.spreadsheets().values().update(spreadsheetId=Output_ID, range="All Sessions!A{}".format(row),
                                           valueInputOption="USER_ENTERED",
                                           body={"values": res}).execute()


def main():
    # Call the Sheets API
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # get input for current row in web-flow sheet
    row = input("Enter current row: ")

    # get input from web-flow sheet
    input1 = sheet.values().get(spreadsheetId=Webflow_ID,
                                range=f"Sheet1!A2:W{row}").execute()
    performer_info = input1.get('values', [])

    # get input from zoom link sheet
    input2 = sheet.values().get(spreadsheetId=Zoom_links_ID,
                                range="May 2021!A2:G9").execute()
    zoom_links = input2.get('values', [])

    # just personal notes ... nothing going on here ...
    """
    INPUT DATA
    0.Submission Time	1.Email Address   2.Performer's First Name  3.Performer's Last Name	
    4.Parent's First Name   5.Parent's Last Name	6.Performer's Age	7.Describe your experience	
    8.When would you like to perform?	9.What will your performance consist of?	10.What instrument?	
    11.Title of piece	12.Key	13.Opus	14.Movement	15.Composer Full Name   16.Length of Performance
    17.Number of Performers	    18.Day 1 changing piece acknowledgement	    19.Performance Method
    20.Answer host questions?	21.How did you hear about us?   22.Receive email updates?


    PROCESSED DATA
    0.Session Number   1.performer email   2.performer full name   3.performer first name 
    4.performer age    5.interview decision   6.performance type   7.performance piece
    8.composer name    9.submission date/time
    """

    # initialize
    title = ['Session #', 'Performer Email', 'Order', 'Performer Full Name', 'Performer First Name', 'Age',
             'Interview Decision', 'Performance Type', 'Performance Piece Name', 'Composer', 'Youtube Link']
    master_list = {i + 1: [] for i in range(8)}
    print_zoom_links = True
    same_tab = True
    hosts = []

    # organize zoom links
    for i in range(len(zoom_links)):
        zoom_info(zoom_links[i])

    # sort original list
    order(performer_info)

    # remove duplicates
    ind = 0
    while ind < len(performer_info) - 1:
        remove = string.punctuation + string.whitespace

        name1, name2 = performer_info[ind][2] + performer_info[ind][3], performer_info[ind+1][2] \
                       + performer_info[ind+1][3]
        if name1.translate(remove) == name2.translate(remove) \
                and performer_info[ind][1] == performer_info[ind+1][1] and \
                performer_info[ind][9] == performer_info[ind+1][9]:
            time1, time2 = timestamp(performer_info[ind][0], performer_info[ind+1][0])
            if time1 < time2:
                performer_info.pop(ind)
            else: performer_info.pop(ind+1)
        else:
            ind += 1

    # gather information for each performer & append to master list -> dictionary
    for i in range(len(performer_info)):
        performer, ind = getinfo(performer_info[i])
        master_list[ind].append(performer)

    # sort performers within each session
    for i in range(1, 9): sort(master_list[i])

    # delete extra info, add zoom links and titles
    for i in range(1, 9):
        for j in range(len(master_list[i])):
            temp = master_list[i][j]
            temp.pop()
            if temp[4] == 20: temp[4] = '18+'
            elif temp[4] == 30: temp[4] = "Prefer not to share"
            temp.insert(2, str(j+1))
        master_list[i].insert(0, title)
        if print_zoom_links:
            master_list[i].insert(0, zoom_links[i - 1])

    # output to file
    row = output_info(service, master_list, same_tab, print_zoom_links)
    if row >= 0:
        output_info2(service, master_list, row, print_zoom_links)

    print("done")
    print("performer information outputted to google sheets")

    # end of program


if __name__ == '__main__':
    main()

# (in the future) automatically create zoom meetings
# delete extra info(not needed currently)
# service.spreadsheets().values().batchClear(spreadsheetId=Output_ID,
#                                            body={'ranges': "Sheet1!A1:F8"}).execute()

# add MC/Host name and email to each session ls

