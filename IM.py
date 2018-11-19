import gspread
import gspread.utils as r2c
import re
import sys
import easygui
from oauth2client.service_account import ServiceAccountCredentials


# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('Inventory Management-576c00641514.json', scope)
gc = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
ss = gc.open_by_url('https://docs.google.com/spreadsheets/d/1Kqj6ybFQmvJw3DYZmj72Hz5jWpZo2tuehRbAxLVFj2w/edit#gid=1783603278')

##########################################

Response = ss.worksheet("Form Responses")


#########################################

last_row = False

index = 1

while last_row == False:

    cell = 'C' + str(index)

    Tag = Response.acell(cell). value

    if Tag == '':

        index = index - 1

        cell = 'C' + str(index)

        Quan_cell = 'D' + str(index)

        Avail_cell = 'E' + str(index)

        Mov_cell = 'F' + str(index)

        update_cell = 'G' + str(index)

        Tag = Response.acell(cell).value

        Quantity = Response.acell(Quan_cell).value

        Availability = Response.acell(Avail_cell).value

        Item_moved = Response.acell(Mov_cell).value

        last_row = True

    else:

        index = index + 1

Center = re.split('-',Tag)

Center = Center[0]

#########################################

DHA = ss.worksheet("DHA")

JT = ss.worksheet("JT")

DHA5 = ss.worksheet("DHA Phase 5")

COS = ss.worksheet("Cosmoplast")

GG = ss.worksheet("Gulberg Galleria")

LAF = ss.worksheet("Laforma")

########################################
try:

    if Center == 'DHA':

        Value = DHA.find(Tag)

        find = str(Value).split(' ')

        find = re.split('[R C]', find[1])

        del find[0]

        Quan_enter_cell = 'E' + str(find[0])

        Avail_enter_cell = 'F' + str(find[0])

        Mov_enter_cell = 'G' + str(find[0])

        DHA.update_acell(Quan_enter_cell,  Quantity)

        DHA.update_acell(Avail_enter_cell, Availability)

        DHA.update_acell(Mov_enter_cell, Item_moved)


########################################################

    elif Center == 'JT':

        Value = JT.find(Tag)

        find = str(Value).split(' ')

        find = re.split('[R C]', find[1])

        del find[0]

        Quan_enter_cell = 'E' + str(find[0])

        Avail_enter_cell = 'F' + str(find[0])

        Mov_enter_cell = 'G' + str(find[0])

        JT.update_acell(Quan_enter_cell, Quantity)

        JT.update_acell(Avail_enter_cell, Availability)

        JT.update_acell(Mov_enter_cell, Item_moved)

###############################################################

    elif Center == 'DHA5':

        Value = DHA5.find(Tag)

        find = str(Value).split(' ')

        find = re.split('[R C]', find[1])

        del find[0]

        Quan_enter_cell = 'E' + str(find[0])

        Avail_enter_cell = 'F' + str(find[0])

        Mov_enter_cell = 'G' + str(find[0])

        DHA5.update_acell(Quan_enter_cell, Quantity)

        DHA5.update_acell(Avail_enter_cell, Availability)

        DHA5.update_acell(Mov_enter_cell, Item_moved)

##################################################################

    elif Center == 'COS':

        Value = COS.find(Tag)

        find = str(Value).split(' ')

        find = re.split('[R C]', find[0])

        del find[0]

        Quan_enter_cell = 'E' + str(find[0])

        Avail_enter_cell = 'F' + str(find[0])

        Mov_enter_cell = 'G' + str(find[0])

        COS.update_acell(Quan_enter_cell, Quantity)

        COS.update_acell(Avail_enter_cell, Availability)

        COS.update_acell(Mov_enter_cell, Item_moved)

###################################################################

    elif Center == 'GG':

        Value = GG.find(Tag)

        find = str(Value).split(' ')

        find = re.split('[R C]', find[1])

        del find[0]

        Quan_enter_cell = 'E' + str(find[0])

        Avail_enter_cell = 'F' + str(find[0])

        Mov_enter_cell = 'G' + str(find[0])

        GG.update_acell(Quan_enter_cell, Quantity)

        GG.update_acell(Avail_enter_cell, Availability)

        GG.update_acell(Mov_enter_cell, Item_moved)

############################################################

    elif Center == 'LAF':

        Value = LAF.find(Tag)

        find = str(Value).split(' ')

        find = re.split('[R C]', find[1])

        del find[0]

        Quan_enter_cell = 'E' + str(find[0])

        Avail_enter_cell = 'F' + str(find[0])

        Mov_enter_cell = 'G' + str(find[0])

        LAF.update_acell(Quan_enter_cell, Quantity)

        LAF.update_acell(Avail_enter_cell, Availability)

        LAF.update_acell(Mov_enter_cell, Item_moved)
except:

    easygui.msgbox('Invalid tag number', 'Result')

    sys.exit("Value not found")

#######################################

easygui.msgbox('Data Updated', 'Done!')

Response.update_acell('C' + str(index),'Updated')

sys.exit("Done")



