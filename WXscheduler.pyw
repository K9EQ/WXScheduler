# VIM :set softtabstop=4 shiftwidth=4 expandtab
"""
    A simple "settings" implementation.  Load/Edit/Save settings for your programs
    Uses json file format which makes it trivial to integrate into a Python program.  If you can
    put your data into a dictionary, you can save it as a settings file.

    Note that it attempts to use a lookup dictionary to convert from the settings file to keys used in 
    your settings window.  Some element's "update" methods may not work correctly for some elements.

    Copyright 2020 PySimpleGUI.com
    Licensed under LGPL-3
"""
import PySimpleGUI as sg
import time
from json import (load as jsonload, dump as jsondump)
from os import path
from pywinauto import Application

################ Global variable definitions #################
ExecutedCommands = []
Sched_nths=['every','1st','2nd','3rd','4th','5th']
Sched_nth_default=Sched_nths[0]
Sched_days=['day','Sun','Mon','Tue','Wed','Thu','Fri','Sat']
Sched_day_default=Sched_days[0]
Sched_hours=['00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23']
Sched_hour_default=Sched_hours[0]
Sched_minutes=['00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19',
               '20','21','22','23','24','25','26','27','28','29','30','31','32','33','34','35','36','37','38','39',
               '40','41','42','43','44','45','46','47','48','49','50','51','52','53','54','55','56','57','58','59']
Sched_minute_default=Sched_minutes[0]
Sched_TOT_minutes=['5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26',
             '27','28','29','30','31','32','33','34','35','36','37','38','39','40','41','42','43','44','45','46','47',
             '48','49','50','51','52','53','54','55','56','57','58','59','60']
Sched_TOT_minutes_default='15'

Sched_commands = [ 'none', 'Connect', 'Disconnect', 'Restart Wires-X App' ]

# Initialize Schedule default values at program startup 
Sched_description_default = ''
Sched_RoundQsoRoomConnection_default = True
Sched_AcceptCallsWhileInRoundQsoRooms_default = True
Sched_BackToRoundQsoAfterDisconnect_default = True
Sched_ReturnToRoomCheckbox_default = True
Sched_ReturnToRoomID_default = ''
Sched_UnlimitedTOT_default = False
Sched_command_default=Sched_commands[0]
Sched_cmdarg_default = ''

###################################################################
# Create a dictionary indexed by first 2 characters of a radio-ID #
###################################################################
RadioNameFromRadioID = {
    'E0' : 'FT1-D',
    'E5' : 'FT2-D',
    'EA' : 'FT3-D',
    'EB' : 'FT5-D',
    'F0' : 'FTM-400',
    'F5' : 'FTM-100',
    'FA' : 'FTM-300',
    'FB' : 'FTM-200',
    'FC' : 'FTM-500',
    'G0' : 'FT-991',
    'H0' : 'FTM-3200',
    'H5' : 'FT-70D',
    'H6' : 'FT-70D',
    'HA' : 'FTM-3207',
    'HF' : 'FTM-7250',
    'R0' : 'DR-1X',
    'R5' : 'DR-2X',
    }

#######################################################################
# Convert Degrees Minutes Seconds to decimal maidenhead.
#######################################################################
def dms2dd(degrees, minutes, seconds, hemisphere):
    dd = float(degrees) + float(minutes)/60 + float(seconds)/(60*60);
    if hemisphere == 'S' or hemisphere == 'W':
        dd *= -1 # makes dd negative
    return dd

#######################################################################
# Convert Degrees Minutes Seconds input string to Grid Square string
#######################################################################
def convertDegMinSec_to_GridSquare(dmsStr):

    if len(dmsStr) == 0:
        return ''

    keepers = '0123456789NEWS: '
    newStr = ''.join([x for x in dmsStr if x in keepers])
    newStr = newStr.replace(':', ' ') # change all colons to spaces
    dmsList = newStr.split(' ') # split the line into a list of fields

    # get the latitude/longitude values from the list skipping over empty entries
    try:
        i = 0
        while i < len(dmsList):

            if len(dmsList[i]) == 0:
                i += 1
            elif dmsList[i] == 'N' or dmsList[i] == 'S':
                LatHemisphere = dmsList[i]
                LatDegrees = float(dmsList[i+1])
                LatMinutes = float(dmsList[i+2])
                LatSeconds = float(dmsList[i+3])
                i += 4
            elif dmsList[i] == 'E' or dmsList[i] == 'W':
                LonHemisphere = dmsList[i]
                LonDegrees = float(dmsList[i+1])
                LonMinutes = float(dmsList[i+2])
                LonSeconds = float(dmsList[i+3])
                break
            else:
                i += 1

    except:
        #print('PROBLEM: convertDegMinSec_to_GridSquare(%s)' % dmsStr)
        return 'xxxxxx'

    # Convert DMS to decimal maidenhead as a positive value
    LatDecimal = dms2dd(LatDegrees, LatMinutes, LatSeconds, LatHemisphere)
    LonDecimal = dms2dd(LonDegrees, LonMinutes, LonSeconds, LonHemisphere)

    # Adjust decimal values for Grid Square conversion
    LatDecimal +=  90.0
    LonDecimal += 180.0

    # Convert latitude/longitude decimal values to grid square characters
    LonField = chr(int(LonDecimal / 20) + ord('A'))
    LatField = chr(int(LatDecimal / 10) + ord('A'))
    LonGrid  = chr(int((LonDecimal % 20) / 2) + ord('0'))
    LatGrid  = chr(int((LatDecimal % 10) / 1) + ord('0'))
    LonSubsq = chr(int((LonDecimal % 2) * (60 / 5)) + ord('a'))
    LatSubsq = chr(int((LatDecimal % 1) * (60 / 2.5)) + ord('a'))

    # return the 6 character Grid Square string (example: "EN51wu")
    return LonField + LatField + LonGrid + LatGrid + LonSubsq + LatSubsq

##################################################################################
# Check a string for a amateur radio callsign, and if found
# convert the callsign into a HTML lookup.
#       <a href=https://www.qrz.com/db/callsign>callsign</a>
##################################################################################
def callsign2html(s):
    try:
        # scan string for digits and create a list of digit positions
        digitPositions = []
        strLen = len(s)
        for i in range(0,strLen):
            if s[i].isdigit():
                digitPositions.append(i)

        if len(digitPositions) == 0:
            # No digits in string - Assume Room Name
            return s

        # qrz.com recognizes / as a callsign delimiter, but not -
        callsign = s.replace('-', '/')

        # assume everything before / is a valid callsign
        if '/' in callsign:
            return '<a href=https://www.qrz.com/db/%s>%s</a>' % (callsign,s)

        # no delimiter - check for name concatenation
        if digitPositions[0] == 0 or (len(digitPositions) > 1 and digitPositions[1] < 4):
            # first char is digit or multiple digits near beginning of callsign - assume International Callsign
            return '<a href=https://www.qrz.com/db/%s>%s</a>' % (callsign,s)

        # single digit
        digitPos = digitPositions[0]
        prefixStr = callsign[0:digitPos]
        digitStr = callsign[digitPos]
        # if 3 or less chars after the digit, then use callsign as is
        if (strLen - digitPos) <= 4:
            return '<a href=https://www.qrz.com/db/%s>%s</a>' % (callsign,s)

        # assume 3 letter suffix and insert a / between callsign and name
        suffixStr = callsign[digitPos+1:digitPos+4]
        nameStr = callsign[digitPos+4:strLen]
        callsignStr = prefixStr + digitStr + suffixStr
        return '<a href=https://www.qrz.com/db/%s>%s</a>' % (callsignStr,callsignStr + '/' + nameStr)
        
    except:
        return '*%s*' % s

##############################################################
# This function is called after the main loop determines
# that ~/Documents/WIRESXA/AccHistory/WiresAccess.log
# has been updated by the Wires-X application.
# 1) Reads the entire WiresAccess.log into memory
# 2) Calculates hash to determine if the contents have
#    actually changed to avoid rewriting the same ouput.
# 3) Parses the contents to produce 2 outputs:
#    - returns new hash and formatted screen output
#    - LastHeard.html file with qrz.com callsign lookup URLs
##############################################################
def refreshLastHeard(settings, previous_cksum):

    ####################################################################################################################################
    # Class object definition encapsulating last heard information from WiresAccess.log with '%' separated lines
    #   W9LBR-BILL%E5cN9%W9LBR-BILL%2019/09/25 19:44:39%V-CH%543155503750745F7C6C2078002020202020%N:41 50' 42" / W:088 07' 58"%0%%%%%
    ####################################################################################################################################
    class UserInfo:
        def __init__(self, f0, f1, f2, f3, f4, f5, f6):
            self.nodeName  = f0
            if len(f1) > 0:
              self.radioID = f1
            else:
              self.radioID = '_null'
            self.callsign  = f2
            self.timestamp = f3
            self.source    = f4
            self.data      = f5
            self.gridsq    = convertDegMinSec_to_GridSquare(f6)

    outputScreen = []
    outputHTML = []
    debuglines = []

    # read the entire file contents
    # NOTE: WXaccesslog sometimes contains non utf-8 characters
    try:
        with open(settings['WXaccesslog'], encoding='utf-8', errors='backslashreplace', mode='r') as content_file:
            content = content_file.read()
    except Exception as e:
        outputScreen.append( 'Exception reading %s::%s' % (settings['WXaccesslog'], e) )
        return previous_cksum, None

    # check if the contents of WiresAccess.log have actually changed
    current_cksum = hash(content)
    if current_cksum == previous_cksum:
        return current_cksum, ''  # nothing to update

    # split content into lines
    accessLogLines = content.splitlines()

    for line in accessLogLines:

        # split line into fields
        f = line.split('%')

        if len(f) < 7:
            # ignore the line
            debuglines.append('Ignoring AccLog %d [%s]' % (len(f), line))
            continue

        # update UserInfo class with field strings
        ui = UserInfo(f[0], f[1], f[2], f[3], f[4], f[5], f[6])

        if   ui.source == 'Net'  : source = 'Internet'
        elif ui.source == 'V-CH' : source = 'Local/RF'
        elif ui.source == 'Room' : source = 'Room    '
        else                     : source = '?source?'

        if ui.radioID.isdigit():
            radioName = 'Room/Node'
        else:
            # determine Radio Name from first 2 chars of ui.radioID
            try:
                radioName = RadioNameFromRadioID[ui.radioID[:2]]
            except:
                # corrupt radioID -- set radioName to radioID
                radioName = '?%s?' % ui.radioID
                debuglines.append('Corrupt radioID {%s} [%s]' % (ui.radioID, line))
            
        if ui.callsign == ui.nodeName:
            outputScreen.append( '%s %s [%s] %-7.7s %s %s' % (ui.timestamp,source,ui.radioID,radioName,ui.callsign,ui.gridsq) )
            outputHTML.append( '%s %s [%s] %-7.7s %s %s' % (ui.timestamp,source,ui.radioID,radioName,callsign2html(ui.callsign),ui.gridsq) )
        else:
            outputScreen.append( '%s %s [%s] %s %s {%s} %s' % (ui.timestamp,source,ui.radioID,radioName,ui.callsign,ui.nodeName,ui.gridsq) )
            outputHTML.append( '%s %s [%s] %s %s {%s} %s' % (ui.timestamp,source,ui.radioID,radioName,callsign2html(ui.callsign),callsign2html(ui.nodeName),ui.gridsq) )

    # sort outputHTML lines by the timestamp field at the beginning of each output line
    outputHTML.sort(reverse=True)

    try:
        # overwrite file if it already exists
        lhFile = open(settings['WXlastheardHTML'], 'w')
        lhFile.write('<!DOCTYPE html>\n')
        lhFile.write('<head>\n')
        lhFile.write('</head>\n')
        lhFile.write('<h1>Wires-X Last Heard Log</h1>\n')
        lhFile.write('<pre>\n')
        lhFile.write('<code>\n')
        for line in outputHTML:
            lhFile.write(line + '\n')
        lhFile.write('</code>\n')
        lhFile.write('</pre>\n')
        lhFile.write('</html>\n')
        lhFile.close()

    except Exception as e:
        outputScreen.append(f'LastHeardLog: {e}')

    # Display debug output on the "console"
    if len(debuglines) > 0:
        for line in debuglines:
            print('debug:',line)

    # reverse sort outputScreen lines by the timestamp field at the beginning of each output line
    outputScreen.sort(reverse=True)

    return current_cksum, outputScreen

##################### Make a Delete Event window #####################
def create_delete_event_window(settings):
    sg.theme(settings['theme'])

    def TextLabel(text): return sg.Text(text+':', justification='r', size=(15,1))

    schedEVs = get_scheduled(settings)

    layout = [  [sg.Text('Delete Event from Schedule', font='Any 15')],
                [TextLabel('Choose one'),sg.Combo(schedEVs,default_value=schedEVs[0],size=(80,5),font=('Consolas',9),key='-DELETE EVENT-')],
                [sg.Button('Delete'), sg.Button('Cancel')]  ]

    window = sg.Window('Delete Event from Schedule', layout, finalize=True)

    return window

##################### Make an Add Event window #####################
def create_add_event_window(settings):
    sg.theme(settings['theme'])

    def TextLabel(text): return sg.Text(text+':', justification='r', size=(15,1))

    layout = [  [sg.Text('Add Event to Schedule', font='Any 15')],
                [TextLabel('Occurs'),sg.Combo(Sched_nths,size=(6,6),default_value=Sched_nth_default,key='-OCCURS-')],
                [TextLabel('Week Day'),sg.Combo(Sched_days,size=(4,8),default_value=Sched_day_default,key='-DOW-')],
                [TextLabel('Hour'),sg.Combo(Sched_hours,size=(3,20),default_value=Sched_hour_default,key='-HOUR-')],
                [TextLabel('Minute'),sg.Combo(Sched_minutes,size=(3,20),default_value=Sched_minute_default,key='-MINUTE-')],
                [TextLabel('Description'),sg.Input(default_text=Sched_description_default,key='-DESCRIPTION-')],
                [sg.Checkbox('Round QSO Room connection (ON:Permit/OFF:Reject)',default=Sched_RoundQsoRoomConnection_default,key='-RoundQsoRoomConnection-')],
                [sg.Checkbox('Accept calls while in Round QSO Rooms',default=Sched_AcceptCallsWhileInRoundQsoRooms_default,key='-AcceptCallsWhileInRoundQsoRooms-')],
                [sg.Checkbox('Back to Round QSO after disconnect',default=Sched_BackToRoundQsoAfterDisconnect_default,key='-BackToRoundQsoAfterDisconnect-')],
                [sg.Checkbox('Return to Room', default=Sched_ReturnToRoomCheckbox_default,key='-ReturnToRoomCheckbox-')],
                [TextLabel('Room ID'),sg.Input(default_text=Sched_ReturnToRoomID_default,key='-ReturnToRoomID-')],
                [sg.Checkbox('Unlimited TOT', default=Sched_UnlimitedTOT_default,key='-UnlimitedTOT-')],
                [TextLabel('TOT(TimeOut Timer)'),sg.Combo(Sched_TOT_minutes,size=(3,20),default_value=Sched_TOT_minutes_default,key='-TimeOutTimer-')],
                [TextLabel('Command'),sg.Combo(Sched_commands,size=(32,32),default_value=Sched_command_default,key='-COMMAND-')],
                [TextLabel('Argument'),sg.Input(default_text=Sched_cmdarg_default,key='-CMDARG-')],
                [sg.Button('Update'), sg.Button('Cancel')]  ]

    window = sg.Window('Add Event to Schedule', layout, finalize=True)

    return window

def add_event_to_schedule(settings, event_values):
    global Sched_nth_default
    global Sched_day_default
    global Sched_hour_default
    global Sched_minute_default
    global Sched_description_default
    global Sched_RoundQsoRoomConnection_default
    global Sched_AcceptCallsWhileInRoundQsoRooms_default
    global Sched_BackToRoundQsoAfterDisconnect_default
    global Sched_ReturnToRoomCheckbox_default
    global Sched_ReturnToRoomID_default
    global Sched_UnlimitedTOT_default
    global Sched_TimeOutTimer_default
    global Sched_command_default
    global Sched_cmdarg_default

    # When -DOW- is specified to be 'day' ensure that -OCCURS- will be 'every' 
    if event_values['-DOW-'] == 'day':
        event_values['-OCCURS-'] = 'every'

    # update defaults to minimize re-entering Add Event values next time
    Sched_nth_default = event_values['-OCCURS-']
    Sched_day_default = event_values['-DOW-']
    Sched_hour_default = event_values['-HOUR-']
    Sched_minute_default = event_values['-MINUTE-']
    Sched_description_default = event_values['-DESCRIPTION-']
    Sched_RoundQsoRoomConnection_default = event_values['-RoundQsoRoomConnection-']
    Sched_AcceptCallsWhileInRoundQsoRooms_default = event_values['-AcceptCallsWhileInRoundQsoRooms-']
    Sched_BackToRoundQsoAfterDisconnect_default = event_values['-BackToRoundQsoAfterDisconnect-']
    Sched_ReturnToRoomCheckbox_default = event_values['-ReturnToRoomCheckbox-']
    Sched_ReturnToRoomID_default = event_values['-ReturnToRoomID-']
    Sched_UnlimitedTOT_default = event_values['-UnlimitedTOT-']
    Sched_TimeOutTimer_default = event_values['-TimeOutTimer-']
    Sched_command_default = event_values['-COMMAND-']
    Sched_cmdarg_default = event_values['-CMDARG-']

    is_error = ''

    if not event_values['-OCCURS-'] in Sched_nths:
        is_error += '%s is not a valid Occurs setting\n' % event_values['-OCCURS-']

    if not event_values['-DOW-'] in Sched_days:
        is_error += '%s is not a valid Day of Week setting\n' % event_values['-DOW-']

    if not event_values['-HOUR-'] in Sched_hours:
        is_error += '%s is not a valid Hour setting\n' % event_values['-HOUR-']

    if not event_values['-MINUTE-'] in Sched_minutes:
        is_error += '%s is not a valid Minute setting\n' % event_values['-MINUTE-']

    command = event_values['-COMMAND-']

    if not command in Sched_commands:
        is_error += '%s is not a valid command\n' % command

    # Validate 'Connect' Node/Room value and verify that 'Round QSO Room connection' is checked
    if command == 'Connect':
        if_error = 'Argument must be a valid Node or Room number (Between 10000 and 99999 inclusive)\n'
        try:
            num = int(event_values['-CMDARG-'])
            if num < 10000 or num > 99999:
                is_error += if_error
        except:
            is_error += if_error

        if event_values['-RoundQsoRoomConnection-'] == False:
            is_error += 'Error: [%s %s] will be Rejected because []Round QSO Room connection is unchecked' % (command, event_values['-CMDARG-'])
    else:
        # Blank out cmdarg for all other commands
        event_values['-CMDARG-'] = ''
        Sched_cmdarg_default = ''

    # Validate ReturnToRoomID if ReturnToRoom is checked
    if event_values['-ReturnToRoomCheckbox-'] == True:
        if len(event_values['-ReturnToRoomID-']) == 5:
            if_error = 'Numeric Return to Room ID must be a valid Room number (Between 20000 and 89999 inclusive)\n'
            try:
                num = int(event_values['-ReturnToRoomID-'])
                if num < 20000 or num > 89999:
                    is_error += if_error
            except:
                is_error += if_error
        elif len(event_values['-ReturnToRoomID-']) == 0:
            is_error += 'Return to Room ID must be supplied\n'
        elif len(event_values['-ReturnToRoomID-']) > 16:
            is_error += 'Max length Return to Room ID is 16 characters\n'

    # Validate TimeOutTimer value if UnlimitedTOT is unchecked
    if event_values['-UnlimitedTOT-'] == False:
        if_error = 'TimeOutTimer must be between 5 and 60 inclusive)\n'
        try:
            num = int(event_values['-TimeOutTimer-'])
            if num < 5 or num > 60:
                is_error += if_error
        except:
            is_error += if_error
    
    if len(is_error) > 0:
        return False, is_error

    # create key and values for settings
    new_key = '@' + event_values['-OCCURS-'] +'-'+ \
                    event_values['-DOW-']    +'-'+ \
                    event_values['-HOUR-']   +'-'+ \
                    event_values['-MINUTE-']
    new_event = [event_values['-OCCURS-'],
                 event_values['-DOW-'],
                 event_values['-HOUR-'],
                 event_values['-MINUTE-'],
                 event_values['-DESCRIPTION-'],
                 event_values['-RoundQsoRoomConnection-'],
                 event_values['-AcceptCallsWhileInRoundQsoRooms-'],
                 event_values['-BackToRoundQsoAfterDisconnect-'],
                 event_values['-ReturnToRoomCheckbox-'],
                 event_values['-ReturnToRoomID-'],
                 event_values['-UnlimitedTOT-'],
                 event_values['-TimeOutTimer-'],
                 event_values['-COMMAND-'],
                 event_values['-CMDARG-'],
                ]

    # if event key already in settings, show previous event 
    if new_key in settings:
        previous_event = settings[new_key]
        ret_str = 'WARNING: Replaced [%s %s %s %s %s] with [%s %s %s %s %s]' % (
                        previous_event[0],
                        previous_event[1],
                        previous_event[2],
                        previous_event[3],
                        previous_event[4],
                        new_event[0],
                        new_event[1],
                        new_event[2],
                        new_event[3],
                        new_event[4] )
    else:
        ret_str = ''

    # Update settings dictionary
    try:
        settings[new_key] = new_event
    except Exception as e:
        return False, f'settings dict update EXCEPTION {e}'
    
    return True, ret_str

##################### Return a list of Scheduled Events #####################
def get_scheduled(settings):

    scheduled = []
    for key in settings:
        # only process settings key that has a leading '@'
        if key[0] == '@':
            vals = settings[key]
            if len(vals[4]) > 0: description='('+vals[4]+')'
            else: description=''
            scheduled.append('%5s %3s %s:%s %s' % (vals[0], vals[1], vals[2], vals[3], description))

    if len(scheduled) > 0:
        return scheduled

    return ['<empty>']

##################### Return current date, time, DOW and nthDOW in a dictionary #####################
def get_current_date_time_dow_nth():
    date_time = {}

    tm = time.localtime()
    date_time['yyyy'] = time.strftime('%Y', tm)  # yyyy
    date_time['mm']   = time.strftime('%m', tm)  # mm
    date_time['dd']   = time.strftime('%d', tm)  # dd
    date_time['HH']   = time.strftime('%H', tm)  # HH
    date_time['MM']   = time.strftime('%M', tm)  # MM
    date_time['SS']   = time.strftime('%S', tm)  # SS
    date_time['dow']  = time.strftime('%a', tm)  # weekday as "Sun", "Mon", etc.
    date_time['nth']  = str(Sched_nths[int((tm.tm_mday + 6) / 7)]) # Nth DOW

    return date_time

##################### Check if it is time to execute a scheduled Wires-X command #####################
def is_scheduled_time(event_to_match, settings):

    etm = event_to_match.split('-')
    match_nth    = etm[0]
    match_dow    = etm[1]
    match_hour   = etm[2]
    match_minute = etm[3]

    for key in settings:
        # only process settings key that has a leading '@'
        if key[0] == '@':
            ev_data = settings[key]
            if ev_data[3] != match_minute:
                continue
            if ev_data[2] != match_hour:
                continue
            if ev_data[1] != 'day':     # inverse logic that skips match_dow check when ev_data[1] == 'day'
                if ev_data[1] != match_dow:
                    continue
            if ev_data[0] == match_nth or ev_data[0] == 'every':
                # key matched -- return the associated data
                return True, ev_data

    # No keys were matched
    noData = []
    return False, noData

##########################################################################################################
# Automate
#|    | CheckBox - 'Back to Round QSO after disconnect'    (L448, T141, R749, B160)
#|    | ['Back to Round QSO after disconnect', 'Back to Round QSO after disconnectCheckBox', 'CheckBox6']
#|    | child_window(title="Back to Round QSO after disconnect", class_name="Button")
##########################################################################################################
def SetBacktoRoundQSOafterdisconnectCheckBox(app, action):

    response = Display_File_Settings_submenu(app, 'Call settings')
    if response != 'ok':
        return response

    try:
        dlg = app.Settings
        chkbox = dlg.BacktoRoundQSOafterdisconnectCheckBox
        if action :
            chkbox.check()
        else:
            chkbox.uncheck()
        dlg.OK.click()

    except:
        return 'EXCEPTION during SetBacktoRoundQSOafterdisconnectCheckBox()'

    return 'OK'

################################################################################
# Automate
#|    | CheckBox - 'Return to Room'    (L448, T217, R686, B236)
#|    | ['Return to Room', 'CheckBox7', 'Return to RoomCheckBox']
#|    | child_window(title="Return to Room", class_name="Button")
################################################################################
def SetReturntoRoomCheckBox(app, action, clickOK=True):

    response = Display_File_Settings_submenu(app, 'Call settings')
    if response != 'ok':
        return response
    SetReturntoRoomCheckBox(app, True, False) # Check Return to Room, but don't click OK

    try:
        dlg = app.Settings
        chkbox = dlg.ReturntoRoomCheckBox
        if action :
            chkbox.check()
        else:
            chkbox.uncheck()
        if clickOK :
            dlg.OK.click()

    except:
        return 'EXCEPTION during SetReturntoRoomCheckBox'

    return 'OK'

#############################################################################
# Automate (Return to)Room ID
#|    | CheckBox - 'Return to Room'    (L710, T236, R948, B255)
#|    | ['CheckBox7', 'Return to Room', 'Return to RoomCheckBox']
#|    | child_window(title="Return to Room", class_name="Button")
#|    | 
#|    | Static - 'Room ID'    (L711, T266, R778, B285)
#|    | ['Room IDStatic2', 'Room ID2', 'Static31']
#|    | child_window(title="Room ID", class_name="Static")
#|    | 
#|    | Edit - ''    (L787, T262, R1002, B285)
#|    | ['Room IDEdit2', 'Edit22']
#|    | child_window(class_name="Edit")
#############################################################################
def SetReturntoRoomID(app, roomNum):

    # Must first uncheck 'Return to Room' because the click() below will toggle it
    SetReturntoRoomCheckBox(app, False, False) # Uncheck Return to Room, but don't click OK

    try:
        dlg = app.Settings
        dlg.ReturntoRoomEdit.click()    # click() will cause unchecked box to be checked
        dlg.RoomIDEdit2.click()         # click() the edit box so it can be typed into
        dlg.type_keys('{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}{DEL}') # Clear any previous entry 
        dlg.type_keys(roomNum)
        dlg.OK.click()

    except:
        return 'EXCEPTION during SetReturntoRoomID()'

    return 'OK'

#######################################################################
# Automate application ConnectTo menu
#######################################################################
def ConnectToRoom(app, roomNum):

    try:
        # Select menu item 'Connect(C)' and its 'Connect To(T)' item to open the 'Input ID' dialog
        exstr = 'app.WiresX.menu_select(Connect(C)->Connect To(T))'
        app.WiresX.menu_select('Connect(C)->Connect To(T)')
        # set the Node/Room text box in the 'Input ID' dialog
        exstr = 'app.InputID.Edit.set_edit_text(%s)' % roomNum
        app.InputID.Edit.set_edit_text(roomNum)
        # Click the 'OK' button in the 'Input ID' dialog
        exstr = 'app.InputID.OK.click()'
        app.InputID.OK.click()
    except:
        return exstr

    return 'OK'

#######################################################################
# Automate application Disconnect menu
#######################################################################
def DisconnectFromAnyRoom(app):

    try:
        # Select menu item 'Connect(C)' and its 'Disconnect(D)' item to disconnect the current ROOM
        app.WiresX.menu_select('Connect(C)->Disconnect(D)')
    except:
        return 'No Node/Room was Connected'

    return 'OK'

###############################################################
# Cause the Wires-X app to display a "File->Settings(P)" menu
# before making automation updates to the specified sub-menu.
###############################################################
def Display_File_Settings_submenu(app, submenuStr):

    try:
        app.WiresX.menu_select('File(F)->Settings(P)') # pop-up Settings window via File menu
        app.Settings.type_keys('{HOME}{DOWN}')         # go to top off Tree and down one

        # Each DOWN causes the screen to update, verified by treeItem[6] containing submenuStr
        for i in range(0,12):
            treeItem = app['Settings'].children()   # Refresh treeItem list for each loop iteration to get screen updates
            if submenuStr in str(treeItem[6]):
                #print('vvvvv----%s----vvvvv' % submenuStr)
                #app.Settings.print_control_identifiers() # on console
                #print('^^^^^----%s----^^^^^' % submenuStr)
                return 'ok'
            app.Settings.type_keys('{DOWN}')    # cause the next sub-menu to be displayed

        # Did not find submenuStr -- Close pop-up by typing an ESC -- return ERROR
        app.Settings.type_keys('{ESC}')
        return 'F-P-SUBMENU \'' + submenuStr + '\' NOT FOUND'

    except:
        return 'EXCEPTION in Display_File_Settings_submenu()'

################################################################################################################
# Automate
#|    | CheckBox - 'Accept calls while in Round QSO Rooms'    (L448, T124, R749, B141)
#|    | ['Accept calls while in Round QSO RoomsCheckBox', 'CheckBox5', 'Accept calls while in Round QSO Rooms']
#|    | child_window(title="Accept calls while in Round QSO Rooms", class_name="Button")
################################################################################################################
def SetAcceptcallswhileinRoundQSORoomsCheckBox(app, action):

    response = Display_File_Settings_submenu(app, 'Call settings')
    if response != 'ok':
        return response

    try:
        dlg = app.Settings
        chkbox = dlg.AcceptcallswhileinRoomsCheckBox
        if action :
            chkbox.check()
        else:
            chkbox.uncheck()
        dlg.OK.click()

    except:
        return 'EXCEPTION during SetAcceptcallswhileinRoundQSORoomsCheckBox()'

    return 'OK'

######################################################################################################################################
# Automate
#|    | CheckBox - 'Round QSO Room connection (ON:Permit/OFF:Reject)'    (L448, T106, R810, B123)
#|    | ['Round QSO Room connection (ON:Permit/OFF:Reject)', 'Round QSO Room connection (ON:Permit/OFF:Reject)CheckBox', 'CheckBox4']
#|    | child_window(title="Round QSO Room connection (ON:Permit/OFF:Reject)", class_name="Button")
######################################################################################################################################
def SetRoundQSORoomconnectionCheckBox(app, action):

    response = Display_File_Settings_submenu(app, 'Call settings')
    if response != 'ok':
        return response

    try:
        dlg = app.Settings
        chkbox = dlg.RoundQSORoomconnectonCheckBox
        if action :
            chkbox.check()
        else:
            chkbox.uncheck()
        dlg.OK.click()

    except:
        return 'EXCEPTION during SetRoundQSORoomconnectionCheckBox()'

    return 'OK'

#############################################################################
# Automate
#| CheckBox - 'Unlimited TOT'    (L701, T158, R873, B177)
#|    | ['CheckBox13', 'Unlimited TOT', 'Unlimited TOTCheckBox']
#|    | child_window(title="Unlimited TOT", class_name="Button")
#############################################################################
def Set_Unlimited_TOT_checkbox(app, action):

    response = Display_File_Settings_submenu(app, 'General settings')
    if response != 'ok':
        return response

    try:
        dlg = app.Settings
        chkbox = dlg.UnlimitedTOTCheckBox
        if action :
            chkbox.check()
        else:
            chkbox.uncheck()
        dlg.OK.click()

    except:
        return 'EXCEPTION during Set_Unlimited_TOT_checkbox()'

    return 'OK'

##################################################################################################
# Automate
#|    | Edit - ''    (L435, T194, R922, B217)
#|    | ['TOT(TimeOut Timer)Edit', 'Edit8', 'TOT(TimeOut Timer)Edit0', 'TOT(TimeOut Timer)Edit1']
#|    | child_window(class_name="Edit")
# Note: Unlimited_TOT must be unchecked to be able to set TOT_TimeOutTimer
##################################################################################################
def Set_TOT_TimeoutTimer(app, minutes):

    response = Display_File_Settings_submenu(app, 'General settings')
    if response != 'ok':
        return response

    try:
        dlg = app.Settings
        # enter minutes into text box
        edit = dlg.TOTTimeOutTimerEdit
        edit.set_edit_text(minutes)
        dlg.OK.click()

    except:
        return 'EXCEPTION during Set_TOT_TimeoutTimer()'

    return 'OK'

#######################################################################
# Automate application Exit -- Wires-X App will automatically restart
#######################################################################
def ExitApplication(app):

    try:
        exstr = 'app.WiresX.menu_select(File(F)->Exit(E)) #1'
        app.WiresX.menu_select('File(F)->Exit(E)')
    except:
        # Settings window left open?
        try:
            exstr = 'app.Settings.Cancel.click()'
            app.Settings.Cancel.click()            # click Cancel
            time.sleep(0.3)
            exstr = 'app.WiresX.menu_select(File(F)->Exit(E)) #2'
            app.WiresX.menu_select('File(F)->Exit(E)')
        except:
            return exstr

    return 'OK'

def ForceDisconnectRoom(settings):

    # Connect to the Wires-X application
    try:
        if_exception = 'Application().connect() EXCEPTION: '
        app = Application(backend='win32').connect(path=settings['WXapplication'])
        # restore WIRES-X window only if minimized
        if_exception = 'app.window() EXCEPTION: '
        window = app.window(title=u'WIRES-X', visible_only=False).restore()
    except Exception as e:
        return if_exception + str(e)

    response = DisconnectFromAnyRoom(app)

    return '[Force Disconnect] {%s}' % response

############## Update Wires-X settings and execute Wires-X command #################
def performWXactions(eventData, settings):

    # Connect to the Wires-X application
    try:
        if_exception = 'Application().connect() EXCEPTION: '
        app = Application(backend='win32').connect(path=settings['WXapplication'])
        # restore WIRES-X window only if minimized
        if_exception = 'app.window() EXCEPTION: '
        window = app.window(title=u'WIRES-X', visible_only=False).restore()
    except Exception as e:
        return if_exception + str(e)

    try:
        app.Settings.Cancel.click() # click Cancel
    except:
        print('IGNORING: app.Settings.Cancel.click() EXCEPTION')

    ## pandas key to json settings file index mapping ######
    # '-OCCURS-' ---------------------------> eventData[0]
    # '-DOW-' ------------------------------> eventData[1]
    # '-HOUR-' -----------------------------> eventData[2]
    # '-MINUTE-' ---------------------------> eventData[3]
    # '-DESCRIPTION-' ----------------------> eventData[4]
    # '-RoundQsoRoomConnection-' -----------> eventData[5]
    # '-AcceptCallsWhileInRoundQsoRooms-' --> eventData[6]
    # '-BackToRoundQsoAfterDisconnect-' ----> eventData[7]
    # '-ReturnToRoomCheckbox-' -------------> eventData[8]
    # '-ReturnToRoomID-' -------------------> eventData[9]
    # '-UnlimitedTOT-' ---------------------> eventData[10]
    # '-TimeOutTimer-' ---------------------> eventData[11]
    # '-COMMAND-' --------------------------> eventData[12]
    # '-CMDARG-' ---------------------------> eventData[13]

    description = '(%s)' % eventData[4]

    # Update the Wires-X 'Call settings'

    # '-RoundQsoRoomConnection-' -----------> eventData[5]
    response = SetRoundQSORoomconnectionCheckBox(app, eventData[5])
    if response != 'OK':
        return description + '::' + response

    # '-AcceptCallsWhileInRoundQsoRooms-' --> eventData[6]
    # '-BackToRoundQsoAfterDisconnect-' ----> eventData[7]
    #### [7] can only be changed if [6] is True #####
    if eventData[6] == False and eventData[7] == False:
        response = SetAcceptcallswhileinRoundQSORoomsCheckBox(app, True)
        if response != 'OK':
            return description + '::' + response
        response = SetBacktoRoundQSOafterdisconnectCheckBox(app, False)
        if response != 'OK':
            return description + '::' + response
        response = SetAcceptcallswhileinRoundQSORoomsCheckBox(app, False)
        if response != 'OK':
            return description + '::' + response
    elif eventData[6] == False and eventData[7] == True:
        response = SetAcceptcallswhileinRoundQSORoomsCheckBox(app, True)
        if response != 'OK':
            return description + '::' + response
        response = SetBacktoRoundQSOafterdisconnectCheckBox(app, True)
        if response != 'OK':
            return description + '::' + response
        response = SetAcceptcallswhileinRoundQSORoomsCheckBox(app, False)
        if response != 'OK':
            return description + '::' + response
    else:
        response = SetAcceptcallswhileinRoundQSORoomsCheckBox(app, True)
        if response != 'OK':
            return description + '::' + response
        response = SetBacktoRoundQSOafterdisconnectCheckBox(app, eventData[7])
        if response != 'OK':
            return description + '::' + response

    # '-ReturnToRoomCheckbox-' -------------> eventData[8]
    # '-ReturnToRoomID-' -------------------> eventData[9]
    if eventData[8] == False:
        # First clear out the Return to Room ID that will leave the CheckBox checked
        response = SetReturntoRoomID(app, '')
        if response != 'OK':
            return description + '::' + response
        # Now uncheck 'Return to Room'
        response = SetReturntoRoomCheckBox(app, False)
        if response != 'OK':
            return description + '::' + response
    else:
        # Set the RoomID which will leave the Checkbox checked
        response = SetReturntoRoomID(app, eventData[9])
        if response != 'OK':
            return description + '::' + response

    # '-TimeOutTimer-' ---------------------> eventData[11]
    ## Note: The TimeOutTimer value can only be changed when UnlimitedTOT is unchecked
    response = Set_Unlimited_TOT_checkbox(app, False)
    if response != 'OK':
        return description + '::' + response
    response = Set_TOT_TimeoutTimer(app, eventData[11])
    if response != 'OK':
        return description + '::' + response

    # '-UnlimitedTOT-' ---------------------> eventData[10]
    response = Set_Unlimited_TOT_checkbox(app, eventData[10])
    if response != 'OK':
        return description + '::' + response

    # '-COMMAND-' --------------------------> eventData[12]
    command = eventData[12]
    # '-CMDARG-' ---------------------------> eventData[13]
    cmdarg = eventData[13]
    if command == 'Connect':
        # FYI--The App ignores Connecting to the same room that is already connected to.
        response = ConnectToRoom(app, cmdarg)
    elif command == 'Disconnect':
        response = DisconnectFromAnyRoom(app)
    elif command == 'Restart Wires-X App':
        response = ExitApplication(app)
    elif command != 'none':
        response = 'UNEXPECTED command [%s]' % command

    if command == 'none':
        return '%s {%s}' % (description, response)

    return '%s %s %s {%s}' % (description, command, cmdarg, response)

##################### Return a list of reverse sorted executed scheduled commands #####################
def get_executed_commands():

    if len(ExecutedCommands) <= 0:
        return [' ']

    return sorted(ExecutedCommands, reverse=True)

##################### Load/Save Settings File #####################
SETTINGS_FILE = path.expanduser('~/Documents/WIRESXA/WXscheduler.cfg')
SETTINGS_KEYS_TO_ELEMENT_KEYS = {'theme':'-THEME-',
                                 'WXapplication':'-WIRES-X EXE-',
                                 'WXaccesslog':'-WIRES-X ACCESS LOG-',
                                 'WXlastheardHTML':'-WIRES-X LAST HEARD-'}

def load_settings(settings_file):

    try:
        with open(settings_file, 'r') as f:
            settings = jsonload(f)

    except Exception as e:
        sg.popup_quick_message(f'exception {e}', 'No settings file found... will create one for you', keep_on_top=True, background_color='red', text_color='white')
        # Create default settings (except Scheduled Events) and save to file
        settings = {}
        settings['theme'] = sg.theme()
        settings['WXapplication'] = 'C:/Program Files (x86)/YAESUMUSEN/WIRES-X/wires-X.exe'
        settings['WXaccesslog']= path.expanduser('~/Documents/WIRESXA/AccHistory/WiresAccess.log')
        settings['WXlastheardHTML']= path.expanduser('~/Desktop/Wires-X_Last_Heard.html')
        save_settings(settings_file, settings, None)

    return settings

def save_settings(settings_file, settings, values):
    if values:      # if there are stuff specified by another window, fill in those values
        for key in SETTINGS_KEYS_TO_ELEMENT_KEYS:  # update window with the values read from settings file
            try:
                settings[key] = values[SETTINGS_KEYS_TO_ELEMENT_KEYS[key]]
            except Exception as e:
                print(f'Problem updating settings from window values. Key = {key}')

    with open(settings_file, 'w') as f:
        jsondump(settings, f)

    #sg.popup('Settings saved')
    sg.popup_quick_message('The settings have been saved', keep_on_top=True, background_color='green', text_color='white')

##################### Make a settings window #####################
def create_settings_window(settings):
    sg.theme(settings['theme'])

    def TextLabel(text): return sg.Text(text+':', justification='r', size=(15,1))

    layout = [  [sg.Text('Change Settings', font='Any 15')],
                [TextLabel('Theme'),sg.Combo(sg.theme_list(), size=(20, 20), key='-THEME-')],
                [TextLabel('Wires-X EXE'),sg.Input(key='-WIRES-X EXE-'), sg.FileBrowse(target='-WIRES-X EXE-')],
                [TextLabel('Wires-X Access Log'),sg.Input(key='-WIRES-X ACCESS LOG-'), sg.FileBrowse(target='-WIRES-X ACCESS LOG-')],
                [TextLabel('Wires-X Last Heard'),sg.Input(key='-WIRES-X LAST HEARD-'), sg.FileBrowse(target='-WIRES-X LAST HEARD-')],
                [sg.Button('Save'), sg.Button('Cancel')]  ]

    window = sg.Window('Settings', layout, keep_on_top=True, finalize=True)

    for key in SETTINGS_KEYS_TO_ELEMENT_KEYS:   # update window with the values read from settings file
        try:
            window[SETTINGS_KEYS_TO_ELEMENT_KEYS[key]].update(value=settings[key])
        except Exception as e:
            print(f'Problem updating PySimpleGUI window from settings. Key = {key}')

    return window

##################### Main Program Window & Event Loop #####################
def create_main_window(settings):
    sg.theme(settings['theme'])

    layout = [  [sg.T(size=(25,1),font=('Helvetica',20),justification='center',key='-DATE TIME DOW-')],
                [sg.T('Last Heard:'),sg.Listbox(['Waiting for %s to be updated'%settings['WXaccesslog']], size=(80, 10),font=('Consolas',9),key='-WIRES-X LAST HEARD-')],
                [sg.T('  Executed:'),sg.Listbox(get_executed_commands(), size=(80, 5),font=('Consolas',9),key='-EXECUTED-')],
                [sg.T('  Schedule:'),sg.Listbox(get_scheduled(settings), size=(80, 5),font=('Consolas',9))],
                [sg.Button('Add Event'), sg.Button('Delete Event'), sg.Button('Force Disconnect'), sg.Button('Change Settings')]  ]

    return sg.Window('Wires-X Scheduler (v1.3)', layout, finalize=True)


def main():

    dt = get_current_date_time_dow_nth()
    previous_minute = dt['MM']
    window, settings = None, load_settings(SETTINGS_FILE)

    while True: # Forever Loop - An iteration occurs when a window.read() returns an event, or the per second timeout occurs

        if window is None:
            # Create the main window if this is the first time through the loop, or window event processing below closed it
            # and cause LastHeard window to be immediately populated
            #
            previousModTime = 0.0
            previous_cksum = hash('nothing yet')
            window = create_main_window(settings)

        # Display the current time in the main window
        #
        dt = get_current_date_time_dow_nth()
        formatted_dt = '%s/%s/%s %s:%s:%s' % (dt['yyyy'],dt['mm'],dt['dd'],dt['HH'],dt['MM'],dt['SS'])  # "yyyy/mm/dd HH:MM:SS"
        window['-DATE TIME DOW-'].update('%s (%s %s)' % (formatted_dt, dt['nth'], dt['dow']) )          # + " (1st Mon)"

        # check if the WiresAccess.log modification time has changed
        #
        try:
            currentModTime = path.getmtime(settings['WXaccesslog']) 
        except Exception as e:
            window['-WIRES-X LAST HEARD-'].update(f'Exception: {e}')
        else:
            if currentModTime != previousModTime:
                previousModTime = currentModTime
                current_cksum, last_heard_lines = refreshLastHeard(settings, previous_cksum)
                if current_cksum != previous_cksum:
                    previous_cksum = current_cksum
                    window['-WIRES-X LAST HEARD-'].update(last_heard_lines)

        # If entering a new minute, then check if it is time to execute a scheduled Wires-X action
        #
        current_minute = dt['MM']
        if previous_minute != current_minute:
            event_to_match = dt['nth']+'-'+dt['dow']+'-'+dt['HH']+'-'+dt['MM'] # e.g. "1st-Mon-HH-MM"
            is_time, ev_data = is_scheduled_time(event_to_match, settings)
            if is_time:
                ## Execute scheduled actions ##
                report = performWXactions(ev_data, settings)
                ExecutedCommands.append('%s %s' % (formatted_dt, report))
                window['-EXECUTED-'].update(get_executed_commands())
            previous_minute = current_minute

        # calculate how many milli-seconds until the start of the next second
        #
        secs = time.time()
        delayms = 1000 - int((secs - int(secs)) * 1000.0)
 
        # Wait until a window event or timeout occurs
        #
        event, values = window.read(timeout=delayms)

        # Exit the forever loop and this program
        #
        if event in (sg.WIN_CLOSED, 'Exit'):
            break

        if event == 'Add Event':
            # the Add Event button in the main window has been clicked
            #
            event, event_values = create_add_event_window(settings).read(close=True)
            if event == 'Update':
                # the Add button in the Add Event window has been clicked
                #
                ret, ew_str = add_event_to_schedule(settings, event_values)
                if ret == True:
                    if len(ew_str) > 0:
                        sg.popup(ew_str, background_color='yellow', text_color='blue')
                    save_settings(SETTINGS_FILE, settings, None)
                    window.close()
                    window = None
                else:
                    sg.popup(ew_str, background_color='red', text_color='white')

        elif event == 'Delete Event':
            # the Delete Event button in the main window has been clicked
            #
            event, event_to_delete = create_delete_event_window(settings).read(close=True)
            if event == 'Delete':
                # the Delete button in the Delete Event window has been clicked
                #
                evL = event_to_delete['-DELETE EVENT-'].strip().replace(':',' ').split(' ')
                if len(evL) >= 4:
                    key = '@'+evL[0]+'-'+evL[1]+'-'+evL[2]+'-'+evL[3]
                    deleted_event = settings.pop(key)
                    save_settings(SETTINGS_FILE, settings, None)
                    window.close()
                    window = None

        elif event == 'Force Disconnect':
            # the Force Disconnect button in the main window has been clicked
            #
            report = ForceDisconnectRoom(settings)
            ExecutedCommands.append('%s %s' % (formatted_dt, report))
            window['-EXECUTED-'].update(get_executed_commands())

        elif event == 'Change Settings':
            # the Change Settings button in the main window has been clicked
            #
            event, values = create_settings_window(settings).read(close=True)
            if event == 'Save':
                # the Save button in the Change Settings window has been clicked
                #
                window.close()
                window = None
                save_settings(SETTINGS_FILE, settings, values)

    window.close()
    return

main()
