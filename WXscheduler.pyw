# VIM :set softtabstop=4 shiftwidth=4 expandtab ai
"""
    A simple "settings" implementation.  Load/Edit/Save settings for your programs
    Uses json file format which makes it trivial to integrate into a Python program.  If you can
    put your data into a dictionary, you can save it as a settings file.

    Note that it attempts to use a lookup dictionary to convert from the settings file to keys used in 
    your settings window.  Some element's "update" methods may not work correctly for some elements.

    Copyright 2020 PySimpleGUI.com
    Licensed under LGPL-3
    
    WXScheduler.py was written by Bill, W9LBR
    Version 2.5
    Modifications by Chris, K9EQ: HamOperator.com
    2023-01-12 Version 2.5.1
     - Removed TxID errors from error logging (too many, no value)
     - Changed titlebar icon
     
"""
import sys
import PySimpleGUI as sg
import time
from datetime import datetime
import pytz
import json
from os import path
from pywinauto import Application

################ Global variable definitions #################
WIRESXA_PATHNAME_1 = R'~\OneDrive\Documents\WIRESXA'
WIRESXA_PATHNAME_2 = R'~\Documents\WIRESXA'
WIRESXA_PATHNAME = None         # Initialized in main()
SETTINGS_FILENAME = R'\WXscheduler.cfg'
SettingsFilePathname = None     # Initialized in main()
USER_DESKTOP_1 = R'~\OneDrive\Desktop'
USER_DESKTOP_2 = R'~\Desktop'
USER_DESKTOP = None             # Initialized in main()
WIRESX_APP = R'C:\Program Files (x86)\YAESUMUSEN\WIRES-X\wires-X.exe'

SETTINGS_KEYS_TO_ELEMENT_KEYS = {'theme':'-THEME-',
                                 #'lastLocation':'-last location-',#### Not needed because automatically updated whenever the window is moved ####
                                 'localTZ':'-LOCAL TZ-',
                                 'WXapplication':'-WIRES-X EXE-',
                                 'WXaccesslog':'-WIRES-X ACCESS LOG-',
                                 'WXlastheardHTML':'-WIRES-X LAST HEARD-',
                                 #'WXinfolog':'-WIRES-X INFO LOG-',
                                 #'WXuserslog':'-WIRES-X USERS LOG-',
                                }
ExecutedCommands = []

# Assign hierarchical values to strings that makes Scheduled Event settings_Keys sortable 
OccursHierarchy = {'every':'0','1st':'1','2nd':'2','3rd':'3','4th':'4','5th':'5'} 
DowHierarchy = {'day':'0','Sun':'1','Mon':'2','Tue':'3','Wed':'4','Thu':'5','Fri':'6','Sat':'7'}

Sched_nths=['every','1st','2nd','3rd','4th','5th']
Sched_days=['day','Sun','Mon','Tue','Wed','Thu','Fri','Sat']
Sched_hours=['00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23']
Sched_minutes=['00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19',
               '20','21','22','23','24','25','26','27','28','29','30','31','32','33','34','35','36','37','38','39',
               '40','41','42','43','44','45','46','47','48','49','50','51','52','53','54','55','56','57','58','59']
Sched_TOT_minutes=['5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26',
             '27','28','29','30','31','32','33','34','35','36','37','38','39','40','41','42','43','44','45','46','47',
             '48','49','50','51','52','53','54','55','56','57','58','59','60']

Sched_commands = [ 'none', 'Connect', 'Disconnect', 'Restart Wires-X App' ]

# Initialize Schedule default values at program startup 
Sched_nth_default=Sched_nths[0]
Sched_day_default=Sched_days[0]
Sched_hour_default=Sched_hours[0]
Sched_minute_default=Sched_minutes[0]
Sched_timezone_default = 'UTC'
Sched_description_default = ''
Sched_RoundQsoRoomConnection_default = True
Sched_AcceptCallsWhileInRoundQsoRooms_default = False
Sched_BackToRoundQsoAfterDisconnect_default = False
Sched_ReturnToRoomCheckbox_default = False
Sched_ReturnToRoomID_default = ''
Sched_UnlimitedTOT_default = False
Sched_TimeOutTimer_default='15'
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
# that ~\Documents\WIRESXA\AccHistory\WiresAccess.log
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
        with open(path.expanduser(settings['WXaccesslog']), encoding='utf-8', errors='backslashreplace', mode='r') as content_file:
            content = content_file.read()
    except Exception as e:
        outputScreen.append( 'Exception reading %s::%s' % (path.expanduser(settings['WXaccesslog']), e) )
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
            if (ord(ui.radioID[0]) & 1) == 1:
                radioName = 'Node'
            else:
                radioName = 'Room'
        else:
            # determine Radio Name from first 2 chars of ui.radioID
            try:
                radioName = RadioNameFromRadioID[ui.radioID[:2]]
            except:                # corrupt radioID -- set radioName to radioID
                radioName = '?%s?' % ui.radioID
                #debuglines.append('Corrupt radioID {%s} [%s]' % (ui.radioID, line)) K9EQ 2.5.1
            
        if ui.callsign == ui.nodeName:
            outputScreen.append( '%s %s [%s] %-7.7s %s %s' % (ui.timestamp,source,ui.radioID,radioName,ui.callsign,ui.gridsq) )
            outputHTML.append(   '%s %s [%s] %-7.7s %s %s' % (ui.timestamp,source,ui.radioID,radioName,callsign2html(ui.callsign),ui.gridsq) )
        else:
            outputScreen.append( '%s %s [%s] %-7.7s %s {%s} %s' % (ui.timestamp,source,ui.radioID,radioName,ui.callsign,ui.nodeName,ui.gridsq) )
            outputHTML.append(   '%s %s [%s] %-7.7s %s {%s} %s' % (ui.timestamp,source,ui.radioID,radioName,callsign2html(ui.callsign),callsign2html(ui.nodeName),ui.gridsq) )

    # sort outputHTML lines by the timestamp field at the beginning of each output line
    outputHTML.sort(reverse=True)

    try:
        # overwrite file if it already exists
        lhFile = open(path.expanduser(settings['WXlastheardHTML']), 'w')
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

##################### Make an Add/Update Event window #####################
def create_AddUpdate_event_window(settings, titleStr):
    sg.theme(settings['theme'])

    def TextLabel(text): return sg.Text(text+':', justification='r', size=(15,1))

    layout = [  [sg.Text(titleStr, font='Any 15')],
                [TextLabel('Occurs'),sg.Combo(Sched_nths,size=(6,6),default_value=Sched_nth_default,key='-OCCURS-')],
                [TextLabel('Week Day'),sg.Combo(Sched_days,size=(4,8),default_value=Sched_day_default,key='-DOW-')],
                [TextLabel('Hour'),sg.Combo(Sched_hours,size=(3,20),default_value=Sched_hour_default,key='-HOUR-')],
                [TextLabel('Minute'),sg.Combo(Sched_minutes,size=(3,20),default_value=Sched_minute_default,key='-MINUTE-')],
                [TextLabel('Event\'s Timezone'), sg.Combo(pytz.all_timezones, size=(20,20),default_value=Sched_timezone_default,key='-TZ-')],
                [TextLabel('Description'),sg.Input(default_text=Sched_description_default,key='-DESCRIPTION-')],
                [sg.Checkbox('Round QSO Room connection (ON:Permit/OFF:Reject)',default=Sched_RoundQsoRoomConnection_default,key='-RoundQsoRoomConnection-')],
                [sg.Checkbox('Accept calls while in Round QSO Rooms',default=Sched_AcceptCallsWhileInRoundQsoRooms_default,key='-AcceptCallsWhileInRoundQsoRooms-')],
                [sg.Checkbox('Back to Round QSO after disconnect',default=Sched_BackToRoundQsoAfterDisconnect_default,key='-BackToRoundQsoAfterDisconnect-')],
                [sg.Checkbox('Return to Room', default=Sched_ReturnToRoomCheckbox_default,key='-ReturnToRoomCheckbox-')],
                [TextLabel('Room ID'),sg.Input(default_text=Sched_ReturnToRoomID_default,key='-ReturnToRoomID-')],
                [sg.Checkbox('Unlimited TOT', default=Sched_UnlimitedTOT_default,key='-UnlimitedTOT-')],
                [TextLabel('TOT(TimeOut Timer)'),sg.Combo(Sched_TOT_minutes,size=(3,20),default_value=Sched_TimeOutTimer_default,key='-TimeOutTimer-')],
                [TextLabel('Command'),sg.Combo(Sched_commands,size=(32,32),default_value=Sched_command_default,key='-COMMAND-')],
                [TextLabel('Argument'),sg.Input(default_text=Sched_cmdarg_default,key='-CMDARG-')],
                [sg.Button('Save'), sg.Button('Cancel')]  ]

    window = sg.Window(titleStr, layout, location=tuple(settings['lastLocation']), finalize=True)
    return window

def add_event_to_schedule(settings, new_or_update, event_values):
    global Sched_nth_default
    global Sched_day_default
    global Sched_hour_default
    global Sched_minute_default
    global Sched_timezone_default 
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
    Sched_timezone_default = event_values['-TZ-']
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

    if not event_values['-TZ-'] in pytz.all_timezones:
        is_error += '%s is not a valid Timezone setting\n' % event_values['-TZ-']

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
        returnToRoomID = event_values['-ReturnToRoomID-']
        if len(returnToRoomID) == 0:
            is_error += 'Return to Room ID must be supplied\n'
        elif returnToRoomID.isdigit():
            num = int(returnToRoomID)
            if num < 20000 or num > 89999:
                is_error += 'Numeric Return to Room ID must be a valid Room number (Between 20000 and 89999 inclusive)\n'
        elif len(returnToRoomID) > 16:
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
    validKey, new_key = makeSettingsEventKey(event_values)
    if validKey is False:
        return False, f'Invalid settings_key {new_key}'
    new_event = [event_values['-OCCURS-'],
                 event_values['-DOW-'],
                 event_values['-HOUR-'],
                 event_values['-MINUTE-'],
                 event_values['-TZ-'],
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

    # if new event key already in settings, show previous event being overwritten 
    if new_or_update == 'New' and new_key in settings:
        previous_event = settings[new_key]
        ret_str = 'WARNING: Replacing [%s %s %s %s %s %s] with [%s %s %s %s %s %s]' % (
                        previous_event[0],
                        previous_event[1],
                        previous_event[2],
                        previous_event[3],
                        previous_event[4],
                        previous_event[5],
                        new_event[0],
                        new_event[1],
                        new_event[2],
                        new_event[3],
                        new_event[4],
                        new_event[5] )
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

    scheduledKeys = []
    scheduledEvents = []

    for key in settings:
        # only get settings key that has a leading '@'
        if key[0] == '@':
            scheduledKeys.append(key)

    if len(scheduledKeys) == 0:
        return ['<empty>']

    # return list of scheduled events that are sorted by their settings_key value
    sortedKeys = sorted(scheduledKeys)
    for key in sortedKeys:
        vals = settings[key]
        if len(vals[5]) > 0:
            description='('+vals[5]+')'
        else:
            description=''
        scheduledEvents.append('%5s %3s %s:%s %s %s' % (vals[0], vals[1], vals[2], vals[3], vals[4], description))

    return scheduledEvents

##################### Return date, time, DOW and nthDOW for specified timezone in a dictionary #####################
def get_timezone_date_time_dow_nth(timezone_str):
    date_time_tz = {}

    try:
        date_time_tz['yyyy'] = datetime.now(pytz.timezone(timezone_str)).strftime('%Y') #year
        date_time_tz['mm']   = datetime.now(pytz.timezone(timezone_str)).strftime('%m') #month
        date_time_tz['dd']   = datetime.now(pytz.timezone(timezone_str)).strftime('%d') #mday
        date_time_tz['HH']   = datetime.now(pytz.timezone(timezone_str)).strftime('%H') #hour
        date_time_tz['MM']   = datetime.now(pytz.timezone(timezone_str)).strftime('%M') #minute
        date_time_tz['SS']   = datetime.now(pytz.timezone(timezone_str)).strftime('%S') #second
        date_time_tz['dow']  = datetime.now(pytz.timezone(timezone_str)).strftime('%a') #Sun,Mon,Tue,Wed...
        date_time_tz['nth']  = Sched_nths[int((int(date_time_tz['dd']) + 6) / 7)]
        date_time_tz['tz']   = timezone_str
    except:
        emsg = '\nEXCEPTION OCCURED:\n'
        emsg += '  \"%s\" is not a valid Timezone designation\n' % timezone_str
        sg.popup_ok('FATAL ISSUE: \n\n' + emsg, background_color='red', text_color='white')
        sys.exit(99) # Exit program

    return date_time_tz

##################### Check if it is time to execute a scheduled Wires-X event #####################
def is_scheduled_time(settings):

    # walk through settings and only process keys that have a leading '@'
    for key in settings:
        if key[0] == '@':
            ev_data = settings[key]
            # get current hour/minute/dow/nth based on event's specified TimeZone
            dtz = get_timezone_date_time_dow_nth(ev_data[4])
            if ev_data[3] != dtz['MM']:
                continue
            if ev_data[2] != dtz['HH']:
                continue
            if ev_data[1] != 'day':     # inverse logic that skips match_dow check when ev_data[1] == 'day'
                if ev_data[1] != dtz['dow']:
                    continue
            if ev_data[0] == dtz['nth'] or ev_data[0] == 'every':
                # key's ev_data matched -- return date/time/tz and event data
                return True, ev_data, dtz

    # No key's ev_data matched
    noData = []
    return False, noData, 'none'

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
        return 'OK - Nothing to disconnect'

    return 'OK'

###############################################################
# Cause the Wires-X app to display a "File->Settings(P)" menu
# before making automation updates to the specified sub-menu.
###############################################################
def Display_File_Settings_submenu(app, submenuStr):

    try:
        xStr = 'app.WiresX.menu_select(\'File(F)->Settings(P)\')'
        app.WiresX.menu_select('File(F)->Settings(P)') # pop-up Settings window via File menu
        xStr = 'app.Settings.type_keys(\'{HOME}{DOWN}\')'
        app.Settings.type_keys('{HOME}{DOWN}')         # go to top off Tree and down one

        # Each DOWN causes the screen to update, verified by treeItem[6] containing submenuStr
        for i in range(0,12):
            xStr = 'app[\'Settings\'].children()'
            treeItem = app['Settings'].children()   # Refresh treeItem list for each loop iteration to get screen updates
            if submenuStr in str(treeItem[6]):
                #print('vvvvv----%s----vvvvv' % submenuStr)
                #app.Settings.print_control_identifiers() # on console
                #print('^^^^^----%s----^^^^^' % submenuStr)
                return 'ok'
            xStr = 'app.Settings.type_keys(\'{DOWN}\')'
            app.Settings.type_keys('{DOWN}')    # cause the next sub-menu to be displayed

        # Did not find submenuStr -- Close pop-up by typing an ESC -- return ERROR
        xStr = 'app.Settings.type_keys(\'{ESC}\')'
        app.Settings.type_keys('{ESC}')
        return 'F-P-SUBMENU \'' + submenuStr + '\' NOT FOUND'

    except:
        return 'EXCEPTION: Display_File_Settings_submenu(\'%s\') [%s]' % (submenuStr, xStr)

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
    # '-TZ-' -------------------------------> eventData[4]
    # '-DESCRIPTION-' ----------------------> eventData[5]
    # '-RoundQsoRoomConnection-' -----------> eventData[6]
    # '-AcceptCallsWhileInRoundQsoRooms-' --> eventData[7]
    # '-BackToRoundQsoAfterDisconnect-' ----> eventData[8]
    # '-ReturnToRoomCheckbox-' -------------> eventData[9]
    # '-ReturnToRoomID-' -------------------> eventData[10]
    # '-UnlimitedTOT-' ---------------------> eventData[11]
    # '-TimeOutTimer-' ---------------------> eventData[12]
    # '-COMMAND-' --------------------------> eventData[13]
    # '-CMDARG-' ---------------------------> eventData[14]

    description = '(%s)' % eventData[5]

    # Update the Wires-X 'Call settings'

    # '-RoundQsoRoomConnection-' -----------> eventData[6]
    response = SetRoundQSORoomconnectionCheckBox(app, eventData[6])
    if response != 'OK':
        return description + '::' + response

    # '-AcceptCallsWhileInRoundQsoRooms-' --> eventData[7]
    # '-BackToRoundQsoAfterDisconnect-' ----> eventData[8]
    #### [8] can only be changed if [7] is True #####
    if eventData[7] == False and eventData[8] == False:
        response = SetAcceptcallswhileinRoundQSORoomsCheckBox(app, True)
        if response != 'OK':
            return description + '::' + response
        response = SetBacktoRoundQSOafterdisconnectCheckBox(app, False)
        if response != 'OK':
            return description + '::' + response
        response = SetAcceptcallswhileinRoundQSORoomsCheckBox(app, False)
        if response != 'OK':
            return description + '::' + response
    elif eventData[7] == False and eventData[8] == True:
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

    # '-ReturnToRoomCheckbox-' -------------> eventData[9]
    # '-ReturnToRoomID-' -------------------> eventData[10]
    if eventData[9] == False:
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
        response = SetReturntoRoomID(app, eventData[10])
        if response != 'OK':
            return description + '::' + response

    # '-TimeOutTimer-' ---------------------> eventData[12]
    ## Note: The TimeOutTimer value can only be changed when UnlimitedTOT is unchecked
    response = Set_Unlimited_TOT_checkbox(app, False)
    if response != 'OK':
        return description + '::' + response
    response = Set_TOT_TimeoutTimer(app, eventData[12])
    if response != 'OK':
        return description + '::' + response

    # '-UnlimitedTOT-' ---------------------> eventData[11]
    response = Set_Unlimited_TOT_checkbox(app, eventData[11])
    if response != 'OK':
        return description + '::' + response

    # '-COMMAND-' --------------------------> eventData[13]
    command = eventData[13]
    # '-CMDARG-' ---------------------------> eventData[14]
    cmdarg = eventData[14]
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
        return ['<none>']

    return sorted(ExecutedCommands, reverse=True)

##################### Load/Save Settings File #####################
def load_settings():

    launchChangeSettings = False
    saveSettingsToFile = False

    if path.isfile(SettingsFilePathname) == False:
        # Create default settings (without any Scheduled Events) and save to file
        settings = {}
        settings['theme'] = sg.theme()
        settings['lastLocation'] = [ 0, 0 ]
        settings['localTZ'] = 'CHANGE-ME'
        settings['WXapplication'] = WIRESX_APP
        settings['WXaccesslog'] = WIRESXA_PATHNAME + R'\AccHistory\WiresAccess.log'
        settings['WXlastheardHTML'] = USER_DESKTOP + R'\Wires-X_Last_Heard.html'
        #settings['WXinfolog']  = WIRESXA_PATHNAME + R'\WXinfo.log'
        #settings['WXuserslog'] = WIRESXA_PATHNAME + R'\WXusers.log'

        launchChangeSettings = True

    else:
        try:
            with open(SettingsFilePathname, 'r') as f:
                settings = json.load(f)

        except Exception as e:
            sg.popup_ok('FATAL ISSUE:\n\n' + e, background_color='red', text_color='white')
            sys.exit(99) # Exit program

    # verify cfg file compatibility
    try:
        someWhere = settings['lastLocation']
    except:
        settings['lastLocation'] = [ 0, 0 ]
        saveSettingsToFile = True

    try:
        localTimezone = settings['localTZ']
        if not localTimezone in pytz.all_timezones:
            settings['localTZ'] = 'Change-ME'
            launchChangeSettings = True
    except:
        settings['localTZ'] = 'Change-Me'
        launchChangeSettings = True

    try:
        WxAccessLogPath = settings['WXaccesslog']
        if WxAccessLogPath[0] != '~':
            settings['WXaccesslog'] = WIRESXA_PATHNAME + R'\AccHistory\WiresAccess.log'
            saveSettingsToFile = True
    except:
        settings['WXaccesslog'] = WIRESXA_PATHNAME + R'\AccHistory\WiresAccess.log'
        saveSettingsToFile = True

    try:
        WxLastHeardHtmlPath = settings['WXlastheardHTML']
        if WxLastHeardHtmlPath[0] != '~':
            settings['WXlastheardHTML'] = USER_DESKTOP + R'\Wires-X_Last_Heard.html'
            saveSettingsToFile = True
    except:
        settings['WXlastheardHTML'] = USER_DESKTOP + R'\Wires-X_Last_Heard.html'
        saveSettingsToFile = True

    # User needs to select their local timezone
    if launchChangeSettings == True:
        event, values = create_settings_window(settings).read(close=True)
        if event == 'Save':
            # the Save button in the Settings window has been clicked
            #
            save_settings(settings, values)

    # update default with local timezone string
    global Sched_timezone_default
    Sched_timezone_default = settings['localTZ']

    # Walk through all scheduled events in settings and if TZ setting is missing:
    #  - create new entry with 'Local Timezone' added from settings
    #  - because TZ is now a component of the settings' key, queue deletion of old event entry (must be done outside loop)
    eventKeysToPop = []
    updatedEvents = {}
    for key in settings:
        # only process settings key that has a leading '@'
        if key[0] == '@':
            splitKey = key.split('-')
            if len(splitKey) == 5:
                continue # already has TZ
            elif len(splitKey) == 4:
                # queue OLDkey pop that can't be done inside for-key-in-settings loop
                eventKeysToPop.append(key)
                oldVals = settings[key]
                newVals = [ oldVals[0], # -OCCURS-
                            oldVals[1], # -DOW-
                            oldVals[2], # -HOUR-
                            oldVals[3], # -MINUTE-
                            settings['localTZ'], # -TZ-
                            oldVals[4], # -DESCRIPTION-
                            oldVals[5], # -RoundQsoRoomConnection-
                            oldVals[6], # -AcceptCallsWhileInRoundQsoRooms-
                            oldVals[7], # -BackToRoundQsoAfterDisconnect-
                            oldVals[8], # -ReturnToRoomCheckbox-
                            oldVals[9], # -ReturnToRoomID-
                            oldVals[10], # -UnlimitedTOT-
                            oldVals[11], # -TimeOutTimer-
                            oldVals[12], # -COMMAND-
                            oldVals[13], # -CMDARG-
                          ]
                newKey = '@' + DowHierarchy[newVals[1]] + \
                         '-' + OccursHierarchy[newVals[0]] + \
                         '-' + newVals[2] + \
                         '-' + newVals[3] + \
                         '-' + newVals[4]
                updatedEvents[newKey] = newVals
            else:
                print('Deleting invalid eventEntry key [{key}]')
                eventKeysToPop.append(key)

    if len(eventKeysToPop) > 0:
        for key in eventKeysToPop:
            settings.pop(key)
        saveSettingsToFile = True

    # merge updatedEvents into settings
    if len(updatedEvents) > 0:
        settings.update(updatedEvents)
        saveSettingsToFile = True

    if saveSettingsToFile == True:
        save_settings(settings, None)

    return settings

######### Save Settings to File #######################################
#
def save_settings(settings, values):
    global SettingsFilePathname

    if values:
        # Update settings with values passed in by GUI window updates
        for key in SETTINGS_KEYS_TO_ELEMENT_KEYS:  # update window with the values read from settings file
            try:
                settings[key] = values[SETTINGS_KEYS_TO_ELEMENT_KEYS[key]]
            except Exception as e:
                print(f'Problem updating settings from window values. Key = {key}')

    with open(SettingsFilePathname, 'w') as f:
        json.dump(settings, f, indent=2)

    #sg.popup_quick_message('The settings have been saved', keep_on_top=True, background_color='green', text_color='white')

##################### Make a Select Event window #####################
def create_select_event_window(settings):
    sg.theme(settings['theme'])

    def TextLabel(text): return sg.Text(text+':', justification='r', size=(15,1))

    schedEVs = get_scheduled(settings)
    numInScheduled = len(schedEVs)

    if schedEVs[0] == '<empty>':
        # create select window with only New and Cancel buttons
        layout = [  [sg.Text('Select Scheduled Event', font='Any 15')],
                    [TextLabel('No scheduled events'),sg.Combo(schedEVs,default_value=schedEVs[0],size=(80,numInScheduled),font=('Consolas',9),key='-EMPTY-')],
                    [sg.Button('New'), sg.Button('Cancel')]  ]
        window = sg.Window('Add New Scheduled Event', layout, location=tuple(settings['lastLocation']), finalize=True)
    else:
        layout = [  [sg.Text('Select Scheduled Event', font='Any 15')],
                    [TextLabel('Choose one'),sg.Combo(schedEVs,default_value=schedEVs[0],size=(80,numInScheduled),font=('Consolas',9),key='-CHOSEN EVENT-')],
                    [sg.Button('New'), sg.Button('Delete'), sg.Button('Update'), sg.Button('Cancel')]  ]
        window = sg.Window('Select Scheduled Event', layout, location=tuple(settings['lastLocation']), finalize=True)

    return window

##################### Make a settings window #####################
def create_settings_window(settings):
    sg.theme(settings['theme'])

    def TextLabel(text): return sg.Text(text+':', justification='r', size=(15,1))

    layout = [  [sg.Text('Settings', font='Any 15')],
                [TextLabel('Theme'),sg.Combo(sg.theme_list(), size=(20, 20), key='-THEME-')],
                [TextLabel('Local Timezone'),sg.Combo(pytz.all_timezones, size=(20, 20), key='-LOCAL TZ-')],
                [TextLabel('Wires-X EXE'),sg.Input(key='-WIRES-X EXE-', size=(65,1)), sg.FileBrowse(target='-WIRES-X EXE-')],
                [TextLabel('Wires-X Access Log'),sg.Input(key='-WIRES-X ACCESS LOG-', size=(65,1))], # don't Browse because loss of ~ ($HOME)
                [TextLabel('Wires-X Last Heard'),sg.Input(key='-WIRES-X LAST HEARD-', size=(65,1))], # don't Browse because loss of ~ ($HOME)
                [sg.Button('Save'), sg.Button('Cancel')]  ]

    window = sg.Window('Settings', layout, keep_on_top=True, location=tuple(settings['lastLocation']), finalize=True)
    for key in SETTINGS_KEYS_TO_ELEMENT_KEYS:   # update window with the values read from settings file
        try:
            window[SETTINGS_KEYS_TO_ELEMENT_KEYS[key]].update(value=settings[key])
        except Exception as e:
            print(f'Problem updating PySimpleGUI window from settings. Key = {key}')

    return window

##################### create Main Program Window #####################
def create_main_window(settings):
    sg.theme(settings['theme'])

    executedCmds = get_executed_commands()
    numExecutedCmds = len(executedCmds)
    if numExecutedCmds > 10:
        numExecutedCmds = 10  # Limit number of executed cmds shown in main window
    scheduledEvents = get_scheduled(settings)
    numScheduledEvents = len(scheduledEvents)
    if numScheduledEvents > 10:
        numScheduledEvents = 10  # Limit number of scheduled events shown in main window

    layout = [  [sg.T(size=(50,1),font=('Helvetica',20),justification='center',key='-DATE TIME DOW-')],
                [sg.T('Last Heard:'),sg.Listbox(['Waiting for %s to be updated'%path.expanduser(settings['WXaccesslog'])], size=(100, 10),font=('Consolas',9),key='-WIRES-X LAST HEARD-')],
                [sg.T('  Executed:'),sg.Listbox(executedCmds, size=(100,numExecutedCmds), font=('Consolas',9), key='-EXECUTED-')],
                [sg.T('  Schedule:'),sg.Listbox(get_scheduled(settings), size=(100,numScheduledEvents), font=('Consolas',9))],
                [sg.Button('Force Disconnect'), sg.Button('Scheduler'), sg.Button('Settings')]  ]

    return sg.Window('WXscheduler (v2.5.1)', layout, location=tuple(settings['lastLocation']), finalize=True, icon='WXScheduler.ico') # K9EQ 2.5.1

##################### Derive Settings Event Key from initial event values #####################
def makeSettingsEventKey( eventValues ):

    keyList = list(eventValues.keys())

    if keyList[0] == '-OCCURS-':
        settings_key = '@' + DowHierarchy[eventValues['-DOW-']] + \
                       '-' + OccursHierarchy[eventValues['-OCCURS-']] + \
                       '-' + eventValues['-HOUR-'] + \
                       '-' + eventValues['-MINUTE-'] + \
                       '-' + eventValues['-TZ-']
        return True, settings_key

    if keyList[0] == '-CHOSEN EVENT-':
        evL = eventValues['-CHOSEN EVENT-'].strip().replace(':',' ').split(' ')
        if len(evL) >= 5:
            settings_key = '@' + DowHierarchy[evL[1]] + \
                           '-' + OccursHierarchy[evL[0]] + \
                           '-' + evL[2] + \
                           '-' + evL[3] + \
                           '-' + evL[4]
            return True, settings_key
        return False, 'invalid chosen'

    if keyList[0] == '-EMPTY-':
        return False, eventValues['-EMPTY-']

    sg.popup_ok(f'makeSettingsEventKey():\n  Unexpected DictKey [ {keyList[0]} ]\n  Returning Invalid-Key\n', background_color='red', text_color='white')

    return False, eventValues['-invalid-']

##################### Program Initialization & Event Loop #####################
def main():
    # global scope necessary to write the following variables
    global Sched_nth_default
    global Sched_day_default
    global Sched_hour_default
    global Sched_minute_default
    global Sched_timezone_default 
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
    global USER_DESKTOP
    global WIRESXA_PATHNAME
    global SettingsFilePathname

    # Construct and verify folder/file pathnames
    emsg = ''

    if path.isdir(path.expanduser(USER_DESKTOP_1)) == True:
        USER_DESKTOP = USER_DESKTOP_1
    elif path.isdir(path.expanduser(USER_DESKTOP_2)) == True:
        USER_DESKTOP = USER_DESKTOP_2
    else:
        emsg += 'Could not find User Desktop folder!\n'
        emsg += '  Tried: %s\n' % path.expanduser(USER_DESKTOP_1)
        emsg += '  Tried: %s\n' % path.expanduser(USER_DESKTOP_2)

    if path.isdir(path.expanduser(WIRESXA_PATHNAME_1)) == True:
        WIRESXA_PATHNAME = WIRESXA_PATHNAME_1
    elif path.isdir(path.expanduser(WIRESXA_PATHNAME_2)) == True:
        WIRESXA_PATHNAME = WIRESXA_PATHNAME_2
    else:
        emsg += 'Could not find WIRESXA folder!\n'
        emsg += '  Tried: %s\n' % path.expanduser(WIRESXA_PATHNAME_1)
        emsg += '  Tried: %s\n' % path.expanduser(WIRESXA_PATHNAME_2)

    SettingsFilePathname = path.expanduser(WIRESXA_PATHNAME + SETTINGS_FILENAME)

    if path.isfile(WIRESX_APP) == False:
        emsg += 'Could not find Wires-X executable!\n'
        emsg += '  Tried: %s\n' % WIRESX_APP

    if len(emsg) > 0:
        emsg += '\nNOTE: This program requires:\n'
        emsg += ' - Wires-X app running on this same PC\n'
        emsg += ' - WXscheduler running under same User ID as Wires-X app\n'
        sg.popup_ok('FATAL ISSUE: \n\n' + emsg, background_color='red', text_color='white')
        return # Exit program

    window, settings = None, load_settings()

    dtL = get_timezone_date_time_dow_nth(settings['localTZ'])
    previous_minute = dtL['MM']

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
        dtL = get_timezone_date_time_dow_nth(settings['localTZ'])
        formatted_dtL = '%s/%s/%s %s:%s:%s' % (dtL['yyyy'],dtL['mm'],dtL['dd'],dtL['HH'],dtL['MM'],dtL['SS'])  # "yyyy/mm/dd HH:MM:SS"
        window['-DATE TIME DOW-'].update('%s (%s %s) %s' % (formatted_dtL, dtL['nth'], dtL['dow'], dtL['tz']))

        # check if WXaccesslog file exists
        outputLines = []
        aPath = path.expanduser(settings['WXaccesslog'])
        if path.isfile(aPath) == False:
            outputLines.append('Unable to access %s' % aPath )
            outputLines.append('[Settings] -> Wires-X Access Log: ???')
            window['-WIRES-X LAST HEARD-'].update(outputLines)
        else:
            # check if the WiresAccess.log modification time has changed
            #
            try:
                currentModTime = path.getmtime(aPath) 
            except Exception as e:
                outputLines.append(f'Exception: {e}')
                window['-WIRES-X LAST HEARD-'].update(outputLines)
            else:
                if currentModTime != previousModTime:
                    previousModTime = currentModTime
                    current_cksum, last_heard_lines = refreshLastHeard(settings, previous_cksum)
                    if current_cksum != previous_cksum:
                        previous_cksum = current_cksum
                        window['-WIRES-X LAST HEARD-'].update(last_heard_lines)

        # If entering a new minute, then check if it is time to execute a scheduled Wires-X action
        #
        current_minute = dtL['MM']
        if previous_minute != current_minute:
            is_time, ev_data, dtz = is_scheduled_time(settings)
            if is_time:
                ## Execute scheduled actions ##
                report = performWXactions(ev_data, settings)
                ExecutedCommands.append('%s/%s/%s %s:%s:%s (%s) %s' % (dtz['yyyy'], dtz['mm'], dtz['dd'], dtz['HH'], dtz['MM'], dtz['SS'], dtz['tz'], report))
                window['-EXECUTED-'].update(get_executed_commands())
            previous_minute = current_minute

        # If the main Window has been moved:
        #   - update the location in settings
        #   - save settings to cfg file
        #
        prevLocation = tuple(settings['lastLocation'])
        currLocation = window.current_location()
        if prevLocation[0] != currLocation[0] or prevLocation[1] != currLocation[1]:
            settings['lastLocation'] = currLocation
            save_settings(settings, None)

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

        if event == 'Scheduler':
            # the Scheduler button in the main window has been clicked
            #
            event, selectedSchedEvent = create_select_event_window(settings).read(close=True)
            before_validKey, before_settings_key = makeSettingsEventKey(selectedSchedEvent)
            if event == 'Cancel':
                pass    # Nothing to do
            elif event == 'New' or event == 'Update':
                new_or_update = event
                if event == 'New' :
                    # Use Sched_*_defaults as last set, except force TZ to local TZ
                    Sched_timezone_default = settings['localTZ']
                    winTitleStr = 'Add New Event to Schedule'
                elif event == 'Update':
                    # Set defaults to selected event's values
                    selected_event_values = settings[before_settings_key]
                    Sched_nth_default = selected_event_values[0] # -OCCURS-
                    Sched_day_default = selected_event_values[1] # -DOW-
                    Sched_hour_default = selected_event_values[2] # -HOUR-
                    Sched_minute_default = selected_event_values[3] # -MINUTE-
                    Sched_timezone_default = selected_event_values[4] # -TZ-
                    Sched_description_default = selected_event_values[5] # -DESCRIPTION-
                    Sched_RoundQsoRoomConnection_default = selected_event_values[6] # -RoundQsoRoomConnection-
                    Sched_AcceptCallsWhileInRoundQsoRooms_default = selected_event_values[7] # -AcceptCallsWhileInRoundQsoRooms-
                    Sched_BackToRoundQsoAfterDisconnect_default = selected_event_values[8] # -BackToRoundQsoAfterDisconnect-
                    Sched_ReturnToRoomCheckbox_default = selected_event_values[9] # -ReturnToRoomCheckbox-
                    Sched_ReturnToRoomID_default = selected_event_values[10] # -ReturnToRoomID-
                    Sched_UnlimitedTOT_default = selected_event_values[11] # -UnlimitedTOT-
                    Sched_TimeOutTimer_default = selected_event_values[12] # -TimeOutTimer-
                    Sched_command_default = selected_event_values[13] # -COMMAND-
                    Sched_cmdarg_default = selected_event_values[14] # -CMDARG-
                    winTitleStr = 'Update Event in Schedule'

                # Get New/Updated event values
                event, event_values = create_AddUpdate_event_window(settings, winTitleStr).read(close=True)

                if event == 'Save':
                    # the Save button in the Add/Update Event window has been clicked
                    #
                    after_validKey, after_settings_key = makeSettingsEventKey(event_values)
                    if new_or_update == 'Update' and before_validKey == True and after_validKey == True and before_settings_key != after_settings_key:
                        # the Update changed settings_key value(s) - therefore must delete before_settings_key (and values) from settings
                        deleted_event = settings.pop(before_settings_key)

                    ret, ew_str = add_event_to_schedule(settings, new_or_update, event_values)
                    if ret == True:
                        if len(ew_str) > 0:
                            sg.popup(ew_str, background_color='yellow', text_color='blue')
                        save_settings(settings, None)
                        window.close()
                        window = None
                    else:
                        sg.popup(ew_str, background_color='red', text_color='white')

            elif event == 'Delete':
                # the Delete button in the Select Event window has been clicked
                #   - derive settings-key from first 5 event values
                validKey, settings_key = makeSettingsEventKey(selectedSchedEvent)
                if validKey:
                    deleted_event = settings.pop(settings_key)
                    save_settings(settings, None)
                    window.close()
                    window = None

        elif event == 'Force Disconnect':
            # the Force Disconnect button in the main window has been clicked
            #
            report = ForceDisconnectRoom(settings)
            ExecutedCommands.append('%s %s' % (formatted_dtL, report))
            window['-EXECUTED-'].update(get_executed_commands())

        elif event == 'Settings':
            # the Settings button in the main window has been clicked
            #
            event, values = create_settings_window(settings).read(close=True)
            if event == 'Save':
                # the Save button in the Settings window has been clicked
                #
                window.close()
                window = None
                save_settings(settings, values)

                # update default with local timezone string
                Sched_timezone_default = settings['localTZ']

    window.close()
    return

main()
