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
Sched_days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat']
Sched_day_default=Sched_days[0]
Sched_hours=['00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23']
Sched_hour_default=Sched_hours[0]
Sched_minutes=['00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19',
               '20','21','22','23','24','25','26','27','28','29','30','31','32','33','34','35','36','37','38','39',
               '40','41','42','43','44','45','46','47','48','49','50','51','52','53','54','55','56','57','58','59']
Sched_minute_default=Sched_minutes[0]
Sched_commands = [ 'Connect', 'Disconnect', 'SetUnlimitedTOT', 'SetTimeoutTOT','Restart Wires-X App' ]
Sched_command_default=Sched_commands[0]
Sched_argument_default = ''
Sched_remarks_default = ''

###################################################################
# Create a dictionary indexed by first 2 characters of a radio-ID #
###################################################################
RadioNameFromRadioID = {
    'E0' : 'FT1-D',
    'E5' : 'FT2D',
    'EA' : 'FT3-D',
    'F0' : 'FTM-400D',
    'F5' : 'FTM-100D',
    'FA' : 'FTM-300D',
    'G0' : 'FT-991',
    'H0' : 'FTM-3200D',
    'H5' : 'FT-70D',
    'HA' : 'FTM-3207D',
    'HF' : 'FTM-7250D',
    'R0' : 'DR-1X',
    'R5' : 'DR-2X',
    }

#######################################################################
# Convert Degrees Minutes Seconds to decimal maidenhead.
#######################################################################
def dms2dd(degrees, minutes, seconds, direction):
    dd = float(degrees) + float(minutes)/60 + float(seconds)/(60*60);
    if direction == 'S' or direction == 'W':
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
    foundAll = False
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
            foundAll = True
            break
        else:
            i += 1

    if foundAll == False:
        #print('PROBLEM: convertDegMinSec_to_GridSquare(%s)' % dmsStr)
        return 'xxxxxx'

    # Convert DMS to decimal maidenhead as a positive value
    LatDecimal = dms2dd(LatDegrees, LatMinutes, LatSeconds, '+')
    LonDecimal = dms2dd(LonDegrees, LonMinutes, LonSeconds, '+')

    # Adjust decimal values for Grid Square conversion
    if LatHemisphere == 'N':
        LatDecimal = 90.0 + LatDecimal
    else:
        LatDecimal = 90.0 - LatDecimal

    if LonHemisphere == 'E':
        LonDecimal = 180.0 + LonDecimal
    else:
        LonDecimal = 180.0 - LonDecimal

    # Convert latitude/longitude decimal values to grid square characters
    LatField = chr(int(LatDecimal / 10) + ord('A'))
    LatGrid = str(int(LatDecimal % 10))
    subsq = (LatDecimal - int(LatDecimal)) * 60
    LatSubsq = chr(int(subsq / 2.5) + ord('a'))

    LonField = chr(int(LonDecimal / 20) + ord('A'))
    LonGrid = str((int((LonDecimal / 2) % 10)))
    subsq = ((LonDecimal - int(LonDecimal)) * 120) + 2.5
    LonSubsq = chr(int(subsq / 5) + 1 + ord('a'))

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
    try:
        with open(settings['WXaccesslog'], 'r') as content_file:
            content = content_file.read()
    except Exception as e:
        # Append lines in reverse order because the output is viewed last line first
        outputScreen.append( 'Exception: %s' % e )
        outputScreen.append( 'Problem reading %s' % settings['WXaccesslog'] )

    else:
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
                outputScreen.append( '%s %s [%s] %-9.9s %s %s' % (ui.timestamp,source,ui.radioID,radioName,ui.callsign,ui.gridsq) )
                outputHTML.append( '%s %s [%s] %-9.9s %s %s' % (ui.timestamp,source,ui.radioID,radioName,callsign2html(ui.callsign),ui.gridsq) )
            else:
                outputScreen.append( '%s %s [%s] %-9.9s %s {%s} %s' % (ui.timestamp,source,ui.radioID,radioName,ui.callsign,ui.nodeName,ui.gridsq) )
                outputHTML.append( '%s %s [%s] %-9.9s %s {%s} %s' % (ui.timestamp,source,ui.radioID,radioName,callsign2html(ui.callsign),callsign2html(ui.nodeName),ui.gridsq) )

        # sort the output lines by the timestamp field at the beginning of each output line
        outputScreen.sort(reverse=True)
        outputHTML.sort()

    try:
        # overwrite file if it already exists
        lhFile = open(settings['WXlastheard'], 'w')
        for line in outputHTML:
            lhFile.write(line + '\n')
        lhFile.close()

    except Exception as e:
        outputScreen.append(f'LastHeardLog: {e}')

    if len(debuglines) > 0:
        for line in debuglines:
            print('debug:',line)

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
                [TextLabel('Occurs'),sg.Combo(Sched_nths,size=(20,20),default_value=Sched_nth_default,key='-OCCURS-')],
                [TextLabel('Week Day'),sg.Combo(Sched_days,size=(20,20),default_value=Sched_day_default,key='-DOW-')],
                [TextLabel('Hour'),sg.Combo(Sched_hours,size=(20,20),default_value=Sched_hour_default,key='-HOUR-')],
                [TextLabel('Minute'),sg.Combo(Sched_minutes,size=(20,20),default_value=Sched_minute_default,key='-MINUTE-')],
                [TextLabel('Command'),sg.Combo(Sched_commands,size=(20,20),default_value=Sched_command_default,key='-COMMAND-')],
                [TextLabel('Argument'),sg.Input(default_text=Sched_argument_default,key='-ARGUMENT-')],
                [TextLabel('Remarks'),sg.Input(default_text=Sched_remarks_default,key='-REMARKS-')],
                [sg.Button('Update'), sg.Button('Cancel')]  ]

    window = sg.Window('Add Event to Schedule', layout, finalize=True)

    return window

def add_event_to_schedule(settings, event_values):
    global Sched_nth_default
    global Sched_day_default
    global Sched_hour_default
    global Sched_minute_default
    global Sched_command_default
    global Sched_argument_default
    global Sched_remarks_default

    # update defaults to minimize re-entering Add Event values next time
    Sched_nth_default      = event_values['-OCCURS-']
    Sched_day_default      = event_values['-DOW-']
    Sched_hour_default     = event_values['-HOUR-']
    Sched_minute_default   = event_values['-MINUTE-']
    Sched_command_default  = event_values['-COMMAND-']
    Sched_argument_default = event_values['-ARGUMENT-']
    Sched_remarks_default  = event_values['-REMARKS-']

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

    # Validate argument or blank out based on command
    if command == 'Connect':
        if_error = 'Argument must be a valid Node or Room number (Between 10000 and 99999 inclusive)\n'
        try:
            num = int(event_values['-ARGUMENT-'])
            if num < 10000 or num > 99999:
                is_error += if_error
        except:
            is_error += if_error

    elif command == 'SetTimeoutTOT':
        if_error = 'Argument must be a valid number between 5 and 60 inclusive)\n'
        try:
            num = int(event_values['-ARGUMENT-'])
            if num < 5 or num > 60:
                is_error += if_error
        except:
            is_error += if_error

    else:
        event_values['-ARGUMENT-'] = ''
        Sched_argument_default = ''
    
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
                 event_values['-COMMAND-'],
                 event_values['-ARGUMENT-'],
                 event_values['-REMARKS-'],
                ]

    # if event key already in settings, save previous event 
    if new_key in settings:
        previous_event = settings[new_key]
        ret_str = 'WARNING: Replaced [%s %s %s %s %s %s %s] with [%s %s %s %s %s %s %s]' % (
                        previous_event[0],
                        previous_event[1],
                        previous_event[2],
                        previous_event[3],
                        previous_event[4],
                        previous_event[5],
                        previous_event[6],
                        new_event[0],
                        new_event[1],
                        new_event[2],
                        new_event[3],
                        new_event[4],
                        new_event[5],
                        new_event[6] )
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
            if len(vals[6]) > 0: remarks='('+vals[6]+')'
            else: remarks=''
            scheduled.append('%5s %3s %s:%s %s %s %s' % (vals[0], vals[1], vals[2], vals[3], vals[4], vals[5], remarks))

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
            if ev_data[3] != match_minute: continue
            if ev_data[2] != match_hour: continue
            if ev_data[1] != match_dow: continue
            if ev_data[0] == match_nth or ev_data[0] == 'every':
                command  = ev_data[4]
                argument = ev_data[5]
                if len(ev_data[6]) > 0:
                    remark='('+ev_data[6]+')'
                else:
                    remark=''
                return True, command, argument, remark

    return False, '', '', ''

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
        exstr = 'app.WiresX.menu_select(Connect(C)->Disconnect(D))'
        app.WiresX.menu_select('Connect(C)->Disconnect(D)')
    except:
        return exstr

    return 'OK'

###############################################################
# Cause the Wires-X app to display a "File->Settings(P)" menu
# before making automation updates to the specified sub-menu.
###############################################################
def Display_File_Settings_submenu(app, submenuStr):

    try:
        exstr = 'File(F)->Settings(P) #1'
        app.WiresX.menu_select('File(F)->Settings(P)') # pop-up Settings window via File menu
    except:
        # Settings window already open or Wires-X app minimized?
        try:
            exstr = 'app.Settings.type_keys({ESC})'
            app.Settings.type_keys('{ESC}')                     # should close any sub windows
            exstr = 'File(F)->Settings(P) #2'
            app.WiresX.menu_select('File(F)->Settings(P)') # pop-up Settings window via File menu
        except:
            return exstr
            
    try:
        # go to top off Tree and down 8 times to 'General Settings'
        exstr = 'type_keys(\'{HOME}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}\')'
        app.Settings.type_keys('{HOME}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}')

    except:
        return 'Display_File_Settings_submenu() EXCEPTION: ' + exstr

    return 'ok'

###########################################################################
# Automate application setting of Unlimited TOT checkbox in Settings menu
###########################################################################
def Set_Unlimited_TOT_checkbox(app):

    response = Display_File_Settings_submenu(app, 'General settings')
    if response != 'ok':
        return response

    try:
        exstr = 'Settings.Checkbox.check()'
        app.Settings.Checkbox.check()                  # check the TOT box
        exstr = 'Settings.OK.click()'
        app.Settings.OK.click()                        # click OK

    except:
        return 'Set_Unlimited_TOT_checkbox() EXCEPTION: ' + exstr

    return 'OK'

###########################################################################
# Automate application setting of TOT(TimeOut Timer) value in Settings menu
# Note: This routine always unchecks the Unlimited TOT checkbox first
###########################################################################
def Set_TOT_TimeoutTimer(app, minutes):

    response = Display_File_Settings_submenu(app, 'General settings')
    if response != 'ok':
        return response

    try:
        exstr = 'Settings.Checkbox.uncheck()'
        app.Settings.Checkbox.uncheck()                # uncheck the TOT box
        exstr = 'Settings.OK.click() #1'
        app.Settings.OK.click()                        # click OK
    except:
        return 'Set_TOT_TimeoutTimer() EXCEPTION: ' + exstr

    response = Display_File_Settings_submenu(app, 'General settings')
    if response != 'ok':
        return response

    try:
        exstr = 'Settings.Edit.set_edit_text(%s)' % minutes
        app.Settings.Edit.set_edit_text(minutes)       # enter minutes into text box
        exstr = 'Settings.OK.click() #2'
        app.Settings.OK.click()                        # click OK

    except:
        return 'Set_TOT_TimeoutTimer() EXCEPTION: ' + exstr

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

############## execute Wires-X command #################
def executeWXcommand(command, argument, remark, settings):

    try:
        if_exception = 'Application().connect() EXCEPTION: '
        app = Application(backend='win32').connect(path=settings['WXapplication'])
        # restore WIRES-X window only if minimized
        if_exception = 'app.window() EXCEPTION: '
        window = app.window(title=u'WIRES-X', visible_only=False).restore()
    except Exception as e:
        return if_exception + str(e)

    if command == 'Connect':
        # FYI--The App ignores Connecting to the same room that is already connected to.
        response = ConnectToRoom(app, argument)

    elif command == 'Disconnect':
        response = DisconnectFromAnyRoom(app)

    elif command == 'SetUnlimitedTOT':
        response = Set_Unlimited_TOT_checkbox(app)

    elif command == 'SetTimeoutTOT':
        response = Set_TOT_TimeoutTimer(app, argument)

    elif command == 'Restart Wires-X App':
        response = ExitApplication(app)

    else:
        return 'UNEXPECTED command [%s] not executed' % command

    if response == 'OK':
        return '%s %s %s' % (command, argument, remark)

    return 'EXCEPTION during: ' + response

##################### Return a list of reverse sorted executed scheduled commands #####################
def get_executed_commands():

    if len(ExecutedCommands) <= 0:
        return [' ']

    return sorted(ExecutedCommands, reverse=True)

SETTINGS_FILE = path.expanduser('~/Documents/WIRESXA/WXscheduler.cfg')
SETTINGS_KEYS_TO_ELEMENT_KEYS = {'theme':'-THEME-',
                                 'WXapplication':'-WIRES-X EXE-',
                                 'WXaccesslog':'-WIRES-X ACCESS LOG-',
                                 'WXlastheard':'-WIRES-X LAST HEARD-'}

##################### Load/Save Settings File #####################
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
        settings['WXlastheard']= path.expanduser('~/Documents/WIRESXA/WX_last_heard.txt')
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

    return sg.Window('Wires-X Scheduler', layout, finalize=True)


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
        formatted_dt = '%s/%s/%s %s:%s:%s (%s %s)' % (dt['yyyy'],dt['mm'],dt['dd'],dt['HH'],dt['MM'],dt['SS'],dt['nth'],dt['dow']) # "yyyy/mm/dd HH:MM:SS (1st Mon)"
        window['-DATE TIME DOW-'].update(formatted_dt)

        # check if the WiresAccess.log modification time has changed
        #
        currentModTime = path.getmtime(settings['WXaccesslog']) 
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
            is_time, command, argument, remark = is_scheduled_time(event_to_match, settings)
            if is_time:
                report = executeWXcommand(command, argument, remark, settings)
                ExecutedCommands.append('%s: %s' % (formatted_dt, report))
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
            report = executeWXcommand('Disconnect', 'Room/Node', 'Force Disconnect button', settings)
            ExecutedCommands.append('%s: %s' % (formatted_dt, report))
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
