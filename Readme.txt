WXScheduler was written by Bill, W9LBR. It has been packaged as an .exe with a Windows installer by Chris K9EQ.
  In addition, the following changes have been made:
 - 2023-01-12 2.5.1 Removed radioID error messages from the log (there were too many of them and they didn't add value)
                    Added icons.

The packaged version displays the GUI without a shell console. If you want the console as well, 
rename the file to *.py instead of *.pyw. You'll then need to open up a shell (command prompt)
and launch with a 32-bit Python interpreter as in: "C:\python32 WXSheduler.py".
The source code is available at https://github.com/K9EQ/WXScheduler.
Also see HamOperator.com for more information about WXScheduler.

Bill has done a fantastic job with this program. I hope that my minor additions make it easier for
a wider audience to benefit from it.

Bill's original text file, WXScheduler.txt, follows:

WXscheduler.pyw is a Python 3 program that automates certain aspects of Wires-X Room/Node connections according
to a user defined schedule.

Schedulable operations:
  - Settings->Call settings
      - Unlimited TOT
      - TOT(TimeOut Timer)
  - Settings->General settings
      - Round QSO Room connection
      - Accept calls while in Round QSO Rooms
      - Back to Round QSO after disconnect
      - Return to Room
      - Room ID
  - Connect <Node/Room>
  - Disconnect
  - Restart Wires-X App

Immediate operation:
  - Force Disconnect

Scheduler Settings Example (Automatically reconnects after network glitch):

  - MnWis Fusion Net every Monday 7:30pm US/Central time

     - First segment (net usually lasts at least 60 minutes)
                   Occurs: [every]
                 Week Day: [Mon]
                     Hour: [19]
                   Minute: [28]
         Event's Timezone: [US/Central]
              Description: [start MnWis Fusion Net]
       [x] RoundQSO Room connection
       [ ] Accept calls while in Round QSO Rooms
       [x] Back to Round QSO after disconnect
       [x] Return to Room
              Room ID: [MNWIS]
       [x] Unlimited TOT
       TOT(TimeOut Timer): [60]
                  Command: [Disconnect]
                 Argument: [          ]

     - Second segment (net running over 60 minutes - switch to LocFav-ROOM on timeout)
                   Occurs: [every]
                 Week Day: [Mon]
                     Hour: [20]
                   Minute: [30]
         Event's Timezone: [US/Central]
              Description: [after MnWis Fusion Net]
       [x] RoundQSO Room connection
       [x] Accept calls while in Round QSO Rooms
       [x] Back to Round QSO after disconnect
       [x] Return to Room
              Room ID: [LocFav-ROOM]
       [ ] Unlimited TOT
       TOT(TimeOut Timer): [10]
                  Command: [none]
                 Argument: [          ]

NOTE #1: WXscheduler.pyw MUST BE RUN ON THE SAME WINDOWS PC and under the SAME USER ID as
         the Wires-X PC application.

NOTE #2: When the Windows screen is LOCKED, this prevents WXscheduler's scheduled keyboard/mouse
         automation from occuring. Suggest: [Settings->System->Power&Sleep->Screen turn off after NEVER]

WXscheduler v2.5 updates:
  - Schedule settings now require a Timezone parameter.
      - Compensation for Daylight Savings Time is now automatic on a per timezone basis
      - Load Python Timezone package using Windows Command Prompt:
                  pip install pytz
  - WXscheduler.cfg is now more portable and user editable
      - User's HOME path is now resolved at run time and no longer saved in WXscheduler.cfg
      - JSON data fields are now saved in pretty mode
      - When WXscheduler-v2.5 detects an existing WXscheduler.cfg without timezone data fields,
        the user will be prompted to set their local timezone value and all existing events will
    be updated with the local timezone value.
  - WXscheduler's main window:
      - Scheduled events are displayed in chronological order
      - When the main window is moved, its new position is remembered. And all sub-windows
          will be opened in the same position. This replaces PySimpleGUI's default of centering
      window within the display.
      - [Add Event] and [Delete Event] buttons have been replaced with the [Scheduler] button that
          when clicked provides a new window with buttons: [New] [Delete] [Update] [Cancel]

WXscheduler v2.4 updates:
  - Additional exception debug information
  - Changed main window title from "Wires-X Scheduler" to "WXscheduler (v...)"

WXscheduler v2.3 update:
  At startup determines whether Microsoft OneDrive is active or not and locates
  the Documents/WIRESXA and user/Desktop folders accordingly.

WXscheduler v2.2 updates:
  At startup verifies:
    - Wires-X.exe is accessable at its standard location
    - /Users/????/Documents/WIRESXA folder is accessable (where WXscheduler.cfg is stored)

WXscheduler v2.1 updates:
  Better displays exception information
  Shows expected WIRESXA pathname to last heard data file 

WXscheduler v2.0 update adds:
  Display Last Heard information that the Wires-X application updates once a minute.
  If available, Lat/Lon displayed as 6-character Grid Square.

Wxscheduler v1.3 updates:
  Last Heard information now correctly displays the newer Yaesu models
       (i.e. FT5-D, FTM-200, and FTM-500)
  Desktop\Wires-X_Last_Heard.html contains hypertext encoded callsigns
       that perform a QRZ.com callsign lookup.

Prerequisites (versions below are what was tested on):

    Windows PC (win7 and higher)
        Wires-x App (Ver-1.550)
        Python 3 (32-bit) because Wires-X App uses the win32 UI
        site-packages: PySimpleGUI, pywinauto, pytz

Installing Python:

    Recommend going to https://www.python.org/downloads/windows/
    to find a Python release that matches Windows on your PC.

    After installing Python, use a CMD window to load site-packages:
        - pip install pywinauto PySimpleGui pytz

Installing WXscheduler.pyw:

    Copy WXscheduler.pyw to your Desktop

    Double click on the WXscheduler.pyw icon to launch the program.

    Hint: To see any startup or run time issues:
          - Open a CMD window
          - Copy WXscheduler.pyw WXscheduler.py
          - WXscheduler.py

Enjoy,

Bill
w9lbr@arrl.net
