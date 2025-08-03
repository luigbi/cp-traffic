Attribute VB_Name = "EngrConst"

'
' Release: 1.0
'
' Description:
'   This file contains the Constants

Option Explicit

'Values must match index values on Job screen
Public Const RESOURCEJOB = 0
Public Const LIBRARYJOB = 1
Public Const TEMPLATEJOB = 2
Public Const SCHEDULEJOB = 3

'Values must match index values on list screen
Public Const AUTOLIST = 0
Public Const BUSLIST = 1
Public Const EVENTTYPELIST = 2
Public Const TIMETYPELIST = 3
Public Const MATERIALTYPELIST = 4
Public Const AUDIOLIST = 5
Public Const RELAYLIST = 6
Public Const FOLLOWLIST = 7
Public Const NETCUELIST = 8
Public Const COMMENTLIST = 9
Public Const SILENCELIST = 10
Public Const SITELIST = 11
Public Const USERLIST = 12
Public Const AUDIOTYPELIST = 13
Public Const AUDIONAMELIST = 14
Public Const BUSGROUPLIST = 15
Public Const AUDIOCONTROLLIST = 16
Public Const BUSCONTROLLIST = 17


Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400

'Alert Menu Item
Public Const MF_BITMAP = &H4&
Public Const MF_BYPOSITION = &H400&
Public Const MF_ENABLED = &H0
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

'List

Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_SELITEMRANGE = &H19B
Public Const LB_GETITEMHEIGHT = &H1A1
Public Const CB_FINDSTRING = &H14C

'List View
Public Const LV_GRIDLINES = 1
Public Const LV_SETEXTENDEDLISTVIEWSTYLE = 4150
Public Const LV_FULLROWSSELECT = 32

'Display setting
Public Const BITSPIXEL As Long = 12
Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10

'Keys
Public Const SHIFTMASK = 1
Public Const CTRLMASK = 2
Public Const ALTMASK = 4
Public Const LEFTBUTTON = 1
Public Const RIGHTBUTTON = 2
Public Const KEYBACKSPACE = 8   'Back space key pressed
Public Const KEYDECPOINT = 46   'Decimal point pressed
Public Const KEYFORWARDSLASH = 47
Public Const KEY0 = 48          '0 key pressed
Public Const KEY9 = 57          '9 key pressed
Public Const KEYLEFT = &H25
Public Const KEYUP = &H26
Public Const KEYRIGHT = &H27
Public Const KEYDOWN = &H28
Public Const KEYESC = 27

'Grid
Public Const GRIDSCROLLWIDTH = 270
Public Const GRIDSCROLLHEIGHT = 270

'Colors
'Color &HBBGGRR
Public Const LIGHTYELLOW = &HC0FFFF '&HBFFFFF '&H80FFFF '&HBFFFFF
Public Const BURGUNDY = &H80&
Public Const DARKGREEN = &H8000&
Public Const LIGHTGREEN = &HD9FFD9  '&HB0FFDF  '&H80FF80
Public Const LIGHTBLUE = &HFDFFD7
Public Const LIGHTRED = &HE7CEFF
Public Const GRAY = &HC0C0C0


'Report
Public Const MATTYPE_RPT = 1
Public Const RELAY_RPT = 2
Public Const USER_RPT = 3
Public Const SILENCE_RPT = 4
Public Const FOLLOW_RPT = 5
Public Const TIMETYPE_RPT = 6
Public Const AUDIONAME_RPT = 7
Public Const AUDIOTYPE_RPT = 8
Public Const AUDIOSOURCE_RPT = 9
Public Const SITE_RPT = 10
Public Const BUSGROUP_RPT = 11
Public Const BUS_RPT = 12
Public Const NETCUE_RPT = 13
Public Const CONTROL_RPT = 14
Public Const COMMENT_RPT = 15
Public Const EVENT_RPT = 16
Public Const ACTIVITY_RPT = 17
Public Const AUTOMATION_RPT = 18
Public Const LIBRARY_RPT = 19
Public Const LIBRARYEVENT_RPT = 20
Public Const AUDIOINUSE_RPT = 21
Public Const ITEMIDCHK_RPT = 22
Public Const SCHED_RPT = 23
Public Const TEXT_RPT = 24
Public Const TEMPLATE_RPT = 25
Public Const TEMPLATEEVENT_RPT = 26
Public Const TEMPLATEAIR_RPT = 27
Public Const ASAIRCOMPARE_RPT = 28

'Module return status
'Call return status
Public Const CALLDONE = 1000 'User press Done
Public Const CALLCANCELLED = 1001 'User pressed Cancel
Public Const CALLTERMINATED = 1002 'Error

'send message by number API
Public Const LB_SETHORIZONTALEXTENT = &H194

