Attribute VB_Name = "CONSTDEF"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CONSTDEF.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Constants (Marked)                                                              *
'*  BYPASSHIPRI                   BYPASSDEMOGRAPHIC             BYPASSDATENEEDED          *
'*  AUDIOVAULTAIR                 AVAILSCOMBO                   COPY_REGIONS              *
'*  ONE_MINUTE                                                                            *
'******************************************************************************************


        ' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
'
' Release: 1.0
'
' Description:
'   This file contains the Constant used by the traffic system

'Global Const FALSE = 0
'Global Const TRUE = -1

Global Const YES = True   'Same as True
Global Const NO = False     'Same as False

Public Const MAX_PATH = 260

' Show parameters
'Replaced with vbModal: Global Const MODAL = 1

'Window (form) size
'Global Const WNDNORMALSIZE = 0  'Normal size
'Global Const WNDMINSIZE = 1     'Minimized (shrunk to an icon)

' MsgBox parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons

Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message

Global Const MB_DEFBUTTON2 = 256       ' Second button is default

Global Const MB_APPLMODAL = 0

' Alignment (label)

'Color &HBBGGRR
' BackColor, ForeColor, FillColor (standard RGB colors: form, controls)
Global Const BLACK = &H0&
Global Const Red = &HFF&
Global Const GREEN = &HFF00&
Global Const Yellow = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF
Global Const BISQUE = &HC4E4FF
Global Const PINK = 13353215 'TTP 10549 - Learfield Cloud printing 911, Crystal Crashes, Use Adobe
Global Const DARKRED = 3942528  'TTP 10585 - Event and Sports Avails export: if team name or event is set to "cancelled", show row in red font color

' Extra background colors
Global Const DARKGREEN = &H8000&
Global Const DARKPURPLE = &H800080
Global Const DARKYELLOW = &H8080&
Global Const DARKGRAY = &H808080
Global Const GRAY = &HC0C0C0
Global Const LIGHTGRAY = &HE0E0E0
Global Const LIGHTYELLOW = &HBFFFFF '&H80FFFF '&HBFFFFF
Global Const LIGHTGREEN = &H80FF80
Global Const LIGHTBLUE = &HFDFFD7
Public Const SPLITBLUE = &HFFFFC0   'Slightly darker blue then LIGHTBLUE
Public Const MIDDLEBLUE = &HCC6600  '&HFF9933  'Used with hidden lines only: CC6600 is the darker blue color that jim is verifing that Entravision is ok with
Public Const Orange = 33023 '&H80FF
Public Const LIGHTRED = &HCCCCFF
Public Const LIGHTERGREEN = &HCCFFE5
Public Const LIGHTERYELLOW = &HCCFFFF
Public Const LIGHTERBISQUE = &HCCE5FF
Public Const YELLOWORANGE = &H56D5F8    '&H66CCFF
Public Const LIGHTMAGENTA = &HFFBEFF
'Non-edit field color

' System Colors

'Window Position setting
'Global Const SWP_NOACTIVATE = &H10
'Global Const SWP_SHOWWINDOW = &H40
'Global Const SWP_NOMOVE = &H2
'Global Const SWP_NOSIZE = &H1

'Display setting
Public Const BITSPIXEL As Long = 12
Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10


' MousePointer (form, controls)
'Replaced with vbDefault: Global Const DEFAULT = 0           ' 0 - Default
'Global Const ARROW = 1             ' 1 - Arrow
'Global Const CROSSHAIR = 2         ' 2 - Cross
'Global Const IBEAM = 3             ' 3 - I-Beam
'Global Const ICONPOINTER = 4       ' 4 - Icon
'Global Const SIZEPOINTER = 5       ' 5 - Size
'Global Const SIZENESW = 6          ' 6 - Size NE SW
'Global Const SIZENS = 7            ' 7 - Size N S
'Global Const SIZENWSE = 8          ' 8 - Size NW SE
'Global Const SIZEWE = 9            ' 9 - Size W E
'Global Const UPARROW = 10          ' 10 - Up Arrow
'Replaced with vbHourGlass: Global Const xHOURGLASS = 11        ' 11 - Hourglass

'Drag
Global Const DRAGENTER = 0

'Number key filters
Global Const KEYBACKSPACE = 8   'Back space key pressed
Global Const KEYPOS = 43        '+ key pressed
Global Const KEYNEG = 45        '- key pressed
Global Const KEYDECPOINT = 46   'Decimal point pressed
Global Const KEY0 = 48          '0 key pressed
Global Const KEY5 = 53
Global Const KEY9 = 57          '9 key pressed
Global Const KEYUA = 65         'Upper case A
Global Const KEYUZ = 90         'Upper case Z
Global Const KEYLA = 97         'Lower case a
Global Const KEYLZ = 122        'Lower case z
Global Const KEYASTERISK = 42   'Asterisk key pressed
Global Const KEYSLASH = 47      '/ Slash key
Global Const KEYCOMMA = 44

'Example of number filter code:
'    ilPos = InStr(edcRate.Text, ".")
'    If ilPos > 0 Then
'        If KeyAscii = KEYDECPOINT Then
'            Beep
'            KeyAscii = 0
'            Exit Sub
'        End If
'    End If
'    If edcRate.SelStart = 0 Then
'        If (KeyAscii = KEYBACKSPACE) Or (KeyAscii = KEYNEG) Or ((KeyAscii >= KEY0) And (KeyAscii <= KEY9)) Or (KeyAscii = KEYDECPOINT) Then
'        Else
'            Beep
'            KeyAscii = 0
'            Exit Sub
'        End If
'    Else
'        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
'            Beep
'            KeyAscii = 0
'            Exit Sub
'        End If
'    End If


' Button and Shift (KeyDown, KeyUp, MouseDown, MouseMove, MouseUp)
Global Const SHIFTMASK = 1
Global Const CTRLMASK = 2
Global Const ALTMASK = 4
'Global Const MIDDLEBUTTON = 4

' KeyCode (KeyDown, KeyUp)
'Global Const KEYMBUTTON = &H4         ' NOT contiguous with L & RBUTTON
'Global Const KEYCLEAR = &HC
'Global Const KEYRETURN = &HD
'Global Const KEYPAUSE = &H13
'Global Const KEYCAPITAL = &H14
'Global Const KEYPRIOR = &H21
'Global Const KEYNEXT = &H22
'Global Const KEYEND = &H23
'Global Const KEYHOME = &H24
Global Const KEYLEFT = &H25
Global Const KEYUP = &H26
Global Const KEYRIGHT = &H27
Global Const KeyDown = &H28
'Global Const KEYSELECT = &H29
'Global Const KEYPRINT = &H2A
'Global Const KEYEXECUTE = &H2B
'Global Const KEYSNAPSHOT = &H2C
Global Const KEYINSERT = &H2D
Global Const KEYDELETE = &H2E

' KEY_A thru KEY_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' KEY_0 thru KEY_9 are the same as their ASCII equivalents: '0' thru '9'

'Global Const KEYNUMPAD0 = &H60
'Global Const KEYNUMPAD1 = &H61
'Global Const KEYNUMPAD2 = &H62
'Global Const KEYNUMPAD3 = &H63
'Global Const KEYNUMPAD4 = &H64
'Global Const KEYNUMPAD5 = &H65
'Global Const KEYNUMPAD6 = &H66
'Global Const KEYNUMPAD7 = &H67
'Global Const KEYNUMPAD8 = &H68
'Global Const KEYNUMPAD9 = &H69
'Global Const KEYMULTIPLY = &H6A
'Global Const KEYADD = &H6B
'Global Const KEYSEPARATOR = &H6C
'Global Const KEYSUBTRACT = &H6D
'Global Const KEYDECIMAL = &H6E
'Global Const KEYDIVIDE = &H6F
Global Const KEYF1 = &H70
Global Const KEYF2 = &H71
Global Const KEYF3 = &H72
Global Const KEYF4 = &H73
Global Const KEYF5 = &H74
'Global Const KEYF6 = &H75
'Global Const KEYF7 = &H76
'Global Const KEYF8 = &H77
'Global Const KEYF9 = &H78
'Global Const KEYF10 = &H79
'Global Const KEYF11 = &H7A
'Global Const KEYF12 = &H7B
'Global Const KEYF13 = &H7C
'Global Const KEYF14 = &H7D
'Global Const KEYF15 = &H7E
'Global Const KEYF16 = &H7F
'Global Const KEYNUMLOCK = &H90

' MsgBox parameters
'Replaced with vbOkOnly: Global Const MBx_OK = 0                 ' OK button only
'Replaced with vbOkCancel: Global Const MBx_OKCANCEL = 1           ' OK and Cancel buttons
'Replaced with vbYesNoCancel: Global Const MBx_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
'Replaced with vbYesNo: Global Const MBx_YESNO = 4              ' Yes and No buttons

'Replaced with vbCritical: Global Const MBx_ICONSTOP = 16          ' Critical message
'Replaced with vbQuestion: Global Const MBx_ICONQUESTION = 32      ' Warning query
'Replaced witg vbExclamation: Global Const MBx_ICONEXCLAMATION = 48   ' Warning message
'Replaced with vbInformation: Global Const MBx_ICONINFORMATION = 64   ' Information message

'Replaced with vbDefaultButton2: Global Const MBx_DEFBUTTON2 = 256       ' Second button is default

'Replaced with vbApplicationModal: Global Const MBx_APPLMODAL = 0

' MsgBox return values
'Replaced with vbOk: Global Const IDx_OK = 1                  ' OK button pressed
'Replaced with vbCancel: Global Const IDx_CANCEL = 2              ' Cancel button pressed
'Replaced with vbYes: Global Const IDx_YES = 6                 ' Yes button pressed
'Replaced with vbNo: Global Const IDx_NO = 7                  ' No button pressed

'Common Dialog Control
'Action Property
Global Const DLG_FILE_SAVE = 2
'Global Const DLG_COLOR = 3
'Global Const DLG_FONT = 4
Global Const DLG_PRINT = 5

'File Open/Save Dialog Flags
'Replaced with cdlOFNOverwritePrompt: Global Const OFNx_OVERWRITEPROMPT = &H2&
'Replaced with cdlOFNHideReadOnly: Global Const OFNx_HIDEREADONLY = &H4&
'Replaced with cdlOFNNoChangeDir: Global Const OFNx_NOCHANGEDIR = &H8&
'Global Const OFN_NOVALIDATE = &H100&
'Global Const OFN_EXTENSIONDIFFERENT = &H400&
'Replaced with cdlOFNPathMustExist: Global Const OFNx_PATHMUSTEXIST = &H800&
'Global Const OFN_FILEMUSTEXIST = &H1000&
'Replaced with cdlOFNCreatePrompt: Global Const OFNx_CREATEPROMPT = &H2000&
'Global Const OFN_SHAREAWARE = &H4000&
'Replaceed with CdlOFNNoReadOnlyReturn: Global Const OFNx_NOREADONLYRETURN = &H8000&

'Printer Dialog Flags
'Global Const PD_ALLPAGES = &H0&
'Global Const PD_SELECTION = &H1&
'Global Const PD_PAGENUMS = &H2&
'Global Const PD_NOSELECTION = &H4&
'Global Const PD_NOPAGENUMS = &H8&
'Global Const PD_COLLATE = &H10&
'Global Const PD_PRINTTOFILE = &H20&
'Replaced with cdlPDPrintSetup: Global Const PDx_PRINTSETUP = &H40&
'Global Const PD_NOWARNING = &H80&
'Global Const PD_RETURNDC = &H100&
'Global Const PD_RETURNIC = &H200&
'Global Const PD_RETURNDEFAULT = &H400&
'Global Const PD_SHOWHELP = &H800&
'Global Const PD_USEDEVMODECOPIES = &H40000
'Global Const PD_DISABLEPRINTTOFILE = &H80000
'Global Const PD_HIDEPRINTTOFILE = &H100000

'Replaced with vbCFText: Global Const CFx_TEXT = 1        'Text format

'  Edit Control Messages
Global Const LB_SELITEMRANGE = &H19B
Global Const EM_SETREADONLY = &HCF

'ListBox
Global Const LB_FINDSTRING = &H18F
'Global Const LB_SELITEMRANGE = &H19B
Global Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_GETITEMHEIGHT = &H1A1
'ComboBox
Global Const CB_FINDSTRING = &H14C

'List View
Public Const LV_GRIDLINES = 1
Global Const LV_SETEXTENDEDLISTVIEWSTYLE = 4150
Global Const LV_FULLROWSSELECT = 32

'Help Constants
'Global Const HELP_CONTEXT = &H1           'Display topic in ulTopic
'Global Const HELP_QUIT = &H2              'Terminate help
'Replaced with cdlHelpIndex: Global Const HELPx_INDEX = &H3             'Display index
'Global Const HELP_CONTENTS = &H3
'Replaced with cdlHelpHelpOnHelp: Global Const HELPx_HELPONHELP = &H4        'Display help on using help
'Global Const HELP_SETINDEX = &H5          'Set the current Index for multi index help
'Global Const HELP_SETCONTENTS = &H5
'Global Const HELP_CONTEXTPOPUP = &H8
'Global Const HELP_FORCEFILE = &H9
'Global Const HELP_KEY = &H101             'Display topic for keyword in offabData
'Global Const HELP_COMMAND = &H102
'Replaced with cdlHelpPartialKey: Global Const HELPx_PARTIALKEY = &H105      'call the search engine in winhelp


'
Global Const STILL_ACTIVE = &H103
Global Const PROCESS_QUERY_INFORMATION = &H400

'Grid constants
'Grid
Public Const GRIDSCROLLWIDTH = 270
Public Const GRIDSCROLLHEIGHT = 270


' Btrieve Error message abort level

' Event Types

'Spot type

'Event Class

'Library types
Global Const ALLLIB = 0
Global Const LATESTLIB = 1
Global Const ALLLIBFRONT = 2

Global Const ADJSSFPASBZ = -1 'Used to adjust all indexies tPas from 1-1200 to 0-1199

'Spot Summary flags
Global Const SSLOCK = &H40 'Avail Locked
Global Const SSLOCKSPOT = &H80    'Spots locked, Used in SSF avail structure and in Event Information structure (Spots.Frm)
Global Const SSSUSTAINING = &H100   'Used in LLC and SSF avail structures
Global Const SSSPONSORSHIP = &H200  'Only used with LLC structure- Not used because Daypart is defined as In or Out
Global Const SSLOCALONLY = &H400    'Only used with LLC structure
Global Const SSFEEDONLY = &H800     'Only used with LLC structure
Global Const SSXMID = &H1000        'across midnight (0=No; 1=Yes)
Global Const sSOverBook = &H2000    'Overbooked

Global Const SSOPENBB = &H10        'Used in SSF Spot structure
Global Const SSCLOSEBB = &H20       'Used in SSF Spot structure
Global Const SSSPLITPRI = &H40
Global Const SSSPLITSEC = &H80
Global Const SSDONUT = &H100        'Used in SSF Spot structure
Global Const SSBOOKEND = &H200      'Used in SSF Spot structure
Global Const SSAVAILBUY = &H400     'Used in SSF Spot structure
Global Const SSPREEMPTIBLE = &H800  'Used in SSF Spot structure
Global Const SSEXAVAILBUY = &H1000  'Used in SSF Spot structure
Global Const SSLIBBUY = &H2000      'Used in SSF Spot structure
Global Const SSEXCLUSIONS = &H4000  'Used in SSF Spot structure

Global Const SSSOLOAVAIL = &H20000000
Global Const SS1STPOSITION = &H40000000

'Contract Header changes
Global Const HDADVTCHG = &H1
Global Const HDCOMP1CHG = &H2
Global Const HDCOMP2CHG = &H4
Global Const HDEXCL1CHG = &H8
Global Const HDEXCL2CHG = &H10
Global Const HDSCHSTATUSPCHG = &H20
Global Const HDSCHSTATUSMCHG = &H20

'Shift values
Global Const SHIFT8 = &H100
Global Const SHIFT11 = &H800    'Used in Price Level
Global Const SHIFT12 = &H1000
Global Const SHIFT17 = &H20000
Global Const SHIFT23 = &H800000


'Btrieve
Global Const INDEXKEY0 = 0
Global Const INDEXKEY1 = 1
Global Const INDEXKEY2 = 2
Global Const INDEXKEY3 = 3
Global Const INDEXKEY4 = 4
Global Const INDEXKEY5 = 5
Global Const INDEXKEY6 = 6
Global Const INDEXKEY7 = 7


'Test input fields
Global Const NOMSG = 8          'Don't show error message
Global Const SHOWMSG = 16       'Show error message
Global Const ALLBLANK = 1       'Test for all blanks
Global Const ALLMANBLANK = 2    'Test for all mandatory fields as all blanks
Global Const ALLMANDEFINED = 4  'Test for all mandatory fields as defined
                                'This field is the only one that should
                                'include SHOWMSG
Global Const TESTALLCTRLS = -1  'Parameter indicating that all controls should be tested

'Format string area
Global Const FMTCOMMA = 1       'Include commas
Global Const FMTDOLLARSIGN = 2  'Include dollar sign
Global Const FMTPERCENTSIGN = 4 'Include percent sign
Global Const FMTNEGASBRACKETS = 8 'Show negative numbers with brackets
Global Const FMTNEGATBACK = 16  'Show negative sign as last character
Global Const FMTLEAVEBLANK = 32 'Leave blank as blank
'Note: default negative sign shown as first character
'              no commas
'              no dollar sign
'              no percent sign
Global Const UNFMTDEFAULT = 0   'Leave value after removing special characters
Global Const UNFMTREMOVEDECZERO = 1 'If value after decimal point = 0, remove

'Filter Types when retrieving records to populated list boxes
Global Const NOFILTER = 0
Global Const CHARFILTER = 1
Global Const INTEGERFILTER = 2
Global Const STRINGFILTER = 3
Global Const CHARFILTERNOT = 4      'Not equal
Global Const INTEGERFILTERNOT = 5   'Not equal
Global Const STRINGFILTERNOT = 6    'Not equal

'Vehicle types
Global Const VEHCONV_WO_FEED = &H1
Global Const VEHCONV_W_FEED = &H2
Global Const VEHSELLING = &H4
Global Const VEHAIRING = &H8
Global Const VEHLOG = &H10
Global Const VEHVIRTUAL = &H20
Global Const VEHLOGVEHICLE = &H40
Global Const VEHPACKAGE = &H80
Global Const VEHREP_WO_CLUSTER = &H100                  '8-30-02
Global Const VEHREP_W_CLUSTER = &H200
Global Const VEHNTR = &H400
Global Const VEHSTDPKG = &H800                  '8-30-02 chg from 100
Global Const VEHSIMUL = &H1000                  '8-30-02 chg from 200
Global Const VEHBYMKT = &H2000   '8-30-02 chg from 400; If set and tgSpf.sMktBase = "Y", then requires igMktCode
Global Const VEHBYPASSNOLOG = &H4000 '8-30-02 chg from 800 Bypass vehicles that are not generating Logs
Global Const VEHSPORT = &H10000  'Sport (Game) vehicles (Can't used H8000)
Global Const VEHEXCLUDESPORT = &H20000  'Exclude sports when VEHCONV_WO_FEED adn VEHCONV_FEED selected
Global Const VEHIMPORTAFFILIATESPOTS = &H40000    'Import (Sports and Non-Sports)
Global Const VEHSPORTMINUELIVE = &H80000
Global Const VEHBYPASSWEGENER_OLA = &H100000
Global Const VEHTESTLOGMERGE = &H200000
Global Const VEHMOVESPORTTONONSPORT = &H400000
Global Const VEHMOVESPORTTOSPORT = &H800000
Global Const VEHLIVELOG = &H1000000
Global Const VEHONINSERTION = &H2000000
Global Const VEHEXCLUDEIMPORTINVOICESPOTS = &H4000000
Global Const VEHALLTYPES = &H11FFF
Global Const VEHCPMPKG = &H8000000
Global Const ACTIVEVEH = &H10000000                '8-30-02 chg from 2000
Global Const DORMANTVEH = &H20000000               '8-30-02 chg from 4000
Global Const VEHEXCLUDEPODNOPRGM = &H40000000      '2/8/21 - renamed from VEHEXCLUDECPM to VEHEXCLUDEPODNOPRGM; exclude podcast w/o program based on ad server.
'Biggest allowed number is H80000000 if

Global Const DIRECTRESPONSERANK = 1010
Global Const REMNANTRANK = 1020
Global Const PERINQUIRYRANK = 1030
Global Const TRADERANK = 1040
Global Const NOCHARGERANK = 1044
Global Const EXTRARANK = 1045
Global Const PROMORANK = 1050
Global Const PSARANK = 1060
Global Const RESERVATIONRANK = 1070



Global Const IMPORTINVOICESPOT = 2000
Global Const RANKMASK = &H7FF
Global Const PRICELEVELMASK = &H7800


'Option Fields
Global Const OFPROJECTION = &H1
Global Const OFBUSCAT = &H2
Global Const OFSHARE = &H4
Global Const OFREVSET = &H8
Global Const OFDELGUAR = &H10
'Global Const OFBILLCYCLE = &H20
Global Const OFCALENDARBILL = &H20
Global Const OFCOOPBILL = &H40
Global Const OFRESEARCH = &H80


Global Const LNALLOCATION = &H1
Global Const LNACQUISITION = &H2
Global Const LN1STPOSITION = &H4
'Global Const LNPREFERREDDT = &H8
'Global Const LNSOLOAVAIL = &H10

Global Const SPALLOCATION = &H1
Global Const SPNTRACQUISITION = &H2
Global Const SP1STPOSITION = &H4
Global Const SPPREFERREDDT = &H8
Global Const SPSOLOAVAIL = &H10
Global Const BESTFITWEIGHTNONE = &H20
Global Const BBSAMELINE = &H40
Global Const BYPASSHIPRI = &H80 'VBC NR

'Using Features2
Global Const REGIONALCOPY = &H1
Global Const SPLITCOPY = &H2
Global Const SPLITNETWORKS = &H4
Global Const BARTER = &H8
Global Const STRONGPASSWORD = &H10
Global Const MIXAIRTIMEANDREP = &H20
Global Const MERCHPROMOBYDOLLAR = &H40
Global Const GREATPLAINS = &H80

'Using Features3
Global Const USINGHUB = &H1
Global Const TAXONNTR = &H2
Global Const PROMOCOPY = &H4
Global Const MEDIACODEBYVEH = &H8
Global Const INCMEDIACODEAUDIOVAULT = &H10
Global Const PSAINTOCONTRACTAVAILS = &H20
Global Const PROMOINTOCONTRACTAVAILS = &H40
Global Const TAXONAIRTIME = &H80

'Using Features4
Global Const LOCKBOXBYVEHICLE = &H1
Global Const BYPASSDEMOGRAPHIC = &H2 'VBC NR
Global Const BYPASSDATENEEDED = &H4 'VBC NR
Global Const ALLOWMOVEONTODAY = &H8
Global Const CHGBILLEDPRICE = &H10
Global Const TAXBYUSA = &H20
Global Const TAXBYCANADA = &H40
Global Const INVSORTBYVEHICLE = &H80

'Using Features5
Global Const REMOTEEXPORT = &H1
Global Const REMOTEIMPORT = &H2
Global Const COMBINEAIRNTR = &H4
Global Const CNTRINVSORTRC = &H8
Global Const CNTRINVSORTLN = &H10
Global Const STATIONINTERFACE = &H20
Global Const RADAR = &H40
Global Const SUPPRESSTIMEFORM1 = &H80

'Using Features6
Global Const INSTALLMENT = &H1
Global Const INSTALLMENTREVENUEEARNED = &H2
Global Const BBNOTSEPARATELINE = &H4
Global Const BBCLOSEST = &H8        '0=Open/Close; 1=Closest open ot close
Global Const GETPAIDEXPORT = &H10
Global Const DIGITALCONTENT = &H20
Global Const GUARBYGRIMP = &H40
Global Const INVEXPORTPARAMETERS = &H80

'Using Features7
Global Const CSIBACKUP = &H1
Global Const BONUSCOMM = &H2
Global Const COMMFISCALYEAR = &H4
Global Const EXPORTREVENUE = &H8        '10-2-08
Global Const REGIONMIXLEN = &H10        '10/10/08
Global Const XDIGITALISCIEXPORT = &H20
Global Const WEGENEREXPORT = &H40
Global Const OLAEXPORT = &H80

'Using Features8
Global Const LRMANDATORY = &H1
Global Const SHOWCMMTONDETAILPAGE = &H2
Global Const ALLOWMSASPLITCOPY = &H4
Global Const RIVENDELLEXPORT = &H8
Global Const XDIGITALBREAKEXPORT = &H10
Global Const ISCIEXPORT = &H20
Global Const PREFEEDDEF = &H40
Global Const REPBYDT = &H80

'Using Features9
Global Const AFFILIATECRM = &H1 'Affiliate CRM
Global Const PC1STPOS = &H2     'Default Pre-recorded copy into 1st position
Global Const PROPOSALXML = &H4
Global Const LIMITISCI = &H8
Global Const IDCRESTRICTION = &H10
Global Const WEEKLYBILL = &H20
Global Const PRINTEDI = &H40
Global Const WORDWRAPVEHICLE = &H80

'Using Features10
Global Const ADDADVTTOISCI = &H1 'X-Digital: Add Advertiser name to ISCI
Global Const MIDNIGHTBASEDHOUR = &H2     'X-Digital: Spot Insertion using Midnight Based time (12=0, 11,...11p=23)
Global Const PKGLNRATEONBR = &H4
Global Const WegenerIPump = &H8
Global Const REPLACEDELWKWITHFILLS = &H10
Global Const VCREATIVEEXPORT = &H20
Global Const CONTRACTVERIFY = &H40
'9114
'Global Const UNITIDBYASTCODE = &H80 'X-Digital: substitue astCode for the Unit ID and generate xml table of ISCI
Global Const UNITIDBYASTCODEFORBREAK = &H80 'X-Digital: substitute astCode for the Unit ID for HB and HBP
'Sport Info
Global Const USINGSPORTS = &H1
Global Const PREEMPTREGPROG = &H2
Global Const USINGFEED = &H4
Global Const USINGLANG = &H8

'Using Feature
Global Const MATRIXEXPORT = &H1
Global Const REVENUEEXPORT = &H2
Global Const LIVECOPY = &H4
Global Const MULTIMEDIA = &H8
Global Const USINGLIVELOG = &H10
Global Const USINGREP = &H20
Global Const HIDDENOVERRIDE = &H40
Global Const BESTFITWEIGHT = &H80

Global Const MGORIGVEHONLY = &H1
Global Const MGSCHVEHONLY = &H2
Global Const MGEITHERVEH = &H4
Global Const FILLORIGVEHONLY = &H8
Global Const FILLSCHVEHONLY = &H10
Global Const FILLEITHERVEH = &H20
Global Const MGRULESINCOPY = &H40
Global Const RSCHCUSTDEMO = &H80

'Automation Equipment
Global Const DALET = &H1
Global Const PROPHETNEXGEN = &H2
Global Const SCOTT = &H4
Global Const DRAKE = &H8
Global Const RCS4DIGITCART = &H10
Global Const PROPHETWIZARD = &H20
Global Const PROPHETMEDIASTAR = &H40
Global Const IMEDIATOUCH = &H80

'spfAutoType2
Global Const AUDIOVAULT = &H1
Global Const AUDIOVAULTRPS = &H2    '11-15-10AUDIOVAULTAIR = &H2  unused for now 'VBC NR
Global Const WIREREADY = &H4
Global Const ENCO = &H8
Global Const RN_REP = &H10
Global Const RN_NET = &H20
Global Const SIMIAN = &H40              '8-20-08
Global Const RCS5DIGITCART = &H80

'spfAutoType3
Global Const AUDIOVAULTAIR = &H1        '2-18-11
Global Const WIDEORBIT = &H2            '1/3/12
Global Const JELLI = &H4
Global Const ENCOESPN = &H8
Global Const SCOTT_V5 = &H10             '8-16-13
Global Const ZETTA = &H20               '1-5-16
Global Const STATIONPLAYLIST = &H40     '5/10/18
Global Const RADIOMAN = &H80     '5/10/18

'Vehicle features
'Using Features1
Global Const EXPORTINSERTION = &H1
Global Const EXPORTLOG = &H2
Global Const IMPORTINSERTION = &H4
Global Const IMPORTAFFILIATESPOTS = &H8
Global Const EXPORTISCIBYPLEDGE = &H10
'H20 left unused because old vehicles have a h20 (space) in this field
Global Const SUPPRESSWEBLOG = &H40
Global Const EXPORTPOSTEDTIMES = &H80

'Using Features2
Global Const XDSAPPLYMERGE = &H1
Global Const INVOICEAIRDATEWO = &H2


'Features1 stored into SAF: Site Option
Global Const MATRIXCAL = &H1
Global Const ENGRHIDEMEDIACODE = &H2    'Engineering Export: Hide Media Code
Global Const SHOWAUDIOTYPEONBR = &H4
Global Const SHOWPRICEONINSERTIONWITHACQUISTION = &H8
Global Const SALESFORCEEXPORT = &H10
Global Const EFFICIOEXPORT = &H20
Global Const JELLIEXPORT = &H40
Global Const COMPENSATION = &H80
'Features1 stored into SAF: Vehicle Option
Global Const SUPPRESSPREEMPTION = &H1  'Suppress Scheduling Pre-emption


'Features2 stored into SAF: Site Option
Global Const EVENTREVENUE = &H1
Global Const HIDEHIDDENLINES = &H2
Global Const CANCELCLAUSEMANDATORY = &H4
Global Const EMAILDISTRIBUTION = &H8
Global Const ACQUISITIONCOMMISSIONABLE = &H10
Global Const PAYMENTONCOLLECTION = &H20
Global Const TABLEAUEXPORT = &H40
Global Const TABLEAUCAL = &H80

'Features3 stored into SAF: Site Option
Global Const SUPPRESSNETCOMM = &H1
Global Const REQSTATIONPOSTING = &H2
Global Const SPLITCOPYSTATE = &HC0      'Split Copy Station State (0=Mailing; 1=License; 2=Physical)
Global Const SPLITCOPYLICENSE = &H4
Global Const SPLITCOPYPHYSICAL = &H8
Global Const FREEZEDEFAULT = &H10       'Freeze calculation default
Global Const INVEMAILINDEX = &H20
Global Const INVSENDEMAILINDEX = &H40
Global Const SUPPRESSZERODOLLARINVEXPT = &H80

'Features4 stored into SAF: Site Option
Global Const FILEMAKERIMPORT = &H1  'Filemaker Contract Import
Global Const ACT1CODES = &H2        'ACT1 Line-up
Global Const MKTNAMEONBR = &H4      'Market Name on BR
Global Const COMPRESSTRANSACTIONS = &H8 'On Collection screen compress transactions.
Global Const AVAILINCLUDERESERVATION = &H10
Global Const AVAILINCLUDEREMNANT = &H20
Global Const AVAILINCLDEDIRECTRESPONSES = &H40
Global Const AVAILINCLUDEPERINQUIRY = &H80

'Features5 stored into SAF: Site Option
Global Const PROGRAMMATICALLOWED = &H1  'Programmatic Allowed
Global Const SHOWAVAILCOUNT = &H2       'Programmatic Show Avail Count
Global Const SHOWCPPTAB = &H4           'Programmatic Show CPP Tab
Global Const SHOWCPMTAB = &H8           'Programmatic Show CPM Tab
Global Const SHOWPRICETAB = &H10        'Programmatic Show Price Tab
Global Const PODCASTAUDPCT = &H20
Global Const SHOWDAYDROPDOWN = &H40
Global Const CSVAFFIDAVITIMPORT = &H80

'Features6 stored into SAF: Site Option
Global Const BILLINGONINSERTIONS = &H1  'Insertion Order include Monthly Billed Summary
Global Const UNITIDBYASTCODEFORISCI = &H2 'X-Digital: substitute astCode for the Unit ID for Isci Method
Global Const SIGNATUREONPROPOSAL = &H4  'Signature on Proposal
Global Const CALCULATESEARCHONLINECHG = &H8 'Calculate Research totals on each line change
Global Const EDIAGYCODES = &H10         'Using EDI Client and Product codes
'Global Const ADVANCEAVAILS = &H20       'Unused: Show Avail/Protection and Resaerch Tabs on Proposal screen
Global Const RABCALENDAR = &H40         'RAB-Calendar export
Global Const OVERDUEEXPORT = &H80       'Affiliate Overdue export


'Features7 stored into SAF: Site Option
Global Const RABSTD = &H1  'RAB standard month export
Global Const RABCALSPOTS = &H2 'RAB Calendar by spots export
Global Const IMEDIA_MEDIACODE = &H4 '11/5/20 - TTP # 10013 - iMediaTouch Replace COM with Media Code
Global Const ISCIEXPORTFEATURE = &H8 '10023 ISCI Export
Global Const INVEMAILAIRONLY = &H10 '10016
Global Const INVEMAILNTRONLY = &H20  '10016
Global Const CUSTOMEXPORT = &H40
Global Const PODBILLOVERDELIVERED = &H80    'Bill Over-Delivered CPM Impressions

'Features8 in SAF 10048
Global Const PODAIRTIME = &H1
Global Const PODSPOTS = &H2
Global Const PODADSERVER = &H4
Global Const PODADSERVERVIEWONLY = &H8  'PODMIXALLOWED = &H8
Global Const PODSHOWWKOF = &H10

'User Options (urfAvailSettings
Global Const AVAILBYWD = &H1    'Weekly/Daily
Global Const AVAILBYHD = &H2    'Hourly/Daypart
Global Const AVAILBYSAI = &HC   'Sold/Avails/Inventory
Global Const AVAILBYNP = &H10   '#/%
Global Const AVAILBYUS = &H20   'Units/Seconds


'Call return status
Global Const CALLDONE = 1000 '100
Global Const CALLCANCELLED = 1001 '101
Global Const CALLTERMINATED = 1002 '102

'Define Job/list index into user table (the numbers must
'match the button index number in the basic10 screen)
Global Const RATECARDSJOB = 0
Global Const PROPOSALSJOB = 1
Global Const CONTRACTSJOB = 2
Global Const PROGRAMMINGJOB = 3
Global Const SPOTSJOB = 4
Global Const COPYJOB = 5
Global Const LOGSJOB = 6
Global Const POSTLOGSJOB = 7
Global Const INVOICESJOB = 8
Global Const COLLECTIONSJOB = 9
Global Const BUDGETSJOB = 10
Global Const SLSPCOMMSJOB = 11
Global Const STATIONFEEDJOB = 12
Global Const FEEDJOB = 13
Global Const REPORTSJOB = 14    'Make report the last index- when adding
                                'buttons to basic10 change reportsjob index
                                'and make this value match
Global Const PODBILLJOB = 15

'Set values which match index into tgURF.sWin(---)
'General
Global Const VEHICLESLIST = 20
Global Const AGENCIESLIST = 21
Global Const ADVERTISERSLIST = 22
Global Const COMPETITIVESLIST = 23
Global Const EXCLUSIONSLIST = 24
Global Const VEHICLEGROUPSLIST = 25
Global Const BUSCATEGORIESLIST = 26
Global Const POTENTIALCODESLIST = 27
Global Const FEEDNAMELIST = 62
Global Const MISSEDREASONSLIST = 56
Global Const SPLITNETSLIST = 64
'Sales
Global Const SALESSOURCESLIST = 28
Global Const SALESREGIONSLIST = 29
Global Const SALESOFFICESLIST = 30
Global Const SALESTEAMSLIST = 31
Global Const SALESPEOPLELIST = 32
Global Const REVENUESETSLIST = 33
Global Const BOILERPLATESLIST = 34
Global Const RESEARCHLIST = 35
Global Const DEMOSLIST = 36
Global Const COMPETITORSLIST = 37
Global Const GEOITEM = 38
Global Const PODITEMSLIST = 39
Global Const PodTarget = 40
'Programming
Global Const EVENTTYPESLIST = 41
Global Const EVENTNAMESLIST = 42
Global Const AVAILNAMESLIST = 43
Global Const FEEDTYPESLIST = 44
Global Const GENRESLIST = 45
'Copy
Global Const MEDIADEFINITIONSLIST = 50
Global Const ANNOUNCERNAMESLIST = 51
'Accounting
Global Const TAXTABLELIST = 63
Global Const ITEMBILLINGTYPESLIST = 57
Global Const INVOICESORTLIST = 58
Global Const LOCKBOXESLIST = 59
Global Const EDISERVICESLIST = 60
Global Const TRANSACTIONSLIST = 61
'Options
Global Const SITELIST = 68
Global Const USERLIST = 69
'Report types
Global Const CHFCONVMENU = 90
Global Const GENERICBUTTON = 91
Global Const DALLASFEED = 92
Global Const NYFEED = 93
Global Const BULKCOPY = 94
Global Const PHOENIXFEED = 95
Global Const CMMLCHG = 96
Global Const EXPORTAFFSPOTS = 97
Global Const PROPOSALPROJECTION = 98
Global Const PRODUCERLIST = 99          '1-15-04
Global Const PROVIDERLIST = 100         '1-15-04
Global Const SPLITREGIONLIST = 101         '9-19-06
Global Const EXPORTENCO = 102           '11-04-06
Global Const AVAILSCOMBO = 103          '12-18-07 'VBC NR


Global Const SPORTPROPOSAL = 200
Global Const SPORTINVENTORY = 201
Global Const SPORTREPORT = 202

'Index for igNoExeWinRes

'Child window call child window
Global Const CALLNONE = -1
Global Const CALLSOURCESALESOFFICE = SALESOFFICESLIST 'Sales source
Global Const CALLSOURCEENAME = EVENTNAMESLIST        'Event Name
Global Const CALLSOURCEVEHICLE = VEHICLESLIST
Global Const CALLSOURCESALESPERSON = SALESPEOPLELIST
Global Const CALLSOURCEAGENCY = AGENCIESLIST
Global Const CALLSOURCEADVERTISER = ADVERTISERSLIST
Global Const CALLSOURCECONTRACT = CONTRACTSJOB
Global Const CALLSOURCERATECARD = RATECARDSJOB
Global Const CALLSOURCECOLLECT = COLLECTIONSJOB
Global Const CALLSOURCESPOTS = SPOTSJOB
Global Const CALLSOURCECOPY = COPYJOB
Global Const CALLSOURCEPRG = PROGRAMMINGJOB
Global Const CALLSOURCECBILL = 100
Global Const CALLSOURCEPEVENT = 101
Global Const CALLSOURCEPOSTITEM = 102
Global Const CALLSOURCEVEHOPT = 103
Global Const CALLSOURCEGAME = 104
Global Const CALLSOURCEFEED = FEEDJOB
Global Const CALLSOURCEALERTS = 210
Global Const CALLSOURCEPODITEM = PODITEMSLIST 'poditem category listing
Global Const CALLSOURCEPODTARGET = PodTarget 'poditem category listing

'Report constants
'Rate Card job report constants
Public Const RC_RCITEMS = 0             'Rate carditems
Public Const RC_DAYPARTS = 1            'Dayparts
'Sales commissions job report constants
Public Const COMM_SALESCOMM = 0         'sales commission on billing
Public Const COMM_PROJECTION = 1        'projection report

'Projections job report constants
Public Const PRJ_SALESPERSON = 0
Public Const PRJ_VEHICLE = 1
Public Const PRJ_OFFICE = 2
Public Const PRJ_CATEGORY = 3
'Global Const PRJ_SCENARIO = 4
Public Const PRJ_POTENTIAL = 4
'Orders, Proposals, and Spots jobs report constants
Public Const CNT_BR = 0                     'BRoadcast contracts (proposal, narrow & wide)
Public Const CNT_PAPERWORK = 1              'Paperwork, summary
Public Const CNT_SPTSBYADVT = 2             'Spots by Advt
Public Const CNT_SPTSBYDATETIME = 3         'Spots by Date and Time
Public Const CNT_BOB_BYCNT = 4              'Business Booked (projection)
Public Const CNT_RECAP = 5                  'Recap
Public Const CNT_PLACEMENT = 6              'Spot Placement
Public Const CNT_DISCREP = 7                'discrepancy
Public Const CNT_MG = 8                     'makegoods (MG)
Public Const CNT_SPOTTRAK = 9               'Spot Tracking
Public Const CNT_COMLCHG = 10               'Commercial Change
Public Const CNT_HISTORY = 11               'History
Public Const CNT_AFFILTRAK = 12             'Affiliate Tracking
Public Const CNT_SPOTSALES = 13             'Spot Sales
Public Const CNT_MISSED = 14                'Missed
Public Const CNT_BOB_BYSPOT = 15            'business Booked by Spots (Spot Projection)
Public Const CNT_BOB_BYSPOT_REPRINT = 16    'Business Booked Reprint (Projection reprint)
Public Const CNT_QTRLY_AVAILS = 17          'Quarterly Avails
Public Const CNT_AVG_PRICES = 18            'Weekly Average Prices
Public Const CNT_ADVT_UNITS = 19            'Advt Units Ordered
Public Const CNT_SALES_CPPCPM = 20          'Sales Analysis by CPP CPM
Public Const CNT_AVGRATE = 21               'Average Rate
Public Const CNT_TIEOUT = 22                'Tie Out
Public Const CNT_BOB = 23                   'Billed & Booked Report
Public Const CNT_SALESACTIVITY = 24         'Sales Activity
Public Const CNT_SALESCOMPARE = 25          'Sales Comparison
Public Const CNT_CUMEACTIVITY = 26          'Cumulative Activity
Public Const CNT_MAKEPLAN = 27              'Avg prices needed to make plan
Public Const CNT_VEHCPPCPM = 28             'Current CPP & CPM by vehicle
Public Const CNT_SALESANALYSIS = 29         'Sales Analysis Summary
Public Const CNT_INSERTION = 30             'Insertion Orders
Public Const CNT_MGREVENUE = 31             'MG Revenue 6-16-00
Public Const CNT_ACCRUEDEFER = 32           '12-20-06 rptselcb exe module
'Duplicate constants OK because they are in different modules
Public Const CNT_HILORATE = 33              '5-14-10 Hi-Lo Spot Rate (Political) (rptselcb)
Public Const CNT_DISCREP_SUM = 34           '6-21-16 spot discrepancy summary by month (rptselcb)
Public Const CNT_SPTCOMBO = 35              'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines

Public Const CNT_DAILY_SALESACTIVITY = 31   'Daily Sales Activity by Contract 6-5-01
Public Const CNT_SALESACTIVITY_SS = 32      'Sales Activity by Sales Source or Market (vehicle) 7-25-02
Public Const CNT_SALESPLACEMENT = 33        'Sales Placement 7-25-02
Public Const CNT_NTRRECAP = 0               '(rptselnt) NTR Recap 4-2-03
Public Const CNT_NTRBB = 1                  ' (rptselnt) NTR Billed & Booked 4-2-03
Public Const CNT_MULTIMBB = 2               '1-25-08 Multimedia NTR Billed and Booked

Public Const CNT_VEH_UNITCOUNT = 34           '7-15-03 Vehicle Unit Count (rptselct)
Public Const CNT_BOBRECAP = 35                '4-14-05 Billed & Booked Recap
Public Const CNT_LOCKED = 36                  '4-5-06 locked avails
Public Const CNT_GAMESUMMARY = 37             '7-14-06
Public Const CNT_PAPERWORKTAX = 38            '4-9-07
Public Const CNT_BOBCOMPARE = 39             '9-13-07
Public Const CNT_CONTRACTVERIFY = 40         '4-8-13 Contract Verification (rptselct)
Public Const CNT_INSERTION_ACTIVITY = 41        '10-6-15 rptselct Insertion Order Activity
Public Const CNT_XML_ACTIVITY = 42           '4-1-16 Proposal XML Activity log (rptselct)

Public Const CNT_CP30UNIT = 0               'rptselCP30Unit

Public Const CNT_AVGRATECOMP = 0            '11/1/2018  average rate comparison
Public Const CNT_AVGPRCCOMP = 1             '11/1/2108

Public Const PRICE_RANKING_RPT = 0          '7-30-12 Spot Price Ranking (rptselrk)

Public Const REV_ON_BOOKS = 0               'rptselspotbb
Public Const SPOTREV_REGISTER = 1           'rptselspotbb: invoice revenue - revenue report from spots

'2-12-09 Copy Regions
Public Const CR_COPYREGION = 0
Public Const CR_SPLITREGION = 1

'2-23-10 Regional Copy Assignment
Public Const RG_REGION_ASSIGN = 0        'regional copy assignment (external)
Public Const RG_REGION_ASSIGNTRACE = 1   'regional copy tracing (internal)

Public Const POOLROTATION = -1234

'5-6-08
Public Const DELIVERY_AUDIENCE = 0           'Audience Deliver
Public Const DELIVERY_POSTBUY = 1              'Post Buy Analysis

'Invoice report options
Public Const INV_REGISTER = 0               'Invoice Registers (by inv #, advt, slsp, vehicle)
Public Const INV_VIEWEXPORT = 1             'View Export
Public Const INV_DISTRIBUTE = 2             'Billing distribution
Public Const INV_REMOTE_INVWS = 3           '8-29-02 Remote Invoice Worksheet
Public Const INV_DELINQUENT = 4             '8-29-02 Delinquent Affidavits
Public Const INV_UNBILLABLE = 5             '8-29-02 Unbillable Invoices
Public Const INV_CREDITMEMO = 6        '10-8-03 Credit/Debit Memo Invoice
Public Const INV_SUMMARY = 7                '6-28-05
Public Const INV_BARTERPAYMENT = 8           '11-28-06
Public Const INV_TAXREGISTER = 9           '1-30-07 tax register
Public Const INV_RECONCILE = 10            '11-30-07 Installment Reconciliation
Public Const INV_UNPOSTED_STATIONS = 11     '8-12-15 Unposted station spots for imported barter spots

'Collections report options
Public Const COLL_CASH = 0                  'Cash receipts
Public Const COLL_AGEPAYEE = 1              'Ageing by Payee
Public Const COLL_AGESLSP = 2               'Ageing by Salesperson
Public Const COLL_AGEVEHICLE = 3            'Ageing by Vehicle
Public Const COLL_DELINQUENT = 4            'Delinquent report
Public Const COLL_STATEMENT = 5             'Statments
Public Const COLL_PAYHISTORY = 6            'Payment History
Public Const COLL_CREDITSTATUS = 7          'Credit Status
Public Const COLL_DISTRIBUTE = 8            'Cash Distribution
Public Const COLL_CASHSUM = 9               'Cash summary
Public Const COLL_ACCTHIST = 10             'Account History
Public Const COLL_MERCHANT = 11             'Merchandising & Promotions
Public Const COLL_MERCHRECAP = 12            'Merchandising & Promotions Recap
Public Const COLL_AGEOWNER = 13             'Ageing by Owner
Public Const COLL_AGESS = 14                'Ageing by Sales Source
Public Const COLL_AGEPRODUCER = 15            '2-10-00 Ageing by Producer
Public Const COLL_POAPPLY = 16                    '9-11-03
Public Const COLL_AGEMONTH = 17             '8-21-15 AGeing by Month
Public Const COLL_PARTPAY = 18              '8-25-17 Participant Payables
Public Const COLL_SALESCOMM_COLL = 19    '1-19-18 sales commission on collection

Public Const LIST_PRODUCER = 0              '1-15-04 Vehicle Producer Lists
Public Const LIST_PROVIDER = 1              '1-15-04 Content Provider Lists

Public Const LIST_STDPKG = 6                   '9-30-15 Package vehicle list

Public Const COPY_ROT = 3                   '10-3-05 copy rotation
Public Const COPY_INVBYNUMBER = 4
Public Const COPY_INVBYISCI = 5
Public Const COPY_INVBYADVT = 6
Public Const COPY_INVBYSTARTDATE = 7
Public Const COPY_INVBYEXPDATE = 8
Public Const COPY_INVBYPURGE = 9
Public Const COPY_INVBYENTRYDATE = 10
Public Const COPY_INVUNAPPROVED = 12
Public Const COPY_REGIONS = 15              '7-18-00 copy by regions 'VBC NR
Public Const COPY_BOOK = 16                 '8-30-05
Public Const COPY_SPLITROT = 17            '1-30-09 Split copy/blackout rotation
Public Const COPY_SCRIPTAFFS = 18            '4-9-12 script affidavits
Public Const COPY_INVPRODUCER = 19              '4-10-13 Copy Inventory for Producer status

Public Const PL_LIVELOG = 2                 '12-8-05
Public Const PL_STATION_POSTING = 3         '1-23-19        station posting activity
Public Const RS_SUMMARY = 0                'Research report
Public Const RS_SPECIALSUMMARY = 1         'special research summary
Public Const RS_DEMORANK = 2                '9-10-15 demo ranking report
'Avails Combo reports - Sports and non-sports versions
Public Const AVAILSCOMBO_SPORTS = 0        '12-18-07
Public Const AVAILSCOMBO_NONSPORTS = 1

'User Reports                   '5-9-11
Public Const USER_OPTIONS = 0
Public Const USER_ACTIVITY = 1
Public Const USER_SUMMARY = 2   'Date: 4/15/2020 User Summary report

Public Const ACQ_PAY = 0        '7-1-16
Public Const ACQ_MONITOR = 1

'Prefeed Delivery and Engineering reports 5/6/10
'Index 0 and 1 for Feed report list are Recap & Pledge reports.  These are not related to Prefeed
Public Const PREFEED_DUMP = 2

'Programming reports from Links
Public Const PRG_SELLTOAIR = 0                  'Selling to Airing
Public Const PRG_AIRTOSELL = 1                 'Airing to selling
Public Const PRG_VEHAVAILCONFLICT = 2           'vehicle avail conflict
Public Const PRG_DELIVERY_BYVEH = 3              'delivery by vehicle
Public Const PRG_DELIVERY_BYFEED = 4             'delivery by feed
Public Const PRG_ENG_BYVEH = 5               'engineering by vehicle
Public Const PRG_ENG_BYFEED = 6                  'engineering by feed
Public Const PRG_AIRING_INV = 7     '3-31-15 Airing vehicle inventory

'6-10-14 Export types from Traffic menu or exports module
Public Const EXP_MATRIX = 14
Public Const EXP_EFFICIOREV = 20
Public Const EXP_EFFICIOPROJ = 21
Public Const EXP_TABLEAU = 22       '7-8-15 Tableau is the same export as matrix
'1-5-18 cash Receipts export and Inv Register exports
Public Const EXP_CASH = 24
Public Const Exp_INVREG = 25
'Public Const exp_RAB = 26           '1-23-20 RAB CRM export
Public Const EXP_MK = 26            'moved from exportlist
Public Const EXP_RAB = 27           '1-23-20 RAB CRM exp
Public Const EXP_CUST_REV = 28      'TTP 9992 - Custom Revenue Export
Public Const EXP_ADSERVERBILLDISC = 29      '3/12/21 - Ad Server Billing Discrepancy Export
Public Const EXP_AUDACYINV = 30     'TTP 10205 - 6/21/21 - JW - WO Invoice Export
Public Const EXP_AUDACYINVLINE = 31 '8/17/21 - JW - TTP 10233 - Audacy: line summary export

'Global spot types for Spots by Advt & spots by Date & Time
'bit selectivity for charged and different types of no charge spots
'bits defined right to left (0 to 9)
'5-25-05  Add bit 10 BB
Public Const SPOT_CHARGE = &H1         'charged,bit 0
Public Const SPOT_00 = &H2          '0.00, bit 1
Public Const SPOT_ADU = &H4         'ADU, bit 2
Public Const SPOT_BONUS = &H8       'bonus, bit 3
Public Const SPOT_EXTRA = &H10      'Extra (+fill), bit 4
Public Const SPOT_FILL = &H20       'Fill (-fill), bit5
Public Const SPOT_NC = &H40         'no charge , bit 6
Public Const SPOT_MG = &H80         'mg, bit 7
Public Const SPOT_RECAP = &H100     'recapturable, bit 8
Public Const SPOT_SPINOFF = &H200   'spinoff, bit 9
Public Const SPOT_BB = &H400        'BB, bit 10

Public Const ONE_MINUTE As Long = 60000 'VBC NR

' Internet Definitions
Public Const INTERNET_FLAG_RAW_DATA = &H40000000
Type MSI
    sName As String
    iVersion As Integer
End Type

'Monitor
'Public Const MONITOR_CSS = 0
'Public Const MONITOR_SSB = 1
'Public Const MONITOR_AEQ = 2
'Public Const MONITOR_ASI = 3
'Public Const MONITOR_ARQ = 4
'Public Const MONITOR_ASG = 5
'Public Const MONITOR_SC = 6
'Public Const MONITOR_CE = 7
'Public Const MONITOR_SFE = 8
'Public Const MONITOR_ME = 9
'Public Const MONITOR_EPE = 10
'Public Const MONITOR_ERE = 11
'Public Const MONITOR_GPE = 12
'Public Const MONITOR_BD = 13
Public Const MONITORTIMEINTERVAL As Long = 10   'In second
Public Const MONITORRUNINTERVAL As Long = 30    'In Seconds
Public Const MONITORELAPSEDINTERVAL As Long = 120    'In Seconds
Public Const MONITORELAPSEDINTERVALPERIOD As Long = 240   'In sec
Public Const MONITOREMAILINTERVAL As Long = 3600

'Extended feature
Public Const ASTEXTENDED_MG = 11
Public Const ASTEXTENDED_BONUS = 12
Public Const ASTEXTENDED_REPLACEMENT = 13
Public Const ASTAIR_MISSED_MG_BYPASS = 14           '12-17-18
Public Const ASTEXTENDED_MISSREASON = 100
Public Const ASTEXTENDED_ISCICHGD = 1000
'8032 be careful!  These are the indexes to radio buttons on vehOpt!
Public Const STATIONXMLWIDEORBIT As Integer = 9
Public Const STATIONXMLMARKETRON As Integer = 7
Public Const STATIONXMLNONE As Integer = 8
Public Const AGENCYXML As Integer = 0
'Features 8
Global Const AIRTIMETAG = &H1
Global Const PODCASTSPOTSTAG = &H2
Global Const PODCASTCPMTAG = &H4
Global Const MIXAIRPOD = &H8
Global Const PKGSHOWWKOF = &H10

'InvFeatures stored in agfxInvFeatures and adfxInvFeatures - TTP 10622
'Bit "0=None, 1=Suppress Net Amount for trade invoices"
Global Const TRADEINVSUPRESSNETAMT = &H1

'SPFX spfxInvExpFeature (Invoice Features)
Public Const INVEXP_AUDACYWO = &H1          'Audacy WO Invoice Export
Public Const INVEXP_AUDACYLINE = &H2        'Audacy Line Invoice Export
Public Const INVEXP_SELECTIVEEMAIL = &H4    'Selective Email Feature TTP 10826 / TTP 10813

'SPFX spfxIntFeature (Interface Features)
Public Const INT_HONORDST = &H1             'Network Honors DST

'SPFX spfxSchdFeature (Schedule Features) TTP 11033 - Schedule screen: disallow allow user who last edited a contract from being able to schedule it
Public Const SCHD_DISALLOWSAMEUSER = &H1             'disallow user who edited a contract from scheduling it

