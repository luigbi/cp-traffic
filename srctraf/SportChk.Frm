VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SportChk 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6030
   ClientLeft      =   240
   ClientTop       =   1620
   ClientWidth     =   8580
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   8580
   Begin VB.OptionButton rbcFix 
      Caption         =   "Fix"
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   45
      Width           =   960
   End
   Begin VB.OptionButton rbcCheck 
      Caption         =   "Check"
      Height          =   255
      Left            =   6540
      TabIndex        =   14
      Top             =   45
      Value           =   -1  'True
      Width           =   960
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   240
      Left            =   165
      TabIndex        =   12
      Top             =   5265
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox edcDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   5325
      MaxLength       =   10
      TabIndex        =   3
      Top             =   390
      Width           =   1170
   End
   Begin VB.TextBox edcDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   2190
      MaxLength       =   10
      TabIndex        =   1
      Top             =   390
      Width           =   1170
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6375
      Top             =   5340
   End
   Begin VB.PictureBox plcChkFix 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2790
      Left            =   1005
      ScaleHeight     =   2730
      ScaleWidth      =   6525
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1035
      Width           =   6585
      Begin VB.CheckBox ckcAll 
         Caption         =   "All Vehicles"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   6
         Top             =   2460
         Width           =   1350
      End
      Begin VB.ListBox lbcVehs 
         Appearance      =   0  'Flat
         Height          =   2340
         Left            =   150
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   60
         Width           =   6300
      End
      Begin VB.Label lacProcessing 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1545
         TabIndex        =   10
         Top             =   2475
         Width           =   4860
      End
   End
   Begin VB.CommandButton cmcChkFix 
      Appearance      =   0  'Flat
      Caption         =   "C&heck"
      Height          =   285
      Left            =   3060
      TabIndex        =   8
      Top             =   5400
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4935
      TabIndex        =   9
      Top             =   5400
      Width           =   1140
   End
   Begin VB.ListBox lbcError 
      Appearance      =   0  'Flat
      Height          =   1290
      Left            =   180
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3870
      Width           =   8160
   End
   Begin VB.Label lacScreen 
      Caption         =   "Check Sports Contracts"
      Height          =   225
      Left            =   105
      TabIndex        =   13
      Top             =   60
      Width           =   2115
   End
   Begin VB.Label lacDBPath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   45
      TabIndex        =   11
      Top             =   5760
      Width           =   6060
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   3405
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "Sunday End Date "
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   3750
      TabIndex        =   2
      Top             =   435
      Width           =   1410
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "Monday Start Date "
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   435
      Width           =   1575
   End
End
Attribute VB_Name = "SportChk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SportChk.Frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Type Defs (Removed)                                                            *
'*  CgfChanged                                                                            *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'**********************************************************
'                Check/Fix MODULE DEFINITIONS
'
'   Created : 4/25/94       By : D. LeVine
'   Modified :              By :
'
'**********************************************************
Option Explicit
Option Compare Text
Dim hmTo As Integer   'From file hanle
Dim hmCHF As Integer
Dim hmClf As Integer
Dim hmCff As Integer
Dim hmCgf As Integer

Dim hmGhf As Integer
Dim tmGhf As GHF
Dim tmGhfSrchKey0 As LONGKEY0
Dim imGhfRecLen As Integer

Dim hmGsf As Integer
Dim tmGsf As GSF
Dim tmGsfSrchKey3 As GSFKEY3
Dim imGsfRecLen As Integer

Dim tmChf As CHF            'CHF record image
'Contract record information
Dim tmClfList() As CLFLIST            'CLF record image
'Contract record information
Dim tmCfflist() As CFFLIST            'CFF record image
Dim tmCffSrchKey As CFFKEY0 'CFF key record image

Dim tmCgfList() As CGFLIST

Dim tmChfAdvtExt() As CHFADVTEXT


'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim lmNowDate As Long   'Todays date
Dim imFirstTime As Integer
Dim imCancelled As Integer
Dim imProcessing As Integer
Dim imFirstActivate As Integer
Dim bmChooseFix As Boolean

Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcVehs.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehs.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub cmcCancel_Click()
    If imProcessing Then
        imCancelled = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcChkFix_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llDate                        llStartDate                   llEndDate                 *
'*  slDate                        llTotalVeh                    llVehProc                 *
'*  llPercent                     llTotalNoToProc               llNoProc                  *
'*  ilError                       slMsg                         llLen                     *
'*  ilRes                         ilSvRet                       ilVef                     *
'*                                                                                        *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slName As String
    Dim ilTestOnly As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilSeeMsg As Integer

    ilTestOnly = True   'Check Only
    lbcError.Clear
    slStartDate = edcDate(0).Text
    If Not gValidDate(slStartDate) Then
        Beep
        MsgBox "Invalid Start Date ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
        edcDate(0).SetFocus
        Exit Sub
    End If
    If gWeekDayStr(slStartDate) <> 0 Then
        Beep
        MsgBox "Start Date must be at Monday ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
        edcDate(0).SetFocus
        Exit Sub
    End If
    slEndDate = UCase(edcDate(1).Text)
    If (Len(slEndDate) <> 0) And (slEndDate <> "TFN") Then
        If Not gValidDate(slEndDate) Then
            Beep
            edcDate(1).SetFocus
            Exit Sub
        End If
        If gDateValue(slEndDate) < gDateValue(slStartDate) Then
            Beep
            MsgBox "End Date must be after Start Date ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
            edcDate(1).SetFocus
            Exit Sub
        End If
        If gWeekDayStr(slEndDate) <> 6 Then
            Beep
            MsgBox "End Date must be at Sunday ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
            edcDate(1).SetFocus
            Exit Sub
        End If
    Else
        slEndDate = "12/31/2069"
    End If
    Screen.MousePointer = vbHourglass
    ilRet = 0
    'On Error GoTo cmcChkFixErr:
    slToFile = sgDBPath & "Messages\" & "SportChk.csv"
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo cmcChkFixErr:
            'hmTo = FreeFile
            'Open slToFile For Append As hmTo
            ilRet = gFileOpen(slToFile, "Append", hmTo)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                cmcCancel.SetFocus
                Exit Sub
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo cmcChkFixErr:
            'hmTo = FreeFile
            'Open slToFile For Output As hmTo
            ilRet = gFileOpen(slToFile, "Output", hmTo)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                cmcCancel.SetFocus
                Exit Sub
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo cmcChkFixErr:
        'hmTo = FreeFile
        'Open slToFile For Output As hmTo
        ilRet = gFileOpen(slToFile, "Output", hmTo)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Print #hmTo, "Contract Sport Check " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, ""
    Print #hmTo, "Contract,Line #,Vehicle,Week or Event #,FlightRecordCount or GsfAirDate,GameRecordCount or CgfAirDate"
    ilSeeMsg = False
    imProcessing = True
    imCancelled = False
    plcGauge.Value = 0
    ReDim ilVefCode(0 To 0) As Integer
    For ilLoop = 0 To lbcVehs.ListCount - 1 Step 1
        If lbcVehs.Selected(ilLoop) Then
            slNameCode = tgVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)    'Get user name
            ilRet = gParseItem(slName, 3, "|", slName)    'Get user name
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get user name
            DoEvents
            If imCancelled Then
                Print #hmTo, "Cancelled " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                Close #hmTo
                imProcessing = False
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            ilVefCode(UBound(ilVefCode)) = Val(slCode)
            ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
        End If
    Next ilLoop
    ilRet = mSportCheck(slStartDate, slEndDate, ilVefCode())
    If imCancelled Then
        Erase ilVefCode
        Print #hmTo, "Cancelled " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Close #hmTo
        lacProcessing.Caption = "Processing- Cancelled "
    Else
        Print #hmTo, ""
        Print #hmTo, "Completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Close #hmTo
        plcGauge.Value = 100
        DoEvents
        lacProcessing.Caption = "Processing- Done "
    End If
    Erase ilVefCode
    cmcCancel.Caption = "Done"
    If Not ilRet Then
        ilRet = MsgBox("Results stored into " & slToFile, vbOKOnly + vbInformation, "Error")
    End If
    Screen.MousePointer = vbDefault
    imProcessing = False
    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Exit Sub
'cmcChkFixErr:
'    ilRet = Err.Number
'    Resume Next
End Sub
Private Sub edcDate_GotFocus(Index As Integer)
    If imFirstTime Then
        Screen.MousePointer = vbHourglass
        tmcStart.Enabled = True
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub edcDate_KeyPress(Index As Integer, KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    
    Erase tmClfList
    Erase tmCfflist
    Erase tmCgfList

    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmCgf)
    btrDestroy hmCgf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    
    Set SportChk = Nothing
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcVehs_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
        mSetCommands
    End If
End Sub
Private Sub lbcVehs_GotFocus()
'    Screen.MousePointer = vbHourGlass
'    lbcVehs.Clear
'    mPopulate

'    mSetCommands      'Disable LinksDef if No Vehicles are selected
'    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()

    Dim ilRet As Integer   'Return from btrieve calls
    Dim slDate As String
    imFirstActivate = True
    imTerminate = False
    imSetAll = True
    imAllClicked = False
    imCancelled = False
    imProcessing = False
    Screen.MousePointer = vbHourglass

    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    SportChk.Height = lacDBPath.Top + 2 * lacDBPath.Height
    gCenterStdAlone SportChk
    'SportChk.Show

'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gGetSchParameters
    Screen.MousePointer = vbHourglass
    lacDBPath.Caption = "D.B. Path=> " & sgDBPath
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    slDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slDate)
    edcDate(0).Text = Format$(gDateValue(slDate) + 1, "m/d/yy")
    bmChooseFix = False
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SportChk
    On Error GoTo 0
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SportChk
    On Error GoTo 0
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SportChk
    On Error GoTo 0
    hmCgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SportChk
    On Error GoTo 0
    hmGhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SportChk
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)
    hmGsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SportChk
    On Error GoTo 0
    imGsfRecLen = Len(tmGsf)
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    mTerminate
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilRet                         slTestSystem              *
'*  ilTestSystem                                                                          *
'******************************************************************************************

    Dim slCommand As String
    slCommand = sgCommandStr    'Command$
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:4/25/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate list box              *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilVehCode As Integer
    Dim ilVehIndex As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slLogDate As String
    'ilRet = gPopUserVehicleBox(SportChk, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH, lbcVehs, Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBox(SportChk, VEHSPORT + ACTIVEVEH, lbcVehs, tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopUserVehicleBox: Vehicle)", SportChk
        On Error GoTo 0
        'Add last log date
        For ilLoop = 0 To lbcVehs.ListCount - 1 Step 1
            slNameCode = tgVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVehCode = Val(slCode)
            ilVehIndex = gVpfFind(SportChk, ilVehCode)
            gUnpackDate tgVpf(ilVehIndex).iLLD(0), tgVpf(ilVehIndex).iLLD(1), slLogDate
            If slLogDate <> "" Then
                lbcVehs.List(ilLoop) = lbcVehs.List(ilLoop) & " " & slLogDate
            End If
        Next ilLoop
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    Exit Sub
End Sub
'************************************************************
'          Procedure Name : mSetCommands
'
'    Created : 4/17/94      By : D. Hannifan
'    Modified :             By :
'
'    Comments:  Set Control properties
'
'
'************************************************************
'
Private Sub mSetCommands()
    If lbcVehs.ListCount <= 0 Then
        cmcChkFix.Enabled = False
    Else
        If lbcVehs.SelCount > 0 Then
            cmcChkFix.Enabled = True
        Else
            cmcChkFix.Enabled = False
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: terminate Links                *
'*                                                     *
'*******************************************************
Private Sub mTerminate()

    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload SportChk
    'End
    igManUnload = NO
End Sub



Private Sub rbcCheck_Click()
cmcChkFix.Caption = "Check"
bmChooseFix = False
End Sub

Private Sub rbcFix_Click()
cmcChkFix.Caption = "Fix"
bmChooseFix = True
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If imFirstTime Then
        'mInitDDE
        imFirstTime = False
    End If
    Screen.MousePointer = vbHourglass
    lbcVehs.Clear
    mPopulate

    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Screen.MousePointer = vbDefault
End Sub

Function mSportCheck(slStartDate As String, slEndDate As String, ilVefCode() As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  blFoundACgf                                                                           *
'******************************************************************************************

    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    Dim ilChf As Integer
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim ilVef As Integer
    Dim ilRet As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilCgf As Integer
    Dim ilCgfCount As Integer
    Dim llDate As Long
    Dim ilFound As Integer
    Dim ilCntrError As Integer  'seems to never be used
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilVefIndex As Integer
    Dim slVehicleName As String
    Dim slGsfAirDate As String
    Dim slCgfAirDate As String
    Dim llNoProc As Long
    Dim llTotalNoToProc As Long
    Dim llPercent As Long
    Dim blFixOk As Boolean
    Dim slCffStartDate As String
    Dim llSeasonStartDate As Long
    Dim llSeasonEndDate As Long
    Dim blSeasonFd As Boolean
    Dim blDaysOk As Boolean
    Dim ilDay As Integer
    
    slCntrStatus = "HO"
    slCntrType = "C"
    ilHOType = 1
    sgCntrForDateStamp = ""
    ilRet = gObtainCntrForDate(SportChk, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOType, tmChfAdvtExt())
    If (ilRet <> CP_MSG_NOPOPREQ) And (ilRet <> CP_MSG_NONE) Then
        Exit Function
    End If
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    llNoProc = 0
    llTotalNoToProc = UBound(tmChfAdvtExt)
    For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        ilRet = gObtainCntrPlusGame(hmCHF, hmClf, hmCff, hmCgf, tmChfAdvtExt(ilChf).lCode, False, tmChf, tmClfList(), tmCfflist(), tmCgfList())
        If ilRet Then
            lacProcessing.Caption = "Processing contract " & tmChfAdvtExt(ilChf).lCntrNo
            ilCntrError = False
            For ilClf = LBound(tmClfList) To UBound(tmClfList) - 1 Step 1
                gUnpackDateLong tmClfList(ilClf).ClfRec.iStartDate(0), tmClfList(ilClf).ClfRec.iStartDate(1), llSDate
                gUnpackDateLong tmClfList(ilClf).ClfRec.iEndDate(0), tmClfList(ilClf).ClfRec.iEndDate(1), llEDate
                If llSDate <= llEDate Then
                    For ilVef = 0 To UBound(ilVefCode) - 1 Step 1
                        If tmClfList(ilClf).ClfRec.iVefCode = ilVefCode(ilVef) Then
                            ilVefIndex = gBinarySearchVef(ilVefCode(ilVef))
                            If ilVefIndex = -1 Then
                                Exit Function
                            End If
                            slVehicleName = Trim$(tgMVef(ilVefIndex).sName)
                            'Check if valid game date
                            For ilCgf = LBound(tmCgfList) To UBound(tmCgfList) - 1 Step 1
                                If (tmCgfList(ilCgf).CgfRec.lClfCode = tmClfList(ilClf).ClfRec.lCode) Then
                                    gUnpackDate tmCgfList(ilCgf).CgfRec.iAirDate(0), tmCgfList(ilCgf).CgfRec.iAirDate(1), slCgfAirDate
                                    tmGsfSrchKey3.iVefCode = tmClfList(ilClf).ClfRec.iVefCode
                                    tmGsfSrchKey3.iGameNo = tmCgfList(ilCgf).CgfRec.iGameNo
                                    ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)  'Get last current record to obtain date
                                    If ilRet = BTRV_ERR_NONE Then
                                        blSeasonFd = False
                                        tmGhfSrchKey0.lCode = tmGsf.lGhfCode
                                        ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get last current record to obtain date
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = tmClfList(ilClf).ClfRec.iVefCode) And (tmGsf.iGameNo = tmCgfList(ilCgf).CgfRec.iGameNo)
                                            gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStartDate
                                            gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEndDate
                                            If (gDateValue(slCgfAirDate) >= llSeasonStartDate) And (gDateValue(slCgfAirDate) <= llSeasonEndDate) Then
                                                blSeasonFd = True
                                                gUnpackDate tmGsf.iAirDate(0), tmGsf.iAirDate(1), slGsfAirDate
                                                If gDateValue(slGsfAirDate) <> gDateValue(slCgfAirDate) Then
                                                    '5/16/08 fix date issue. cgf will now match gsf.
                                                    If bmChooseFix = True Then
                                                        blFixOk = mFixCGFDate(slGsfAirDate, tmCgfList(ilCgf).CgfRec.lCode)
                                                    End If
                                                    If (blFixOk) And (bmChooseFix) Then
                                                        Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & tmCgfList(ilCgf).CgfRec.iGameNo & "," & slGsfAirDate & "," & slCgfAirDate & "-Fixed"
                                                        lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & tmCgfList(ilCgf).CgfRec.iGameNo & "," & slGsfAirDate & "," & slCgfAirDate & "-Fixed"
                                                        mFixCffDateAfterFixingCgf slCgfAirDate, slGsfAirDate, tmClfList(ilClf).ClfRec.iLine
                                                        gPackDate slGsfAirDate, tmCgfList(ilCgf).CgfRec.iAirDate(0), tmCgfList(ilCgf).CgfRec.iAirDate(1) 'don't think I need this anymore
                                                        tmCgfList(ilCgf).lStartDate = gDateValue(slGsfAirDate)
                                                        tmCgfList(ilCgf).lEndDate = gDateValue(slGsfAirDate)
                                                    Else
                                                        Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & tmCgfList(ilCgf).CgfRec.iGameNo & "," & slGsfAirDate & "," & slCgfAirDate
                                                        lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & tmCgfList(ilCgf).CgfRec.iGameNo & "," & slGsfAirDate & "," & slCgfAirDate
                                                    End If
                                                    ilCntrError = True
                                                    mSportCheck = False
                                                End If
                                                Exit Do
                                            Else
                                                ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    Exit Do
                                                End If
                                                tmGhfSrchKey0.lCode = tmGsf.lGhfCode
                                                ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get last current record to obtain date
                                            End If
                                        Loop
                                        If Not blSeasonFd Then
                                            Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & tmCgfList(ilCgf).CgfRec.iGameNo & "," & "Season Not Found"
                                            lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & tmCgfList(ilCgf).CgfRec.iGameNo & "," & "Season Not Found"
                                            ilCntrError = True
                                            mSportCheck = False
                                        End If
                                    Else
                                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                            ilRet = csiHandleValue(0, 7)
                                        End If
                                        Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & tmCgfList(ilCgf).CgfRec.iGameNo & "," & "GSF GetEqual Error" & "," & ilRet
                                        lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & tmCgfList(ilCgf).CgfRec.iGameNo & "," & "GSF GetEqual Error" & "," & ilRet
                                        ilCntrError = True
                                        mSportCheck = False
                                    End If
                                End If
                            Next ilCgf
                            'Check if cgf is in balance with cff
                            For llDate = llStartDate To llEndDate Step 7
                                ilCgfCount = 0
                                For ilCgf = LBound(tmCgfList) To UBound(tmCgfList) - 1 Step 1
                                    If (tmCgfList(ilCgf).lStartDate >= llDate) And (tmCgfList(ilCgf).lStartDate <= llDate + 6) And (tmCgfList(ilCgf).CgfRec.lClfCode = tmClfList(ilClf).ClfRec.lCode) Then
                                        ilCgfCount = ilCgfCount + tmCgfList(ilCgf).CgfRec.iNoSpots
                                    End If
                                Next ilCgf
                                If ilCgfCount > 0 Then
                                    ilFound = False
                                    For ilCff = LBound(tmCfflist) To UBound(tmCfflist) - 1 Step 1
                                        If (tmCfflist(ilCff).CffRec.iClfLine = tmClfList(ilClf).ClfRec.iLine) And (llDate + 6 >= tmCfflist(ilCff).lStartDate) And (llDate <= tmCfflist(ilCff).lEndDate) Then
                                            ilFound = True
                                            '9/27/16: Check days
                                            blDaysOk = True
                                            For ilDay = 0 To 6 Step 1
                                                If tmCfflist(ilCff).CffRec.iDay(ilDay) <> 1 Then
                                                    blDaysOk = False
                                                    Exit For
                                                End If
                                            Next ilDay
                                            'If tmCfflist(ilCff).CffRec.iSpotsWk <> ilCgfCount Then
                                            If (tmCfflist(ilCff).CffRec.iSpotsWk <> ilCgfCount) Or (Not blDaysOk) Then
                                            '5/16/08 fix spot problem by changing cff to match cgf.

                                                If bmChooseFix = True Then
                                                    blFixOk = False
                                                    blFixOk = mFixCffSpots(ilCgfCount, tmCfflist(ilCff))
                                                End If
                                                If bmChooseFix And blFixOk Then
                                                    Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & "," & tmCfflist(ilCff).CffRec.iSpotsWk & "," & ilCgfCount & "-fixed"
                                                    lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & "," & tmCfflist(ilCff).CffRec.iSpotsWk & "," & ilCgfCount & "-fixed"
                                                    tmCfflist(ilCff).CffRec.iSpotsWk = ilCgfCount
                                                Else
                                                    Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & "," & tmCfflist(ilCff).CffRec.iSpotsWk & "," & ilCgfCount
                                                    lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & "," & tmCfflist(ilCff).CffRec.iSpotsWk & "," & ilCgfCount
                                               End If
                                               ilCntrError = True
                                               mSportCheck = False
                                            End If
                                            Exit For
                                        End If
                                    Next ilCff
                                    If Not ilFound Then
                                        'create cff
                                        If bmChooseFix Then
                                            slCffStartDate = Format(llDate, "m/d/yy")
                                            ilRet = btrBeginTrans(hmCff, 1000)
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilRet = btrAbortTrans(hmCff)
                                                MsgBox "Transaction Not Completed", vbOKOnly + vbExclamation, "Transaction"
                                            End If
                                            ilRet = mModelCff(slCffStartDate, tmClfList(ilClf).ClfRec.iLine)

                                            If ilRet <> -1 Then
                                                blFixOk = mModifyCffDates(ilRet, slCffStartDate)
                                                If blFixOk = False Then
                                                    ilRet = btrAbortTrans(hmCff)
                                                    MsgBox "A cff file could not be created.", vbInformation + vbOKOnly
                                                Else
                                                    blFixOk = mFixCffSpots(ilCgfCount, tmCfflist(UBound(tmCfflist) - 1))
                                                    If blFixOk = False Then
                                                        ilRet = btrAbortTrans(hmCff)
                                                        MsgBox "The new cff files spots are incorrect", vbInformation + vbOKOnly
                                                    End If
                                                End If
                                            Else
                                                MsgBox "Couldn't model Cff record.  No record was created.", vbInformation + vbOKOnly
                                                ilRet = btrAbortTrans(hmCff)
                                            End If
                                            ilRet = btrEndTrans(hmCff)
                                        End If
                                        If bmChooseFix And blFixOk Then
                                            Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & ",None" & "," & ilCgfCount & "-fixed"
                                            lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & ",None" & "," & ilCgfCount & "-fixed"
                                        Else
                                            Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & ",None" & "," & ilCgfCount
                                            lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & ",None" & "," & ilCgfCount
                                        End If
                                        ilCntrError = True
                                        mSportCheck = False
                                    End If
                                Else
                                    For ilCff = LBound(tmCfflist) To UBound(tmCfflist) - 1 Step 1
                                        If (tmCfflist(ilCff).CffRec.iClfLine = tmClfList(ilClf).ClfRec.iLine) And (llDate + 6 >= tmCfflist(ilCff).lStartDate) _
                                        And (llDate <= tmCfflist(ilCff).lEndDate) And (tmCfflist(ilCff).iStatus <> 2) Then
                                            If tmCfflist(ilCff).CffRec.iSpotsWk > 0 Then
                                                'delete cff!
                                                If bmChooseFix = True Then
                                                    blFixOk = mFixCffByDelete(tmCfflist(ilCff))
                                                    If blFixOk Then
                                                        tmCfflist(ilCff).iStatus = 2
                                                    End If
                                                End If
                                                If bmChooseFix And blFixOk Then
                                                    Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & "," & tmCfflist(ilCff).CffRec.iSpotsWk & ",None" & "-fixed"
                                                    lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & "," & tmCfflist(ilCff).CffRec.iSpotsWk & ",None" & "-fixed"

                                                Else
                                                    Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & "," & tmCfflist(ilCff).CffRec.iSpotsWk & ",None"
                                                    lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llDate, "m/d/yy") & "," & tmCfflist(ilCff).CffRec.iSpotsWk & ",None"
                                                End If
                                                ilCntrError = True
                                                mSportCheck = False
                                            End If
                                            Exit For
                                        End If
                                    Next ilCff
                                End If
                                'If ilCntrError Then
                                '    Exit For
                                'End If

 '                               Next ilCff

                            Next llDate
                        End If
                        'If ilCntrError Then
                        '    Exit For
                        'End If

                    Next ilVef
                End If
                If imCancelled Then
                    Exit Function
                End If
                'If ilCntrError Then
                '    Exit For
                'End If
            Next ilClf
        End If
        If imCancelled Then
            Exit Function
        End If
        llNoProc = llNoProc + 1
        llPercent = (llNoProc * CSng(100)) / llTotalNoToProc
        If llPercent >= 99 Then
            llPercent = 99
        ElseIf llPercent < 0 Then
            llPercent = 0
        End If
        plcGauge.Value = llPercent
        DoEvents
    Next ilChf
End Function



Private Function mFixCGFDate(slGsfAirDate As String, llSearchKey As Long) As Boolean
Dim ilRecLen As Integer
Dim tlCgf As CGF
Dim ilRet As Integer

ilRecLen = Len(tlCgf)
ilRet = btrGetEqual(hmCgf, tlCgf, ilRecLen, llSearchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
If ilRet = BTRV_ERR_NONE Then
    gPackDate slGsfAirDate, tlCgf.iAirDate(0), tlCgf.iAirDate(1)
    ilRet = btrUpdate(hmCgf, tlCgf, ilRecLen)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "There was a problem updating the cgf file."
        mFixCGFDate = False
        Exit Function
    End If
Else
    MsgBox "There was a problem fixing the cgf file."
    mFixCGFDate = False
    Exit Function
End If

mFixCGFDate = True
End Function

Private Function mFixCffSpots(ilCgfCount As Integer, tlCffList As CFFLIST) As Boolean
Dim ilRet As Integer
Dim tlCff As CFF
Dim ilRecLen As Integer
Dim ilDay As Integer

mFixCffSpots = False
tmCffSrchKey.lChfCode = tlCffList.CffRec.lChfCode
tmCffSrchKey.iClfLine = tlCffList.CffRec.iClfLine
tmCffSrchKey.iCntRevNo = tlCffList.CffRec.iCntRevNo
tmCffSrchKey.iPropVer = tlCffList.CffRec.iPropVer
tmCffSrchKey.iStartDate(0) = tlCffList.CffRec.iStartDate(0)
tmCffSrchKey.iStartDate(1) = tlCffList.CffRec.iStartDate(1)

ilRecLen = Len(tlCff)
ilRet = btrGetEqual(hmCff, tlCff, ilRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
If ilRet = BTRV_ERR_NONE Then
    tlCff.iSpotsWk = ilCgfCount
    '9/27/16: reset days
    For ilDay = 0 To 6 Step 1
        tlCff.iDay(ilDay) = 1
    Next ilDay
    ilRet = btrUpdate(hmCff, tlCff, ilRecLen)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "There was a problem updating the cff file."
        Exit Function
    End If
Else
    MsgBox "There was a problem fixing the cff file."
    Exit Function
End If
mFixCffSpots = True
End Function

Private Function mFixCffByDelete(tlCffList As CFFLIST) As Boolean
Dim ilRet As Integer
Dim tlCff As CFF
Dim ilRecLen As Integer

tmCffSrchKey.lChfCode = tlCffList.CffRec.lChfCode
tmCffSrchKey.iClfLine = tlCffList.CffRec.iClfLine
tmCffSrchKey.iCntRevNo = tlCffList.CffRec.iCntRevNo
tmCffSrchKey.iPropVer = tlCffList.CffRec.iPropVer
tmCffSrchKey.iStartDate(0) = tlCffList.CffRec.iStartDate(0)
tmCffSrchKey.iStartDate(1) = tlCffList.CffRec.iStartDate(1)

ilRecLen = Len(tlCff)
ilRet = btrGetEqual(hmCff, tlCff, ilRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
If ilRet = BTRV_ERR_NONE Then
    ilRet = btrDelete(hmCff)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "There was a problem deleting the cff file."
        mFixCffByDelete = False
        Exit Function
    End If
Else
    MsgBox "There was a problem finding the cff file to delete."
    mFixCffByDelete = False
    Exit Function
End If
mFixCffByDelete = True
End Function


Private Sub mFixCffDateAfterFixingCgf(ByVal slOldDate As String, ByVal slNewDate As String, ByVal ilClfLine As Integer)
'I'm worried about the chfcode, as I took it from the clf I'm using istatus to skip records--hope that's ok.
'I assumed that cff start dates were always a monday.

'Input: the two dates are for cff search and the chfcode and clfline to match gsf record.
'Output: change to tmclfflist array and cff file. change tmcffsearchkey when looking for records

'if the before and after days are in the same week, we're done
'Otherwise, search cff's for any matches on the week of the 'new' date
'if we have a match, we're done. 'sportscheck' will fix the spots later in the module.
'if we didn't match the new date,
'search cffs again, this time against the old date. On match, change the cff start and end dates to include the gsf date.
'On no match, create a new cff record.

Dim slNewMonday As String
Dim slOldMonday As String
Dim ilCff As Integer
Dim ilOldCff As Integer
Dim ilRet As Integer

Dim blChanged As Boolean
Dim blFoundNewCff As Boolean

blFoundNewCff = False
slNewMonday = gObtainPrevMonday(slNewDate)
slOldMonday = gObtainPrevMonday(slOldDate)
If gDateValue(slNewMonday) <> gDateValue(slOldMonday) Then
    For ilCff = LBound(tmCfflist) To UBound(tmCfflist) - 1 Step 1
        If tmCfflist(ilCff).lStartDate = gDateValue(slNewMonday) And ((tmCfflist(ilCff).CffRec.iClfLine) = ilClfLine) Then
            blFoundNewCff = True    'found new.  Check for old in next loop and set it to ok delete.
        End If
   Next ilCff
   For ilOldCff = LBound(tmCfflist) To UBound(tmCfflist) - 1 Step 1
        If tmCfflist(ilOldCff).lStartDate = gDateValue(slOldMonday) And ((tmCfflist(ilOldCff).CffRec.iClfLine) = ilClfLine) Then
            If blFoundNewCff = False Then  'found old and no new, so ok to copy old cff and create new
                blChanged = mModifyCffDates(ilOldCff, slNewMonday)
                If blChanged = False Then
                    MsgBox "there was an error creating a new cff date from older cff date", vbInformation + vbOKOnly
                End If
            End If
            'this record had its date changed.  Use this status and cgf spot=0 to delete record.
            'tmcfflist(ilOldCff).iStatus = 2
            Exit Sub
        End If
    Next ilOldCff
    If blFoundNewCff = True Then            'cff on new date, but not on old.
        Exit Sub
    End If
            'if we got here, no cff found on old or new date.  Create new
    ilRet = btrBeginTrans(hmCff, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrAbortTrans(hmCff)
        MsgBox "Transaction Not Completed", vbOKOnly + vbExclamation, "Transaction"
    End If

    ilRet = mModelCff(slNewMonday, ilClfLine)
    If ilRet > -1 Then   'found a model
        blChanged = mModifyCffDates(ilRet, slNewMonday)
        If blChanged = False Then
            ilRet = btrAbortTrans(hmCff)
            MsgBox "there was an error creating a new cff date via modeling", vbInformation + vbOKOnly
        Else

        End If
    Else
        ilRet = btrAbortTrans(hmCff)
        MsgBox "Couldn't find a cff record to model from", vbExclamation + vbOKOnly
    End If
    ilRet = btrEndTrans(hmCff)
End If
End Sub

Private Function mModelCff(ByVal slNewDate As String, ilLineNo As Integer) As Integer
'   Copy a cff from an earlier week.
'   Input= date we want to get close to.  Output= the cff index # of record found.

    Dim llNewDate As Long
    
    llNewDate = gDateValue(slNewDate)
    
    'search cfflist and find latest date less than slnewdate.
    Dim ilCff As Integer
    For ilCff = UBound(tmCfflist) - 1 To LBound(tmCfflist) Step -1
        If (tmCfflist(ilCff).CffRec.iClfLine = ilLineNo) And (llNewDate > tmCfflist(ilCff).lStartDate) Then
            mModelCff = ilCff
            Exit Function
        End If
    Next ilCff
    'If UBound(tmcfflist) > LBound(tmcfflist) Then
    '    mModelCff = LBound(tmcfflist)   'lowest date to model from
    'Else
    '    mModelCff = -1
    'End If
    
    For ilCff = LBound(tmCfflist) To UBound(tmCfflist) - 1 Step 1
        If (tmCfflist(ilCff).CffRec.iClfLine = ilLineNo) Then
            mModelCff = ilCff
            Exit Function
        End If
    Next ilCff
    mModelCff = -1

End Function

Private Function mModifyCffDates(ilCffIndex As Integer, slMonday As String) As Boolean
' change cff record's dates to match new gsf dates
' input ilcffindex- Record to change. Start date to change to. index of clf.
'output.  Change cff file   Also changes tmcffSrchKey.
Dim ilRet As Integer
Dim ilRecLen As Integer
Dim tlCff As CFF
Dim llNewSunday As Long
Dim ilDay As Integer

ilRecLen = Len(tlCff)
tmCffSrchKey.lChfCode = tmCfflist(ilCffIndex).CffRec.lChfCode
tmCffSrchKey.iClfLine = tmCfflist(ilCffIndex).CffRec.iClfLine
tmCffSrchKey.iCntRevNo = tmCfflist(ilCffIndex).CffRec.iCntRevNo
tmCffSrchKey.iPropVer = tmCfflist(ilCffIndex).CffRec.iPropVer
tmCffSrchKey.iStartDate(0) = tmCfflist(ilCffIndex).CffRec.iStartDate(0)
tmCffSrchKey.iStartDate(1) = tmCfflist(ilCffIndex).CffRec.iStartDate(1)
ilRet = btrGetEqual(hmCff, tlCff, ilRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
If ilRet = BTRV_ERR_NONE Then
    gPackDate slMonday, tlCff.iStartDate(0), tlCff.iStartDate(1)
    llNewSunday = DateAdd("d", 6, slMonday)
    gPackDateLong llNewSunday, tlCff.iEndDate(0), tlCff.iEndDate(1)
    tlCff.lCode = 0
    '9/27/16: reset days
    For ilDay = 0 To 6 Step 1
        tlCff.iDay(ilDay) = 1
    Next ilDay
    ilRet = btrInsert(hmCff, tlCff, ilRecLen, INDEXKEY1)
    If ilRet <> BTRV_ERR_NONE Then
        mModifyCffDates = False
        Exit Function
    Else
        tmCfflist(UBound(tmCfflist)).CffRec = tlCff
        tmCfflist(UBound(tmCfflist)).lStartDate = gDateValue(slMonday)
        tmCfflist(UBound(tmCfflist)).lEndDate = llNewSunday
        ReDim Preserve tmCfflist(LBound(tmCfflist) To UBound(tmCfflist) + 1) As CFFLIST
    End If
Else
    mModifyCffDates = False
    Exit Function
End If
mModifyCffDates = True
End Function

