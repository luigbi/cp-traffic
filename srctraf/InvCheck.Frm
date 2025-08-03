VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form InvCheck 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5910
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
   ScaleHeight     =   5910
   ScaleWidth      =   8580
   Begin V81Utilities.CSI_ComboBoxList cbcStartMonth 
      Height          =   270
      Left            =   2340
      TabIndex        =   1
      Top             =   405
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   476
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin V81Utilities.CSI_ComboBoxList cbcEndMonth 
      Height          =   285
      Left            =   2325
      TabIndex        =   5
      Top             =   795
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   503
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin VB.TextBox edcEndYear 
      Height          =   315
      Left            =   4515
      TabIndex        =   7
      Top             =   795
      Width           =   1095
   End
   Begin VB.TextBox edcStartYear 
      Height          =   315
      Left            =   4515
      TabIndex        =   3
      Top             =   405
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   240
      Left            =   165
      TabIndex        =   12
      Top             =   4980
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6375
      Top             =   5055
   End
   Begin VB.CommandButton cmcChkFix 
      Appearance      =   0  'Flat
      Caption         =   "C&heck"
      Height          =   285
      Left            =   3060
      TabIndex        =   9
      Top             =   5115
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4935
      TabIndex        =   10
      Top             =   5115
      Width           =   1140
   End
   Begin VB.ListBox lbcError 
      Appearance      =   0  'Flat
      Height          =   3180
      Left            =   180
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1545
      Width           =   8160
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   3885
      TabIndex        =   6
      Top             =   825
      Width           =   555
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "Invoice End Month"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   585
      TabIndex        =   4
      Top             =   825
      Width           =   1605
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   180
      TabIndex        =   14
      Top             =   1260
      Width           =   8160
   End
   Begin VB.Label lacScreen 
      Caption         =   "Invoice Check"
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
      Top             =   5520
      Width           =   8190
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   3900
      TabIndex        =   2
      Top             =   435
      Width           =   555
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "Invoice Start Month"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   435
      Width           =   1605
   End
End
Attribute VB_Name = "InvCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of InvCheck.Frm on Wed 6/17/09 @ 12:56 P
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

Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey1 As CHFKEY1  'CHF key record image (contract #)
Dim imCHFRecLen As Integer        'RVF record length

'Contract record information
Dim tmClfList() As CLFLIST            'CLF record image
'Contract record information
Dim tmCfflist() As CFFLIST            'CFF record image
Dim tmCffSrchKey As CFFKEY0 'CFF key record image

Dim tmChfAdvtExt() As CHFADVTEXT

Dim hmRvf As Integer            'Receivable file handle
Dim tmRvf As RVF                'RVF record image
Dim tmRvfSrchKey3 As RVFKEY3            'RVF record image (Advertiser code)
Dim tmRvfSrchKey4 As RVFKEY4            'RVF record image (Advertiser code)
Dim imRvfRecLen As Integer        'RVF record length
Dim hmPhf As Integer            'Receivable History file handle- same structure as rvf
'Spot file
Dim tmSdf As SDF            'SDF record image
Dim tmSdfSrchKey0 As SDFKEY0 'SDF key record image (agency)
Dim tmSdfSrchKey3 As LONGKEY0 'SDF key record image (agency)
Dim tmSdfSrchKey5 As LONGKEY0 'SDF key record image (agency)
Dim hmSdf As Integer        'SDF Handle
Dim imSdfRecLen As Integer     'SDF record length
'Spot file
Dim tmSmf As SMF            'SMF record image
Dim tmSmfSrchKey0 As SMFKEY0 'SMF key record image (agency)
Dim hmSmf As Integer        'SMF Handle
Dim imSmfRecLen As Integer     'SMF record length
'Agency
Dim hmAgf As Integer            'Advertsier name file handle
Dim tmAgf As AGF                'ADF record image
Dim tmAgfSrchKey As INTKEY0            'ADF record image
Dim imAgfRecLen As Integer        'ADF record length

'Podcast CPM
Dim hmPcf As Integer
Dim tmPcf As PCF
Dim imPcfRecLen As Integer
Dim tmPcfSrchKey5 As PCFKEY5    'SBF key record image


Dim hmIbf As Integer            'Special Billing file handle
Dim tmIbf As IBF                'SBF record image
Dim tmIbfSrchKey0 As LONGKEY0
Dim tmIbfSrchKey1 As IBFKEY1
Dim tmIbfSrchKey2 As IBFKEY2
Dim tmIbfSrchKey3 As IBFKEY3
Dim imIbfRecLen As Integer       'SBF record length

'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim lmNowDate As Long   'Todays date
Dim imFirstTime As Integer
Dim imCancelled As Integer
Dim imProcessing As Integer
Dim imFirstActivate As Integer
Dim imNoMonths As Integer
'Dim lmInvMonthEndDate() As Long
Private Type INVMONTHENDDATE
    lStdMonthEndDate As Long
    lCalMonthEndDate As Long
    iMonth As Integer
    iYear As Integer
End Type
Dim tmInvMonthEndDate() As INVMONTHENDDATE

Dim lmCntrMonthGross() As Long
Dim lmRecMonthGross() As Long
Dim lmRecMonthNet() As Long



Private Sub cbcEndMonth_OnChange()
    mSetCommands
End Sub

Private Sub cbcStartMonth_GotFocus()
    If imFirstTime Then
        Screen.MousePointer = vbHourglass
        tmcStart.Enabled = True
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cbcStartMonth_OnChange()
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
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilSeeMsg As Integer
    Dim llDate As Long
    Dim slStr As String
    Dim ilMonth As Integer
    Dim slStartYear As String
    Dim slEndYear As String

    lbcError.Clear
    'slStartDate = edcDate(0).Text
    'If Not gValidDate(slStartDate) Then
    '    Beep
    '    MsgBox "Invalid Invoice Start Date ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
    '    edcDate(0).SetFocus
    '    Exit Sub
    'End If
    'If gDateValue(gObtainStartStd(slStartDate)) <> gDateValue(slStartDate) Then
    '    Beep
    '    MsgBox "Start Date must be start date of Invoice month ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
    '    edcDate(0).SetFocus
    '    Exit Sub
    'End If
    'slEndDate = UCase(edcDate(1).Text)
    'If Not gValidDate(slStartDate) Then
    '    Beep
    '    MsgBox "Invalid Invoice End Date ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
    '    edcDate(0).SetFocus
    '    Exit Sub
    'End If
    If cbcStartMonth.ListIndex < 0 Then
        Beep
        MsgBox "Start Month must be Specified ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
        cbcStartMonth.SetFocus
        Exit Sub
    End If
    If edcStartYear.Text = "" Then
        Beep
        MsgBox "Start Year must be Specified ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
        cbcStartMonth.SetFocus
        Exit Sub
    End If
    If cbcEndMonth.ListIndex < 0 Then
        Beep
        MsgBox "End Month must be Specified ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
        cbcStartMonth.SetFocus
        Exit Sub
    End If
    If edcEndYear.Text = "" Then
        Beep
        MsgBox "End Year must be Specified ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
        cbcStartMonth.SetFocus
        Exit Sub
    End If
    slStartDate = cbcStartMonth.GetItemData(cbcStartMonth.ListIndex) & "/15/" & edcStartYear.Text
    If Not gValidDate(slStartDate) Then
        Beep
        MsgBox "Invalid Start Information ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
        Exit Sub
    End If
    slEndDate = cbcEndMonth.GetItemData(cbcEndMonth.ListIndex) & "/16/" & edcEndYear.Text
    If Not gValidDate(slEndDate) Then
        Beep
        MsgBox "Invalid End Information ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
        Exit Sub
    End If
        
    If gDateValue(slEndDate) < gDateValue(slStartDate) Then
        Beep
        MsgBox "End Month/Year must be on or after Start Month/Year ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
        Exit Sub
    End If
    'If gDateValue(gObtainEndStd(slEndDate)) <> gDateValue(slEndDate) Then
    '    Beep
    '    MsgBox "End Date must be end date of Invoice month ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
    '    Exit Sub
    'End If
    Screen.MousePointer = vbHourglass
    imNoMonths = 1
    llDate = gDateValue(gObtainEndStd(slStartDate))
    ReDim tmInvMonthEndDate(0 To 1) As INVMONTHENDDATE
    tmInvMonthEndDate(0).lStdMonthEndDate = llDate
    tmInvMonthEndDate(0).lCalMonthEndDate = gDateValue(gObtainEndCal(Format(llDate, "ddddd")))
    tmInvMonthEndDate(0).iMonth = Month(Format(llDate, "ddddd"))
    tmInvMonthEndDate(0).iYear = Year(Format(llDate, "ddddd"))
    Do While llDate < gDateValue(slEndDate)
        llDate = gDateValue(gObtainEndStd(Format(llDate + 1, "m/d/yy")))
        imNoMonths = imNoMonths + 1
        tmInvMonthEndDate(UBound(tmInvMonthEndDate)).lStdMonthEndDate = llDate
        tmInvMonthEndDate(UBound(tmInvMonthEndDate)).lCalMonthEndDate = gDateValue(gObtainEndCal(Format(llDate, "ddddd")))
        tmInvMonthEndDate(UBound(tmInvMonthEndDate)).iMonth = Month(Format(llDate, "ddddd"))
        tmInvMonthEndDate(UBound(tmInvMonthEndDate)).iYear = Year(Format(llDate, "ddddd"))
        ReDim Preserve tmInvMonthEndDate(0 To UBound(tmInvMonthEndDate) + 1) As INVMONTHENDDATE
    Loop
    ilRet = 0
    'On Error GoTo cmcChkFixErr:
    slToFile = sgDBPath & "Messages\" & "InvCheck.csv"
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
                MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
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
                MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
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
            MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Print #hmTo, "Invoice Check " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, ""
    slStr = "Months Checked:"
    'For ilMonth = 0 To UBound(lmInvMonthEndDate) - 1 Step 1
    '    slStr = slStr & "," & gObtainStartStd(Format$(lmInvMonthEndDate(ilMonth) - 1, "m/d/yy")) & "-" & Format$(lmInvMonthEndDate(ilMonth), "m/d/yy")
    'Next ilMonth
    slStr = slStr & cbcStartMonth.Text & "/" & edcStartYear.Text & "-" & cbcEndMonth.Text & "/" & edcEndYear.Text
    Print #hmTo, slStr
    Print #hmTo, "Contract #,Advertiser,Invoice Month,Contract Adjusted Gross Amount,Receivable Gross Amount,Contract Adjusted Net Amount,Receivable Net Amount"
    ilSeeMsg = False
    imProcessing = True
    imCancelled = False
    plcGauge.Value = 0
    If gDateValue(gObtainStartStd(slStartDate)) < gDateValue(gObtainStartCal(slStartDate)) Then
        slStartDate = gObtainStartStd(slStartDate)
    Else
        slStartDate = gObtainStartCal(slStartDate)
    End If
    If gDateValue(gObtainEndStd(slEndDate)) > gDateValue(gObtainEndCal(slEndDate)) Then
        slEndDate = gObtainEndStd(slEndDate)
    Else
        slEndDate = gObtainEndCal(slEndDate)
    End If
    ilRet = mInvCheck(slStartDate, slEndDate)
    If imCancelled Then
        Print #hmTo, "Cancelled " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Print #hmTo, ""
        Print #hmTo, "If Billing as Aired:"
        Print #hmTo, "Adjusted Gross = Contract Gross - Missed spots - Cancelled spots - Hidden spots + MG/Outsides spots"
        Print #hmTo, "If Billing as Ordered:"
        Print #hmTo, "Adjusted Gross = Contract Gross"
        Print #hmTo, "Billing as Aired or Billing as Ordered:"
        Print #hmTo, "Adjusted Net = Adjusted Gross * (1 - Agency Commission)"
        Print #hmTo, "Note: Hidden lines bypassed, Missed, Cancelled, Hidden and MG/Outsides of hidden lines bypassed"
        Print #hmTo, "Note: Only Cash portion of contract checked. NTR Transactions bypassed.  Installment contracts bypassed"
        Close #hmTo
        lacProcessing.Caption = "Processing- Cancelled "
    Else
        Print #hmTo, ""
        Print #hmTo, "Completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Print #hmTo, ""
        Print #hmTo, "If Billing as Aired:"
        Print #hmTo, "Adjusted Gross = Contract Gross - Missed spots - Cancelled spots - Hidden spots + MG/Outsides spots"
        Print #hmTo, "If Billing as Ordered:"
        Print #hmTo, "Adjusted Gross = Contract Gross"
        Print #hmTo, "Billing as Aired or Billing as Ordered:"
        Print #hmTo, "Adjusted Net = Adjusted Gross * (1 - Agency Commission)"
        Print #hmTo, "Note: Hidden lines bypassed, Missed, Cancelled, Hidden and MG/Outsides of hidden lines bypassed"
        Print #hmTo, "Note: Only Cash portion of contract checked. NTR Transactions bypassed.  Installment contracts bypassed"
        Close #hmTo
        plcGauge.Value = 100
        DoEvents
        lacProcessing.Caption = "Processing- Done "
    End If
    cmcCancel.Caption = "Done"
    If Not ilRet Then
        ilRet = MsgBox("Results stored into " & slToFile, vbOKOnly + vbInformation, "Information")
    End If
    Screen.MousePointer = vbDefault
    imProcessing = False
    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Exit Sub
'cmcChkFixErr:
'    ilRet = Err.Number
'    Resume Next
End Sub


Private Sub edcEndYear_Change()
    mSetCommands
End Sub

Private Sub edcStartYear_Change()
    mSetCommands
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
    Erase tmInvMonthEndDate
    Erase lmCntrMonthGross
    Erase lmRecMonthGross
    Erase lmRecMonthNet

    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    ilRet = btrClose(hmPhf)
    btrDestroy hmPhf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmPcf)
    btrDestroy hmPcf
    ilRet = btrClose(hmIbf)
    btrDestroy hmIbf

    Set InvCheck = Nothing

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
    Dim slStr As String
    
    imFirstActivate = True
    imTerminate = False
    imCancelled = False
    imProcessing = False
    Screen.MousePointer = vbHourglass

    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    InvCheck.Height = lacDBPath.Top + 2 * lacDBPath.Height
    gCenterStdAlone InvCheck
    'InvCheck.Show

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
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", InvCheck
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", InvCheck
    On Error GoTo 0
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", InvCheck
    On Error GoTo 0
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", InvCheck
    On Error GoTo 0
    imRvfRecLen = Len(tmRvf)
    hmPhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", InvCheck
    On Error GoTo 0
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", InvCheck
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", InvCheck
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", InvCheck
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
    hmPcf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPcf, "", sgDBPath & "Pcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pcf.Btr)", InvCheck
    On Error GoTo 0
    imPcfRecLen = Len(tmPcf)
    hmIbf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmIbf, "", sgDBPath & "Ibf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ibf.Btr)", InvCheck
    On Error GoTo 0
    imIbfRecLen = Len(tmIbf)
    
    'Missed Reason
    slStr = ""
    'ReDim tgMRMnf(1 To 1) As MNF
    ReDim tgMRMnf(0 To 0) As MNF
    ilRet = gObtainMnfForType("M", slStr, tgMRMnf())
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mPopMonth
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
    'If Not gValidDate(edcDate(0).Text) Then
    If (cbcStartMonth.ListIndex < 0) Or (cbcEndMonth.ListIndex < 0) Or (edcStartYear.Text = "") Or (edcEndYear.Text = "") Then
        cmcChkFix.Enabled = False
    Else
        cmcChkFix.Enabled = True
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
    Unload InvCheck
    'End
    igManUnload = NO
End Sub




Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If imFirstTime Then
        'mInitDDE
        imFirstTime = False
    End If
    Screen.MousePointer = vbHourglass
    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Screen.MousePointer = vbDefault
End Sub

Function mInvCheck(slStartDate As String, slEndDate As String) As Integer
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
    Dim ilTest As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim ilRet As Integer
    Dim ilVefIndex As Integer
    Dim slVehicleName As String
    Dim llNoProc As Long
    Dim llTotalNoToProc As Long
    Dim llPercent As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilMonth As Integer
    Dim llDate As Long
    Dim slStr As String
    Dim slGross As String
    Dim slNet As String
    Dim llGross As Long
    Dim llNet As Long
    Dim slAgyRate As String
    Dim slPctTrade As String
    Dim ilAdf As Integer
    Dim slAdfName As String
    Dim slVefType As String
    
    mInvCheck = False
    slCntrStatus = "HO"
    slCntrType = "C"
    ilHOType = 1
    sgCntrForDateStamp = ""
    ilRet = gObtainCntrForDate(InvCheck, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOType, tmChfAdvtExt())
    If (ilRet <> CP_MSG_NOPOPREQ) And (ilRet <> CP_MSG_NONE) Then
        Exit Function
    End If
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    ilRet = mGetReceivableCntr(hmRvf, llStartDate, llEndDate)
    ilRet = mGetReceivableCntr(hmPhf, llStartDate, llEndDate)
    ReDim lmCntrMonthGross(0 To imNoMonths) As Long
    ReDim lmRecMonthGross(0 To imNoMonths) As Long
    ReDim lmRecMonthNet(0 To imNoMonths) As Long
    llNoProc = 0
    llTotalNoToProc = UBound(tmChfAdvtExt)
    For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, tmChfAdvtExt(ilChf).lCode, False, tmChf, tmClfList(), tmCfflist())
        If (ilRet) And (tmChf.sInstallDefined <> "Y") Then
            lacProcessing.Caption = "Processing contract " & tmChfAdvtExt(ilChf).lCntrNo
            For ilMonth = 0 To UBound(lmCntrMonthGross) - 1 Step 1
                lmCntrMonthGross(ilMonth) = 0
                lmRecMonthGross(ilMonth) = 0
                lmRecMonthNet(ilMonth) = 0
            Next ilMonth
            For ilClf = LBound(tmClfList) To UBound(tmClfList) - 1 Step 1
                ilVefIndex = gBinarySearchVef(tmClfList(ilClf).ClfRec.iVefCode)
                If ilVefIndex <> -1 Then
                    slVehicleName = Trim$(tgMVef(ilVefIndex).sName)
                    slVefType = Trim$(tgMVef(ilVefIndex).sType)
                Else
                    slVehicleName = ""
                    slVefType = ""
                End If
                gUnpackDateLong tmClfList(ilClf).ClfRec.iStartDate(0), tmClfList(ilClf).ClfRec.iStartDate(1), llSDate
                gUnpackDateLong tmClfList(ilClf).ClfRec.iEndDate(0), tmClfList(ilClf).ClfRec.iEndDate(1), llEDate
                If (llSDate <= llEDate) And (tmClfList(ilClf).ClfRec.sType <> "H") And (slVefType <> "R") Then
                    For ilCff = LBound(tmCfflist) To UBound(tmCfflist) - 1 Step 1
                        If tmClfList(ilClf).ClfRec.iLine = tmCfflist(ilCff).CffRec.iClfLine Then
                            gUnpackDateLong tmCfflist(ilCff).CffRec.iStartDate(0), tmCfflist(ilCff).CffRec.iStartDate(1), llFlStartDate
                            gUnpackDateLong tmCfflist(ilCff).CffRec.iEndDate(0), tmCfflist(ilCff).CffRec.iEndDate(1), llFlEndDate
                            If llFlStartDate <= llFlEndDate Then
                                Select Case tmCfflist(ilCff).CffRec.sPriceType
                                    Case "T"    'True
                                        If (tmCfflist(ilCff).CffRec.sDyWk <> "D") Then    'Weekly
                                            For llDate = llFlStartDate To llFlEndDate Step 7
                                                If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                                                    For ilMonth = 0 To UBound(tmInvMonthEndDate) - 1 Step 1
                                                        'If lmInvMonthEndDate(ilMonth) = gDateValue(gObtainEndStd(Format(llDate, "m/d/yy"))) Then
                                                        If ((tmInvMonthEndDate(ilMonth).lStdMonthEndDate = gDateValue(gObtainEndStd(Format(llDate, "m/d/yy")))) And (tmChf.sBillCycle = "S")) Or ((tmInvMonthEndDate(ilMonth).lCalMonthEndDate = gDateValue(gObtainEndCal(Format(llDate, "m/d/yy")))) And (tmChf.sBillCycle = "C")) Then
                                                            lmCntrMonthGross(ilMonth) = lmCntrMonthGross(ilMonth) + tmCfflist(ilCff).CffRec.lActPrice * tmCfflist(ilCff).CffRec.iSpotsWk
                                                        End If
                                                    Next ilMonth
                                                End If
                                            Next llDate
                                        Else
                                            For llDate = llFlStartDate To llFlEndDate Step 1
                                                For ilMonth = 0 To UBound(tmInvMonthEndDate) - 1 Step 1
                                                    'If lmInvMonthEndDate(ilMonth) = gDateValue(gObtainEndStd(Format(llDate, "m/d/yy"))) Then
                                                    If ((tmInvMonthEndDate(ilMonth).lStdMonthEndDate = gDateValue(gObtainEndStd(Format(llDate, "m/d/yy")))) And (tmChf.sBillCycle = "S")) Or ((tmInvMonthEndDate(ilMonth).lCalMonthEndDate = gDateValue(gObtainEndCal(Format(llDate, "m/d/yy")))) And (tmChf.sBillCycle = "C")) Then
                                                        lmCntrMonthGross(ilMonth) = lmCntrMonthGross(ilMonth) + tmCfflist(ilCff).CffRec.lActPrice * tmCfflist(ilCff).CffRec.iDay(gWeekDayLong(llDate))
                                                    End If
                                                Next ilMonth
                                            Next llDate
                                        End If
                                End Select
                            End If
                        End If
                    Next ilCff
                    
                End If
                If imCancelled Then
                    Exit Function
                End If
            Next ilClf
            mGetAdServerDollars
            If (tgSpf.sInvAirOrder = "A") Or (tgSpf.sInvAirOrder = "2") Then
                'Adjust for missed, MG and Outside
                ilRet = mMissedAdjustment(llStartDate, llEndDate)
                ilRet = mMGOutsideAdjustment(llStartDate, llEndDate)
            End If
            'Receivables
            ilRet = mGetReceivableDollars(hmRvf)
            ilRet = mGetReceivableDollars(hmPhf)
        End If
        ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
        If ilAdf <> -1 Then
            slAdfName = Trim$(tgCommAdf(ilAdf).sName)
        Else
            slAdfName = ""
        End If
        For ilMonth = 0 To UBound(tmInvMonthEndDate) - 1 Step 1
            If (lmCntrMonthGross(ilMonth) <> 0) Or (lmRecMonthNet(ilMonth) <> 0) Then
                slGross = gLongToStrDec(lmCntrMonthGross(ilMonth), 2)
                slPctTrade = gIntToStrDec(tmChf.iPctTrade, 0)
                If Val(slPctTrade) <> 0 Then
                    slGross = gDivStr(gMulStr(slGross, gSubStr("100", slPctTrade)), "100")
                End If
                slAgyRate = ""
                If tmChf.iAgfCode > 0 Then
                    tmAgfSrchKey.iCode = tmChf.iAgfCode
                    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slAgyRate = gIntToStrDec(tmAgf.iComm, 2)
                    End If
                End If
                If slAgyRate <> "" Then
                    slNet = gDivStr(gMulStr(slGross, gSubStr("100.00", slAgyRate)), "100.00")
                    llNet = gStrDecToLong(slNet, 2)
                    If (llNet <> lmRecMonthNet(ilMonth)) And (llNet + 1 <> lmRecMonthNet(ilMonth)) And (llNet - 1 <> lmRecMonthNet(ilMonth)) And (llNet <> lmRecMonthNet(ilMonth) + 1) And (llNet <> lmRecMonthNet(ilMonth) - 1) Then
                        'slStr = Trim$(Str(tmChf.lCntrNo)) & "," & slAdfName & "," & Format(lmInvMonthEndDate(ilMonth), "m/d/yy") & "," & slGross & "," & gLongToStrDec(lmRecMonthGross(ilMonth), 2) & "," & gLongToStrDec(llNet, 2) & "," & gLongToStrDec(lmRecMonthNet(ilMonth), 2)
                        slStr = Trim$(Str(tmChf.lCntrNo)) & "," & slAdfName & "," & tmInvMonthEndDate(ilMonth).iMonth & "/" & tmInvMonthEndDate(ilMonth).iYear & "," & slGross & "," & gLongToStrDec(lmRecMonthGross(ilMonth), 2) & "," & gLongToStrDec(llNet, 2) & "," & gLongToStrDec(lmRecMonthNet(ilMonth), 2)
                        Print #hmTo, slStr
                        'lbcError.AddItem Trim$(Str(tmChf.lCntrNo)) & " " & slAdfName & " " & Format(lmInvMonthEndDate(ilMonth), "m/d/yy") & " " & gLongToStrDec(llNet, 2) & " " & gLongToStrDec(lmRecMonthNet(ilMonth), 2)
                        lbcError.AddItem Trim$(Str(tmChf.lCntrNo)) & " " & slAdfName & " " & tmInvMonthEndDate(ilMonth).iMonth & "/" & tmInvMonthEndDate(ilMonth).iYear & " " & gLongToStrDec(llNet, 2) & " " & gLongToStrDec(lmRecMonthNet(ilMonth), 2)
                    End If
                Else
                    llGross = gStrDecToLong(slGross, 2)
                    If (llGross <> lmRecMonthGross(ilMonth)) Then
                        If tmChf.iAgfCode = 0 Then
                            slStr = Trim$(Str(tmChf.lCntrNo)) & "," & slAdfName & "," & tmInvMonthEndDate(ilMonth).iMonth & "/" & tmInvMonthEndDate(ilMonth).iYear & "," & gLongToStrDec(llGross, 2) & "," & gLongToStrDec(lmRecMonthGross(ilMonth), 2) & "," & gLongToStrDec(llGross, 2) & "," & gLongToStrDec(lmRecMonthNet(ilMonth), 2)
                            Print #hmTo, slStr
                            lbcError.AddItem Trim$(Str(tmChf.lCntrNo)) & " " & slAdfName & " " & tmInvMonthEndDate(ilMonth).iMonth & "/" & tmInvMonthEndDate(ilMonth).iYear & " " & gLongToStrDec(llGross, 2) & " " & gLongToStrDec(lmRecMonthGross(ilMonth), 2)
                        Else
                            slStr = Trim$(Str(tmChf.lCntrNo)) & "," & slAdfName & "," & tmInvMonthEndDate(ilMonth).iMonth & "/" & tmInvMonthEndDate(ilMonth).iYear & "," & gLongToStrDec(llGross, 2) & "," & gLongToStrDec(lmRecMonthGross(ilMonth), 2) & "," & "???" & "," & gLongToStrDec(lmRecMonthNet(ilMonth), 2)
                            Print #hmTo, slStr
                            lbcError.AddItem Trim$(Str(tmChf.lCntrNo)) & " " & slAdfName & " " & tmInvMonthEndDate(ilMonth).iMonth & "/" & tmInvMonthEndDate(ilMonth).iYear & " " & gLongToStrDec(llGross, 2) & " " & gLongToStrDec(lmRecMonthGross(ilMonth), 2)
                        End If
                    End If
                End If
            End If
        Next ilMonth
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

Private Function mMissedAdjustment(llStartDate As Long, llEndDate As Long) As Integer
    Dim slDate As String
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim tlClf As CLF
    Dim ilCff As Integer
    Dim llDate As Long
    Dim ilMonth As Integer
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim ilLoop As Integer
    
    mMissedAdjustment = True
    tmSdfSrchKey5.lCode = tmChf.lCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.lChfCode = tmChf.lCode)
        If tmSdf.sSchStatus = "M" Then
            For ilLoop = LBound(tgMRMnf) To UBound(tgMRMnf) - 1 Step 1
                If tmSdf.iMnfMissed = tgMRMnf(ilLoop).iCode Then
                    If tgMRMnf(ilLoop).iGroupNo = 2 Then
                        'Bypass taking dollars from contract as missed billed
                        tmSdf.sSpotType = "X"
                        tmSdf.sSchStatus = "S"
                    End If
                    Exit For
                End If
            Next ilLoop
        End If
        If ((tmSdf.sSchStatus = "M") Or (tmSdf.sSchStatus = "C") Or (tmSdf.sSchStatus = "H")) And (tmSdf.sSpotType <> "X") Then
            For ilClf = LBound(tmClfList) To UBound(tmClfList) - 1 Step 1
                If tmSdf.iLineNo = tmClfList(ilClf).ClfRec.iLine Then
                    tlClf = tmClfList(ilClf).ClfRec
                    If (tlClf.sType <> "H") And (tlClf.sType <> "P") Then
                        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                        If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                            For ilMonth = 0 To UBound(tmInvMonthEndDate) - 1 Step 1
                                'If lmInvMonthEndDate(ilMonth) = gDateValue(gObtainEndStd(Format(llDate, "m/d/yy"))) Then
                                If ((tmInvMonthEndDate(ilMonth).lStdMonthEndDate = gDateValue(gObtainEndStd(Format(llDate, "m/d/yy")))) And (tmChf.sBillCycle = "S")) Or ((tmInvMonthEndDate(ilMonth).lCalMonthEndDate = gDateValue(gObtainEndCal(Format(llDate, "m/d/yy")))) And (tmChf.sBillCycle = "C")) Then
                                    For ilCff = LBound(tmCfflist) To UBound(tmCfflist) - 1 Step 1
                                        If tmClfList(ilClf).ClfRec.iLine = tmCfflist(ilCff).CffRec.iClfLine Then
                                            gUnpackDateLong tmCfflist(ilCff).CffRec.iStartDate(0), tmCfflist(ilCff).CffRec.iStartDate(1), llFlStartDate
                                            gUnpackDateLong tmCfflist(ilCff).CffRec.iEndDate(0), tmCfflist(ilCff).CffRec.iEndDate(1), llFlEndDate
                                            If (llFlStartDate <= llFlEndDate) And (llDate >= llFlStartDate) And (llDate <= llFlEndDate) Then
                                                Select Case tmCfflist(ilCff).CffRec.sPriceType
                                                    Case "T"    'True
                                                       lmCntrMonthGross(ilMonth) = lmCntrMonthGross(ilMonth) - tmCfflist(ilCff).CffRec.lActPrice
                                                End Select
                                                Exit For
                                            End If
                                        End If
                                    Next ilCff
                                    Exit For
                                End If
                            Next ilMonth
                        End If
                    End If
                    Exit For
                End If
            Next ilClf
        End If
        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
End Function

Private Function mMGOutsideAdjustment(llStartDate As Long, llEndDate As Long) As Integer
    Dim slDate As String
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim tlClf As CLF
    Dim ilCff As Integer
    Dim llDate As Long
    Dim llSchDate As Long
    Dim ilMonth As Integer
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    
    mMGOutsideAdjustment = True
    For ilClf = LBound(tmClfList) To UBound(tmClfList) - 1 Step 1
        tlClf = tmClfList(ilClf).ClfRec
        If (tlClf.sType <> "H") And (tlClf.sType <> "P") Then
            tmSmfSrchKey0.lChfCode = tlClf.lChfCode
            tmSmfSrchKey0.iLineNo = tlClf.iLine
            tmSmfSrchKey0.lFsfCode = 0
            slDate = Format$(llStartDate, "m/d/yy")
            gPackDate slDate, ilDate0, ilDate1
            tmSmfSrchKey0.iMissedDate(0) = 0
            tmSmfSrchKey0.iMissedDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tlClf.lChfCode)
                If tmSmf.iLineNo <> tlClf.iLine Then
                    Exit Do
                End If
                tmSdfSrchKey3.lCode = tmSmf.lSdfCode
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) And (tmSdf.sSpotType <> "X") Then
                    gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llDate
                    If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                        For ilMonth = 0 To UBound(tmInvMonthEndDate) - 1 Step 1
                            'If lmInvMonthEndDate(ilMonth) = gDateValue(gObtainEndStd(Format(llDate, "m/d/yy"))) Then
                            If ((tmInvMonthEndDate(ilMonth).lStdMonthEndDate = gDateValue(gObtainEndStd(Format(llDate, "m/d/yy")))) And (tmChf.sBillCycle = "S")) Or ((tmInvMonthEndDate(ilMonth).lCalMonthEndDate = gDateValue(gObtainEndCal(Format(llDate, "m/d/yy")))) And (tmChf.sBillCycle = "C")) Then
                                For ilCff = LBound(tmCfflist) To UBound(tmCfflist) - 1 Step 1
                                    If tmClfList(ilClf).ClfRec.iLine = tmCfflist(ilCff).CffRec.iClfLine Then
                                        gUnpackDateLong tmCfflist(ilCff).CffRec.iStartDate(0), tmCfflist(ilCff).CffRec.iStartDate(1), llFlStartDate
                                        gUnpackDateLong tmCfflist(ilCff).CffRec.iEndDate(0), tmCfflist(ilCff).CffRec.iEndDate(1), llFlEndDate
                                        If (llFlStartDate <= llFlEndDate) And (llDate >= llFlStartDate) And (llDate <= llFlEndDate) Then
                                            Select Case tmCfflist(ilCff).CffRec.sPriceType
                                                Case "T"    'True
                                                   lmCntrMonthGross(ilMonth) = lmCntrMonthGross(ilMonth) - tmCfflist(ilCff).CffRec.lActPrice
                                            End Select
                                            Exit For
                                        End If
                                    End If
                                Next ilCff
                                Exit For
                            End If
                        Next ilMonth
                    End If
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSchDate
                    If (llSchDate >= llStartDate) And (llSchDate <= llEndDate) Then
                        For ilMonth = 0 To UBound(tmInvMonthEndDate) - 1 Step 1
                            'If lmInvMonthEndDate(ilMonth) = gDateValue(gObtainEndStd(Format(llSchDate, "m/d/yy"))) Then
                            If ((tmInvMonthEndDate(ilMonth).lStdMonthEndDate = gDateValue(gObtainEndStd(Format(llSchDate, "m/d/yy")))) And (tmChf.sBillCycle = "S")) Or ((tmInvMonthEndDate(ilMonth).lCalMonthEndDate = gDateValue(gObtainEndCal(Format(llSchDate, "m/d/yy")))) And (tmChf.sBillCycle = "C")) Then
                                For ilCff = LBound(tmCfflist) To UBound(tmCfflist) - 1 Step 1
                                    If tmClfList(ilClf).ClfRec.iLine = tmCfflist(ilCff).CffRec.iClfLine Then
                                        gUnpackDateLong tmCfflist(ilCff).CffRec.iStartDate(0), tmCfflist(ilCff).CffRec.iStartDate(1), llFlStartDate
                                        gUnpackDateLong tmCfflist(ilCff).CffRec.iEndDate(0), tmCfflist(ilCff).CffRec.iEndDate(1), llFlEndDate
                                        If (llFlStartDate <= llFlEndDate) And (llDate >= llFlStartDate) And (llDate <= llFlEndDate) Then
                                            Select Case tmCfflist(ilCff).CffRec.sPriceType
                                                Case "T"    'True
                                                   lmCntrMonthGross(ilMonth) = lmCntrMonthGross(ilMonth) + tmCfflist(ilCff).CffRec.lActPrice
                                            End Select
                                            Exit For
                                        End If
                                    End If
                                Next ilCff
                                Exit For
                            End If
                        Next ilMonth
                    End If
                End If
                ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Loop
        End If
    Next ilClf
End Function


Private Function mGetReceivableDollars(hlFile As Integer) As Integer
    Dim ilMonth As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim slStr As String
    Dim slDate As String
    
    mGetReceivableDollars = True
    For ilMonth = 0 To UBound(tmInvMonthEndDate) - 1 Step 1
        tmRvfSrchKey4.lCntrNo = tmChf.lCntrNo
        If tmChf.sBillCycle = "C" Then
            gPackDateLong gDateValue(gObtainStartCal(Format(tmInvMonthEndDate(ilMonth).lCalMonthEndDate, "ddddd"))), tmRvfSrchKey4.iTranDate(0), tmRvfSrchKey4.iTranDate(1)
        Else
            gPackDateLong gDateValue(gObtainStartStd(Format(tmInvMonthEndDate(ilMonth).lStdMonthEndDate, "ddddd"))), tmRvfSrchKey4.iTranDate(0), tmRvfSrchKey4.iTranDate(1)
        End If
        ilRet = btrGetGreaterOrEqual(hlFile, tmRvf, imRvfRecLen, tmRvfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmRvf.lCntrNo = tmChf.lCntrNo)
            'gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slDate
            'llDate = gDateValue(gObtainEndStd(slDate))
            If ((gDateValue(gObtainEndStd(slDate)) = tmInvMonthEndDate(ilMonth).lStdMonthEndDate) And (tmChf.sBillCycle = "S")) Or ((gDateValue(gObtainEndCal(slDate)) = tmInvMonthEndDate(ilMonth).lCalMonthEndDate) And (tmChf.sBillCycle = "C")) Then
                If (tmRvf.sTranType = "IN") And (tmRvf.sCashTrade = "C") And (tmRvf.iMnfItem = 0) Then
                    gPDNToStr tmRvf.sGross, 2, slStr
                    lmRecMonthGross(ilMonth) = lmRecMonthGross(ilMonth) + gStrDecToLong(slStr, 2)
                    gPDNToStr tmRvf.sNet, 2, slStr
                    lmRecMonthNet(ilMonth) = lmRecMonthNet(ilMonth) + gStrDecToLong(slStr, 2)
                End If
            Else
                Exit Do
            End If
            ilRet = btrGetNext(hlFile, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next ilMonth
    
End Function


Private Function mGetReceivableCntr(hlFile As Integer, llStartDate As Long, llEndDate As Long) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim slStr As String
    Dim ilFound As Integer
    Dim ilChf As Integer
    
    mGetReceivableCntr = True
    gPackDateLong llStartDate, tmRvfSrchKey3.iTranDate(0), tmRvfSrchKey3.iTranDate(1)
    ilRet = btrGetGreaterOrEqual(hlFile, tmRvf, imRvfRecLen, tmRvfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE)
        gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
        If (llDate <= llEndDate) Then
            If (tmRvf.sTranType = "IN") And (tmRvf.sCashTrade = "C") And (tmRvf.iMnfItem = 0) Then
                For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
                    If tmRvf.lCntrNo = tmChfAdvtExt(ilChf).lCntrNo Then
                        ilFound = True
                        Exit For
                    End If
                Next ilChf
                If Not ilFound Then
                    'Get contract code
                    tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F")
                        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus = "F") And (tmChf.sInstallDefined <> "Y") Then
                        tmChfAdvtExt(UBound(tmChfAdvtExt)).lCode = tmChf.lCode
                        tmChfAdvtExt(UBound(tmChfAdvtExt)).lCntrNo = tmRvf.lCntrNo
                        ReDim Preserve tmChfAdvtExt(LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) + 1) As CHFADVTEXT
                    End If
                End If
            End If
        Else
            Exit Do
        End If
        ilRet = btrGetNext(hlFile, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
End Function


Private Sub mGetAdServerDollars()
    Dim ilPass As Integer
    Dim ilRet As Integer
    Dim slPCFStartDate As String
    Dim slPCFEndDate As String
    Dim ilMonth As Integer
    Dim ilInvMonth As Integer
    Dim ilInvYear As Integer
    Dim llGross As Long
        
    For ilPass = 0 To 1 Step 1
        tmPcfSrchKey5.lChfCode = tmChf.lCode
        tmPcfSrchKey5.sType = "S"
        If ilPass = 1 Then
            tmPcfSrchKey5.sType = "H"
        End If
        ilRet = btrGetEqual(hmPcf, tmPcf, imPcfRecLen, tmPcfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmPcf.lChfCode = tmChf.lCode) And (tmPcfSrchKey5.sType = tmPcf.sType)
            
            gUnpackDate tmPcf.iStartDate(0), tmPcf.iStartDate(1), slPCFStartDate
            gUnpackDate tmPcf.iEndDate(0), tmPcf.iEndDate(1), slPCFEndDate
            For ilMonth = 0 To UBound(tmInvMonthEndDate) - 1 Step 1
                If (gDateValue(gObtainEndStd(slPCFEndDate)) >= tmInvMonthEndDate(ilMonth).lStdMonthEndDate And (gDateValue(gObtainEndStd(slPCFStartDate)) >= tmInvMonthEndDate(ilMonth).lStdMonthEndDate) And (tmChf.sBillCycle = "S")) _
                    Or (gDateValue(gObtainEndCal(slPCFEndDate)) >= tmInvMonthEndDate(ilMonth).lCalMonthEndDate And (gDateValue(gObtainEndCal(slPCFStartDate)) >= tmInvMonthEndDate(ilMonth).lCalMonthEndDate) And (tmChf.sBillCycle = "S")) Then
                    If tmPcf.sPriceType <> "F" Then
                        ilInvMonth = tmInvMonthEndDate(ilMonth).iMonth      ' Month(Format(lmInvMonthEndDate(ilMonth), "ddddd"))
                        ilInvYear = tmInvMonthEndDate(ilMonth).iYear        'Year(Format(lmInvMonthEndDate(ilMonth), "ddddd"))
                        tmIbfSrchKey1.lCntrNo = tmChf.lCntrNo
                        tmIbfSrchKey1.iPodCPMID = tmPcf.iPodCPMID
                        ilRet = btrGetEqual(hmIbf, tmIbf, imIbfRecLen, tmIbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmIbf.lCntrNo = tmChf.lCntrNo) And (tmIbf.iPodCPMID = tmPcf.iPodCPMID)
                            If (tmIbf.iBillMonth = ilInvMonth) And (tmIbf.iBillYear = ilInvYear) Then
                                Exit Do
                            End If
                            ilRet = btrGetNext(hmIbf, tmIbf, imIbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If (ilRet <> BTRV_ERR_NONE) Or (tmIbf.lCntrNo <> tmChf.lCntrNo) Or (tmIbf.iPodCPMID <> tmPcf.iPodCPMID) Or (tmIbf.iBillMonth <> ilInvMonth) Or (tmIbf.iBillYear <> ilInvYear) Then
                            llGross = 0
                        Else
                            llGross = tmIbf.lImpressions * tmPcf.lPodCPM / 1000
                        End If
                    Else
                        llGross = tmPcf.lTotalCost / mNumberOfMonths(slPCFStartDate, slPCFEndDate)
                    End If
                    lmCntrMonthGross(ilMonth) = lmCntrMonthGross(ilMonth) + llGross
                End If
            Next ilMonth
            ilRet = btrGetNext(hmPcf, tmPcf, imPcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next ilPass
End Sub

Private Function mNumberOfMonths(slStartDate As String, slEndDate As String) As Integer
    Dim ilStartMonth As Integer
    Dim ilStartYear As Integer
    Dim ilEndMonth As Integer
    Dim ilEndYear As Integer
    'ilStartMonth = Month(slStartDate)
    'ilStartYear = Year(slStartDate)
    'ilEndMonth = Month(slEndDate)
    'ilEndYear = Year(slEndDate)
    'If ilStartYear = ilEndYear Then
    '    mNumberOfMonths = ilEndMonth - ilStartMonth + 1
    'Else
    '    mNumberOfMonths = (13 - ilStartMonth) + ilEndMonth + 12 * (ilEndYear - ilStartYear - 1)
    'End If
    mNumberOfMonths = 0
    If tmChf.sBillCycle = "C" Then   'Cal
        ilStartMonth = Month(gDateValue(gObtainStartCal(slStartDate)))
        ilStartYear = Year(gDateValue(gObtainStartCal(slStartDate)))
        gObtainMonthYear 1, slEndDate, ilEndMonth, ilEndYear
    ElseIf tmChf.sBillCycle = "W" Then   'Week
        Exit Function
    Else
        ilStartMonth = Month(gDateValue(gObtainEndStd(slStartDate)))
        ilStartYear = Year(gDateValue(gObtainEndStd(slStartDate)))
        gObtainMonthYear 0, slEndDate, ilEndMonth, ilEndYear
    End If
    If ilStartYear = ilEndYear Then
        mNumberOfMonths = ilEndMonth - ilStartMonth + 1
    Else
        mNumberOfMonths = (13 - ilStartMonth) + ilEndMonth + 12 * (ilEndYear - ilStartYear - 1)
    End If

End Function

Private Sub mPopMonth()
    cbcStartMonth.Clear
    cbcStartMonth.AddItem "January"
    cbcStartMonth.SetItemData = 1
    cbcStartMonth.AddItem "February"
    cbcStartMonth.SetItemData = 2
    cbcStartMonth.AddItem "March"
    cbcStartMonth.SetItemData = 3
    cbcStartMonth.AddItem "April"
    cbcStartMonth.SetItemData = 4
    cbcStartMonth.AddItem "May"
    cbcStartMonth.SetItemData = 5
    cbcStartMonth.AddItem "June"
    cbcStartMonth.SetItemData = 6
    cbcStartMonth.AddItem "July"
    cbcStartMonth.SetItemData = 7
    cbcStartMonth.AddItem "August"
    cbcStartMonth.SetItemData = 8
    cbcStartMonth.AddItem "September"
    cbcStartMonth.SetItemData = 9
    cbcStartMonth.AddItem "October"
    cbcStartMonth.SetItemData = 10
    cbcStartMonth.AddItem "November"
    cbcStartMonth.SetItemData = 11
    cbcStartMonth.AddItem "December"
    cbcStartMonth.SetItemData = 12

    cbcEndMonth.Clear
    cbcEndMonth.AddItem "January"
    cbcEndMonth.SetItemData = 1
    cbcEndMonth.AddItem "February"
    cbcEndMonth.SetItemData = 2
    cbcEndMonth.AddItem "March"
    cbcEndMonth.SetItemData = 3
    cbcEndMonth.AddItem "April"
    cbcEndMonth.SetItemData = 4
    cbcEndMonth.AddItem "May"
    cbcEndMonth.SetItemData = 5
    cbcEndMonth.AddItem "June"
    cbcEndMonth.SetItemData = 6
    cbcEndMonth.AddItem "July"
    cbcEndMonth.SetItemData = 7
    cbcEndMonth.AddItem "August"
    cbcEndMonth.SetItemData = 8
    cbcEndMonth.AddItem "September"
    cbcEndMonth.SetItemData = 9
    cbcEndMonth.AddItem "October"
    cbcEndMonth.SetItemData = 10
    cbcEndMonth.AddItem "November"
    cbcEndMonth.SetItemData = 11
    cbcEndMonth.AddItem "December"
    cbcEndMonth.SetItemData = 12

End Sub
