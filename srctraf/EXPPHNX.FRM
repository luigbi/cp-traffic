VERSION 5.00
Begin VB.Form ExpPhnx 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4395
   ClientLeft      =   225
   ClientTop       =   1620
   ClientWidth     =   9135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4395
   ScaleWidth      =   9135
   Begin VB.ListBox lbcMsg 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   3720
      MultiSelect     =   2  'Extended
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   780
      Width           =   5235
   End
   Begin VB.CommandButton cmcStartDate 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2085
      Picture         =   "EXPPHNX.frx":0000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   390
      Width           =   195
   End
   Begin VB.TextBox edcStartDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1155
      MaxLength       =   10
      TabIndex        =   2
      Top             =   390
      Width           =   930
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1155
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "EXPPHNX.frx":00FA
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   17
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   5
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   45
      ScaleHeight     =   270
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3135
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2370
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2790
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   165
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   780
      Width           =   3375
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3345
      TabIndex        =   7
      Top             =   3990
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4785
      TabIndex        =   8
      Top             =   3990
      Width           =   1050
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   3900
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   195
      TabIndex        =   19
      Top             =   3690
      Width           =   8730
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Export Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   75
      TabIndex        =   1
      Top             =   375
      Width           =   1065
   End
   Begin VB.Label lacErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5055
      TabIndex        =   14
      Top             =   1800
      Width           =   1725
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   5175
      TabIndex        =   12
      Top             =   1395
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3225
      TabIndex        =   10
      Top             =   1395
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "ExpPhnx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EXPPHNX.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpPhnx.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export feed (for Phoenix) input screen code
Option Explicit
Option Compare Text
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmTo As Integer   'From file hanle
Dim hmMsg As Integer   'From file hanle
Dim lmNowDate As Long   'Todays date
'Required by gMakeSsf
Dim tmSsf As SSF                'SSF record image
Dim hmSsf As Integer
Dim hmCTSsf As Integer
Dim tmCTSsf As SSF               'Ssf for conflict test
'Dim tmSsfOld As SSF
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmAvailTest As AVAILSS
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0 'ANF key record image
Dim imAdfRecLen As Integer  'ANF record length
'Agency name
Dim hmAgf As Integer
Dim tmAgf As AGF
Dim imAgfRecLen As Integer  'AGF record length
'Multi- name
Dim hmMnf As Integer
Dim tmMnf As MNF
Dim tmMnfSrchKey As INTKEY0 'MNF key record image
Dim imMnfRecLen As Integer  'MNF record length
'Copy/Product
Dim hmCpf As Integer
Dim tmCpf As CPF
Dim tmCpfSrchKey As LONGKEY0 'CPEF key record image
Dim imCpfRecLen As Integer  'CPF record length
'Avail name
Dim hmAnf As Integer
Dim tmAnf As ANF
Dim tmAnfSrchKey As INTKEY0 'ANF key record image
Dim imAnfRecLen As Integer  'ANF record length
'Contract record information
Dim hmCHF As Integer        'Contract header file handle
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
'Contract Line record information
Dim hmClf As Integer        'Contract Line file handle
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image
'Short Title Vehicle Table record information
Dim hmVsf As Integer        'Short Title Vehicle Table file handle
Dim imVsfRecLen As Integer  'VSF record length
Dim tmVsf As VSF            'VSF record image
'Short Title record information
Dim hmSif As Integer        'Short Title file handle
Dim imSifRecLen As Integer  'SIF record length
Dim tmSif As SIF            'SIF record image
Dim tmSifSrchKey As LONGKEY0
'Regional Scheduled Copy record information
Dim hmRsf As Integer        'Regional Scheduled Copy file handle
Dim imRsfRecLen As Integer  'RSF record length
Dim tmRsf As RSF            'RSF record image
'Blackout record information
Dim hmBof As Integer        'Blackout file handle
Dim imBofRecLen As Integer  'BOF record length
Dim tmBof As BOF            'BOF record image
'Media code record information
Dim hmMcf As Integer        'Contract line file handle
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer  'MCF record length
Dim tmMcf As MCF            'MCF record image
'Copy inventory record information
Dim hmCif As Integer        'Copy line file handle
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer  'CIF record length
Dim tmCif As CIF            'CIF record image
Dim hmTzf As Integer        'Time zone Copy file handle
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfRecLen As Integer  'TZF record length
Dim tmTzf As TZF            'TZF record image
Dim imCopyMissing As Integer
' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmSVef As VEF           'Selling Vehicle
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer     'VEF record length
Dim smVehName As String
'Vehicle linkage record information
Dim hmVlf As Integer        'Vehicle linkage file handle
Dim tmVlfSrchKey1 As VLFKEY1 'VLF key record image
Dim imVlfRecLen As Integer  'VLF record length
Dim tmVlf As VLF            'VLF record image
'Delivery file (DLF)
Dim hmDlf As Integer        'Delivery link file
Dim imDlfRecLen As Integer  'DLF record length
Dim tmDlfSrchKey As DLFKEY0 'DLF key record image
Dim tmDlf As DLF            'DLF record image
'Spot record
Dim tmSdf As SDF
Dim hmSdf As Integer
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey3 As LONGKEY0
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imEvtType(0 To 14) As Integer
Dim tmSpotTimes() As SPOTTIMES
Dim smCurrentLines() As String
Dim smNewLines() As String
Dim imFirstActivate As Integer
' MsgBox parameters
Const vbOkOnly = 0                 ' OK button only
Const vbCritical = 16          ' Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcExport_Click()
    Dim slToFile As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim slStr As String
    Dim ilDays As Integer
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slLetter As String
    Dim slFileName As String
    Dim ilLine As Integer
    Dim slDateTime As String
    Dim slMissingCopyNames As String
    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        Beep
        edcStartDate.SetFocus
        Exit Sub
    End If
    slStartDate = slStr
    ilDays = 1  'Val(edcNoDays.Text)
    'If ilDays = 0 Then
    '    Beep
    '    edcNoDays.SetFocus
    '    Exit Sub
    'End If
    slEndDate = Format$(gDateValue(slStartDate) + ilDays - 1, "m/d/yy")
    lbcMsg.Clear
    If Not mOpenMsgFile() Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    'Determine file name from date: YYMMDDn.NY
    'where n= version letter
    gObtainYearMonthDayStr slStartDate, True, slFYear, slFMonth, slFDay
    slFileName = right$(slFYear, 2) & slFMonth & slFDay
    slLetter = "CS"
    ilRet = 0
    'On Error GoTo cmcExportErr:
    slToFile = slFileName & slLetter & ".PHX"
    'slDateTime = FileDateTime(sgExportPath & slToFile)
    ilRet = gFileExist(sgExportPath & slToFile)
    If ilRet = 0 Then
        Kill sgExportPath & slToFile
    End If
    On Error GoTo 0
    slToFile = sgExportPath & slFileName & slLetter & ".PHX" 'slToFile
    ReDim smCurrentLines(0 To 0) As String
    On Error GoTo 0
    
    gAutomationAlertAndLogHandler "* StartDate = " & edcStartDate.Text
    
    
    ilRet = 0
    'On Error GoTo cmcExportErr:
    'hmTo = FreeFile
    'Open slToFile For Output As hmTo
    ilRet = gFileOpen(slToFile, "Output", hmTo)
    If ilRet <> 0 Then
        'Print #hmMsg, "** Terminated **"
        gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        'edcTo.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ReDim smNewLines(0 To 0) As String
    imExporting = True
    slMissingCopyNames = ""
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            imCopyMissing = False
            slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            smVehName = Trim$(slName)
            'Print #hmMsg, "Generating Data for " & Trim$(slName)
            gAutomationAlertAndLogHandler "Generating Data for " & Trim$(slName)
            lacProcessing.Caption = "Generating Data for " & Trim$(slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            If Not mExpPhnxDay("C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType()) Then
                'Print #hmMsg, "** Terminated **"
                gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                Close #hmMsg
                Close #hmTo
                imExporting = False
                'MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                Screen.MousePointer = vbDefault
                cmcCancel.SetFocus
                Exit Sub
            End If
            If imCopyMissing Then
                If slMissingCopyNames = "" Then
                    slMissingCopyNames = slName
                Else
                    slMissingCopyNames = slMissingCopyNames & ", " & slName
                End If
            End If
        End If
    Next ilLoop
    lacProcessing.Caption = "Writing Data to " & slFileName & slLetter & ".TXT"
    DoEvents
    ilRet = 0
    'On Error GoTo cmcExportErr:
    For ilLine = LBound(smNewLines) To UBound(smNewLines) - 1 Step 1
        Print #hmTo, RTrim$(smNewLines(ilLine))
        If ilRet <> 0 Then
            imExporting = False
            'Print #hmMsg, "** Terminated **"
            gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Close #hmTo
            Screen.MousePointer = vbDefault
            ''MsgBox "Error writing to " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            gAutomationAlertAndLogHandler "Error writing to " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            cmcCancel.SetFocus
            Exit Sub
        End If
    Next ilLine
    'Print #hmMsg, "** Completed **"
    gAutomationAlertAndLogHandler "Export Completed"
    Close #hmMsg
    Close #hmTo
    On Error GoTo 0

    lacProcessing.Caption = ""
    Screen.MousePointer = vbDefault
    'If slMissingCopyNames <> "" Then
    '    MsgBox "Copy missing on " & slMissingCopyNames, vbOkOnly + vbExclamation + vbApplicationModal, "Copy Missing"
    'End If
    imExporting = False
    Exit Sub
'cmcExportErr:
'    ilRet = Err.Number
'    Resume Next
ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    

End Sub
Private Sub cmcExport_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub cmcStartDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcStartDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcStartDate_Change()
    Dim slStr As String
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcStartDate_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
        'mInitDDE
        'mSendHelpMsg "BT"
    End If
    gCtrlGotFocus edcStartDate
End Sub
Private Sub edcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcStartDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcStartDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
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
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
    End If

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    ilRet = btrClose(hmCTSsf)
    btrDestroy hmCTSsf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmRsf)
    btrDestroy hmRsf
    ilRet = btrClose(hmBof)
    btrDestroy hmBof
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmTzf)
    btrDestroy hmTzf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    ilRet = btrClose(hmDlf)
    btrDestroy hmDlf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    
    Set ExpPhnx = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcVehicle_GotFocus()
    plcCalendar.Visible = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = edcStartDate.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mExpPhnxDay                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build export record             *
'*                     ("EST" zone, Cmml Sch = "Y",    *
'*                      and Avail Name = "N")          *
'*                                                     *
'*******************************************************
Private Function mExpPhnxDay(sLCP As String, ilVefCode As Integer, slSDate As String, slEDate As String, slStartTime As String, slEndTime As String, ilEvtType() As Integer) As Integer
'
'   ilRet = mExpPhnxDay (slCp, ilVefCode, slSDate, slEDate, slStartTime, slEndTime, ilEvtType())
'
'   Where:
'       slCP (I)- "C"=Current only; "P"=Pending only; "B"=Both
'       ilVefCode (I)-Vehicle code number(slFor = L or C) or feed code (slFor = D)
'       slSDate (I)- Start Date that events are to be obtained
'       slEDate (I)- Start Date that events are to be obtained
'       slStartTime (I)- Start Time (included)
'       slEndTime (I)- End time (not included)
'       ilEvtType (I)- Array of which events are to be included (True or False)
'                       Index description
'                         0   Library
'                         1   Program event
'                         2   Contract avail
'                         3   Open BB
'                         4   Floating BB
'                         5   Close BB
'                         6   Cmml promo
'                         7   Feed avail
'                         8   PSA avail
'                         9   Promo avail
'                         10  Page eject
'                         11  Line space 1
'                         12  Line space 2
'                         13  Line space 3
'                         14  Other event types
'
'
'
    'Dim slType As String
    Dim ilType As Integer
    Dim ilRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilSsfDate0 As Integer
    Dim ilSsfDate1 As Integer
    Dim ilEvt As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim ilSpot As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilTerminated As Integer
    Dim slTime As String
    Dim slDayName As String
    Dim ilVpfIndex As Integer
    Dim ilVehCode As Integer
    Dim ilDlfDate0 As Integer
    Dim ilDlfDate1 As Integer
    Dim ilDlfFound As Integer
    Dim ilVlfDate0 As Integer
    Dim ilVlfDate1 As Integer
    Dim ilAirHour As Integer
    Dim ilLocalHour As Integer
    Dim ilSIndex As Integer
    Dim slSsfDate As String
    'Spot summary
    Dim tlSsfSrchKey As SSFKEY0 'SSF key record image
    Dim ilSsfRecLen As Integer  'SSF record length
    Dim llEvtTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilWithinTime As Integer
    Dim ilEvtFdIndex As Integer
    Dim slAirDate As String
    Dim llAirDate As Long
    Dim ilTest As Integer
    Dim llSpotTime As Long
    Dim llHour As Long
    Dim llAvailTime As Long
    Dim llPrevTime As Long
    Dim ilCopy As Integer
    Dim ilDlfExist As Integer
    Dim llCopyMissingSdfCode As Long
    'Spot detail record information
    Dim slStr As String
    Dim slMsg As String
    Dim ilAvailUnits As Integer
    Dim ilAvailLen As Integer
    Dim ilSpotLen As Integer
    Dim ilSpotUnits As Integer
    Dim ilSubFeed As Integer    'True=Output only Dlf records with subfeeds;
                                'False=Output only records without subfeeds
    Dim slZone As String
    Dim slShortTitle As String
    Dim slAdvtName As String
    Dim slCodeStn As String
    Dim slCart As String * 6
    Dim slAdvtProd As String * 45   '52
    Dim slLen As String
    Dim slLine As String
    Dim slHeader As String
    Dim ilLineNo As Integer
    Dim ilPageNo As Integer
    Dim ilShowTime As Integer
    Dim llPrevHour As Long
    Dim ilNoLines As Integer
    ReDim tmSpotTimes(0 To 0) As SPOTTIMES
    'ReDim smNewLines(0 To 0) As String * 80
    Dim ilBlackoutDefined As Integer
    Dim llBChfCode As Long
    Dim llBCifCode As Long
    Dim llBSifCode As Long
    Dim llBBofCode As Long
    Dim ilVff As Integer
    Dim blBypassZeroUnits As Boolean


    ilLineNo = 66  'Force header
    ilPageNo = 0
    'slType = "O"
    ilType = 0
    llStartTime = CLng(gTimeToCurrency(slStartTime, False))
    llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
    ilSsfRecLen = Len(tmSsf)  'Get and save SSF record length
    On Error GoTo 0
    tmVef.iCode = 0
    llSDate = gDateValue(slSDate)
    llEDate = gDateValue(slEDate)
    For llDate = llSDate To llEDate + 1 Step 1  'Process next date for avails that map back one day
        ilWithinTime = False
        slDate = Format$(llDate, "m/d/yy")
        ilDay = gWeekDayStr(slDate)
        gPackDate slDate, ilLogDate0, ilLogDate1
        ilVehCode = ilVefCode
        If ilVehCode <> tmVef.iCode Then
            tmVefSrchKey.iCode = ilVehCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If (ilRet <> BTRV_ERR_NONE) Then
                Screen.MousePointer = vbDefault
                mExpPhnxDay = False
                ''MsgBox smVehName & " get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                gAutomationAlertAndLogHandler smVehName & " get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                Exit Function
                Screen.MousePointer = vbHourglass
            End If
            ilVpfIndex = -1
            'For ilLoop = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
            '    If tmVef.iCode = tgVpf(ilLoop).iVefKCode Then
                ilLoop = gBinarySearchVpf(tmVef.iCode)
                If ilLoop <> -1 Then
                    ilVpfIndex = ilLoop
            '        Exit For
                End If
            'Next ilLoop
            If ilVpfIndex = -1 Then
                Screen.MousePointer = vbDefault
                mExpPhnxDay = False
                ''MsgBox smVehName & " preference get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                gAutomationAlertAndLogHandler smVehName & " preference get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                Screen.MousePointer = vbHourglass
                Exit Function
            End If
        End If
        ilDlfFound = False
        'If (tmVef.sType = "A") Or ((tmVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(1) <> 0)) Then
        If (tmVef.sType = "A") Or ((tmVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(0) <> 0)) Then
            'Obtain Engineering records for date
            If (ilDay >= 0) And (ilDay <= 4) Then
                slDay = "0"
            ElseIf ilDay = 5 Then
                slDay = "6"
            Else
                slDay = "7"
            End If
            'Obtain the start date of Dlf
            tmDlfSrchKey.iVefCode = ilVehCode
            tmDlfSrchKey.sAirDay = slDay
            tmDlfSrchKey.iStartDate(0) = ilLogDate0
            tmDlfSrchKey.iStartDate(1) = ilLogDate1
            tmDlfSrchKey.iAirTime(0) = 0
            tmDlfSrchKey.iAirTime(1) = 6144 '24*256
            ilRet = btrGetLessOrEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) Then
                ilDlfDate0 = tmDlf.iStartDate(0)
                ilDlfDate1 = tmDlf.iStartDate(1)
                ilDlfFound = True
            Else
                ilDlfDate0 = 0
                ilDlfDate1 = 0
            End If
            'Obtain the start date of VLF
            If tmVef.sType = "A" Then
                ilVlfDate0 = 0
                ilVlfDate1 = 0
                tmVlfSrchKey1.iAirCode = ilVehCode
                tmVlfSrchKey1.iAirDay = Val(slDay)
                tmVlfSrchKey1.iEffDate(0) = ilLogDate0
                tmVlfSrchKey1.iEffDate(1) = ilLogDate1
                tmVlfSrchKey1.iAirTime(0) = 0
                tmVlfSrchKey1.iAirTime(1) = 6144    '24*256
                tmVlfSrchKey1.iAirPosNo = 32000
                tmVlfSrchKey1.iAirSeq = 32000
                ilRet = btrGetLessOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode)
                    ilTerminated = False
                    'Check for CBS
                    If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                        If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                            ilTerminated = True
                        End If
                    End If
                    If (tmVlf.sStatus <> "P") And (tmVlf.iAirDay = Val(slDay)) And (Not ilTerminated) Then
                        ilVlfDate0 = tmVlf.iEffDate(0)
                        ilVlfDate1 = tmVlf.iEffDate(1)
                        Exit Do
                    End If
                    ilRet = btrGetPrevious(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                'If (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) And (tmVlf.iAirDay = Val(slDay)) Then
                '    ilVlfDate0 = tmVlf.iEffDate(0)
                '    ilVlfDate1 = tmVlf.iEffDate(1)
                'Else
                '    ilVlfDate0 = 0
                '    ilVlfDate1 = 0
                'End If
            End If
        End If
        
        blBypassZeroUnits = False
        If (tmVef.sType = "A") Then
            ilVff = gBinarySearchVff(tmVef.iCode)
            If ilVff <> -1 Then
                If tgVff(ilVff).sHonorZeroUnits = "Y" Then
                    blBypassZeroUnits = True
                End If
            End If
        End If
        
        'If Not ilDlfFound Then
        '    Screen.MousePointer = vbDefault
        '    MsgBox smVehName & " delivery missing", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
        '    mExpPhnxDay = False
        '    Screen.MousePointer = vbHourGlass
        '    Exit Function
        'End If
        'gObtainVlf hlVlf, ilVehCode, llDate, tlVlf0(), tlVlf5(), tlVlf6()
        ilDay = gWeekDayStr(slDate)
        gPackDate slDate, ilLogDate0, ilLogDate1
        ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
        ilSsfDate0 = ilLogDate0
        ilSsfDate1 = ilLogDate1
        'tlSsfSrchKey.sType = slType
        tlSsfSrchKey.iType = ilType
        tlSsfSrchKey.iVefCode = ilVehCode
        tlSsfSrchKey.iDate(0) = ilSsfDate0
        tlSsfSrchKey.iDate(1) = ilSsfDate1
        tlSsfSrchKey.iStartTime(0) = 0
        tlSsfSrchKey.iStartTime(1) = 0
        ilRet = gSSFGetEqual(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        'If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.sType <> slType) Or (tmSsf.iVefCode <> ilVehCode) Or (tmSsf.iDate(0) <> ilSsfDate0) Or (tmSsf.iDate(1) <> ilSsfDate1) Then
        If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilType) Or (tmSsf.iVefCode <> ilVehCode) Or (tmSsf.iDate(0) <> ilSsfDate0) Or (tmSsf.iDate(1) <> ilSsfDate1) Then
            'If airing- then use first Ssf prior to date defined
            If tmVef.sType = "A" Then
                ilSsfDate0 = 0
                ilSsfDate1 = 0
                ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                tlSsfSrchKey.iType = ilType
                tlSsfSrchKey.iVefCode = ilVehCode
                tlSsfSrchKey.iDate(0) = ilLogDate0
                tlSsfSrchKey.iDate(1) = ilLogDate1
                tlSsfSrchKey.iStartTime(0) = 0
                tlSsfSrchKey.iStartTime(1) = 6144   '24*256
                ilRet = gSSFGetLessOrEqual(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefCode = ilVehCode)
                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode)
                    gUnpackDate tmSsf.iDate(0), tmSsf.iDate(1), slSsfDate
                    If (ilDay = gWeekDayStr(slSsfDate)) And (tmSsf.iStartTime(0) = 0) And (tmSsf.iStartTime(1) = 0) Then
                        ilSsfDate0 = tmSsf.iDate(0)
                        ilSsfDate1 = tmSsf.iDate(1)
                        Exit Do
                    End If
                    ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetPrevious(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
        End If
        DoEvents
        'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefCode = ilVehCode) Then
        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode) Then
            gUnpackDate ilSsfDate0, ilSsfDate1, slSsfDate
            'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefCode = ilVehCode) And (tmSsf.iDate(0) = ilSsfDate0) And (tmSsf.iDate(1) = ilSsfDate1)
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode) And (tmSsf.iDate(0) = ilSsfDate0) And (tmSsf.iDate(1) = ilSsfDate1)
                'Loop thru Ssf and move records to export
                ilEvt = 1
                Do While ilEvt <= tmSsf.iCount
                   LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If (tmProg.iRecType = 1) Or ((tmProg.iRecType >= 2) And (tmProg.iRecType <= 9)) Then
                        gUnpackTimeLong tmProg.iStartTime(0), tmProg.iStartTime(1), False, llEvtTime
                        If llEvtTime > llEndTime Then
                            ilWithinTime = False
                            Exit Do
                        End If
                        If llEvtTime >= llStartTime Then
                            ilWithinTime = True
                        End If
                    End If
                    If blBypassZeroUnits And (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        If (tmAvail.iAvInfo And &H1F <= 0) Or (tmAvail.iLen <= 0) Then
                            tmProg.iRecType = 0
                        End If
                    End If
                    ilEvtFdIndex = -1
                    If ilWithinTime Then
                        If tmProg.iRecType = 1 Then    'Program
                            ilAirHour = tmProg.iStartTime(1) \ 256  'Obtain Hour
                            If (llDate > llEDate) And (ilAirHour >= 6) Then
                                'Done processing
                                mExpPhnxDay = True
                                Exit Function
                            End If
                        ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then 'Avail
                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                            ilAirHour = tmAvail.iTime(1) \ 256  'Obtain Hour
                            If (llDate > llEDate) And (ilAirHour >= 6) Then
                                'Done processing
                                mExpPhnxDay = True
                                Exit Function
                            End If
                            ilSubFeed = False
                            ilDlfExist = False
                            If (tmVef.sType = "A") Or (ilDlfFound) Then
                                'Obtain engineering entry to see if avail is sent
                                tmDlfSrchKey.iVefCode = ilVehCode
                                tmDlfSrchKey.sAirDay = slDay
                                tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                    ilDlfExist = True
                                    ilTerminated = False
                                    If (tmDlf.sFed = "N") Or (tmDlf.sZone <> "CST") Then
                                        ilTerminated = True
                                    Else
                                        If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                            If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                ilTerminated = True
                                            End If
                                        End If
                                    End If
                                    If Not ilTerminated Then
                                        If tmDlf.iMnfSubFeed > 0 Then
                                            ilSubFeed = True
                                            Exit Do
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            Else
                                ilDlfExist = True   'Set to aviod error message that link is missing
                            End If
                            'If Not ilDlfExist Then
                            If (Not ilDlfExist) And (tgSpf.sSDelNet = "Y") Then
                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
'                                ilRet = MsgBox(smVehName & " missing Delivery link on " & slDate & " at " & slTime, vbOKCancel + vbQuestion + vbApplicationModal, "Find Error")
'                                If ilRet = 2 Then   'Cancel
'                                    mExpPhnxDay = False
'                                    Exit Function
'                                End If

                                gAutomationAlertAndLogHandler smVehName & " missing Delivery link on " & slDate & " at " & slTime 'log
                                If igExportType < 1 Then
                                    gAutomationAlertAndLogHandler smVehName & " missing Delivery link on " & slDate & " at " & slTime 'log
                                    ilRet = MsgBox(smVehName & " missing Delivery link on " & slDate & " at " & slTime, vbOKCancel + vbQuestion + vbApplicationModal, "Find Error")
                                    If ilRet = 2 Then   'Cancel
                                        mExpPhnxDay = False
                                        Exit Function
                                    End If
                                End If
                                
                            Else
                                'Loop on spots, then add conflicting spots
                                If (tmVef.sType = "A") Then
                                    tmVlfSrchKey1.iAirCode = ilVehCode
                                    tmVlfSrchKey1.iAirDay = Val(slDay)
                                    tmVlfSrchKey1.iEffDate(0) = ilVlfDate0
                                    tmVlfSrchKey1.iEffDate(1) = ilVlfDate1
                                    tmVlfSrchKey1.iAirTime(0) = tmAvail.iTime(0)
                                    tmVlfSrchKey1.iAirTime(1) = tmAvail.iTime(1)
                                    tmVlfSrchKey1.iAirPosNo = 0
                                    tmVlfSrchKey1.iAirSeq = 1
                                    ilRet = btrGetGreaterOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) And (tmVlf.iAirDay = Val(slDay)) And (tmVlf.iEffDate(0) = ilVlfDate0) And (tmVlf.iEffDate(1) = ilVlfDate1) And (tmVlf.iAirTime(0) = tmAvail.iTime(0)) And (tmVlf.iAirTime(1) = tmAvail.iTime(1))
                                        ilTerminated = False
                                        If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                            If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                                ilTerminated = True
                                            End If
                                        End If
                                        If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                            'If (tmCTSsf.sType <> slType) Or (tmCTSsf.iVefCode <> tmVlf.iSellCode) Or (tmCTSsf.iDate(0) <> ilLogDate0) Or (tmCTSsf.iDate(1) <> ilLogDate1) Then
                                            If (tmCTSsf.iType <> ilType) Or (tmCTSsf.iVefCode <> tmVlf.iSellCode) Or (tmCTSsf.iDate(0) <> ilLogDate0) Or (tmCTSsf.iDate(1) <> ilLogDate1) Then
                                                tmVefSrchKey.iCode = tmVlf.iSellCode
                                                ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                If (ilRet = BTRV_ERR_NONE) Then
                                                    ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                    'tlSsfSrchKey.sType = slType
                                                    tlSsfSrchKey.iVefCode = tmVlf.iSellCode
                                                    tlSsfSrchKey.iDate(0) = ilLogDate0
                                                    tlSsfSrchKey.iDate(1) = ilLogDate1
                                                    tlSsfSrchKey.iStartTime(0) = 0
                                                    tlSsfSrchKey.iStartTime(1) = 0
                                                    ilRet = gSSFGetEqual(hmCTSsf, tmCTSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                End If
                                            End If
                                            'Do While (ilRet = BTRV_ERR_NONE) And (tmCTSsf.sType = slType) And (tmCTSsf.iVefCode = tmVlf.iSellCode) And (tmCTSsf.iDate(0) = ilLogDate0) And (tmCTSsf.iDate(1) = ilLogDate1)
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmCTSsf.iType = ilType) And (tmCTSsf.iVefCode = tmVlf.iSellCode) And (tmCTSsf.iDate(0) = ilLogDate0) And (tmCTSsf.iDate(1) = ilLogDate1)
                                                For ilSIndex = 1 To tmCTSsf.iCount Step 1
                                                    tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSIndex)
                                                    If ((tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9)) Then
                                                        If (tmAvailTest.iTime(0) = tmVlf.iSellTime(0)) And (tmAvailTest.iTime(1) = tmVlf.iSellTime(1)) Then
                                                            'Determine if any unsold time
                                                            ilAvailUnits = tmAvailTest.iAvInfo And &H1F
                                                            ilAvailLen = tmAvailTest.iLen
                                                            ilSpotLen = 0
                                                            For ilSpot = 1 To tmAvailTest.iNoSpotsThis Step 1
                                                               LSet tmSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilSpot + ilSIndex)
                                                                If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                                                    ilSpotLen = ilSpotLen + (tmSpot.iPosLen And &HFFF)
                                                                    ilSpotUnits = ilSpotUnits + 1
                                                                End If
                                                            Next ilSpot
                                                            If ilAvailLen > ilSpotLen Then
                                                                If ilDlfFound Then
                                                                    tmDlfSrchKey.iVefCode = ilVehCode
                                                                    tmDlfSrchKey.sAirDay = slDay
                                                                    tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                    tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                    tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                    tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                    ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                    Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                                        ilTerminated = False
                                                                        If (tmDlf.sFed = "N") Or (StrComp(Trim$(tmDlf.sZone), "CST", 1) <> 0) Then
                                                                            ilTerminated = True
                                                                        Else
                                                                            If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                                If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                                    ilTerminated = True
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        If Not ilTerminated Then
                                                                            If ilSubFeed Then
                                                                                If tmDlf.iMnfSubFeed > 0 Then
                                                                                    '6/1/16: Replaced GoSub
                                                                                    'GoSub lProcAdjDate  'Result stored into slAirDate
                                                                                    mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour
                                                                                    If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                                        slMsg = "Unsold: " & Trim$(tmSVef.sName)
                                                                                        gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                                        slMsg = slMsg & " at " & slTime
                                                                                        tmAnfSrchKey.iCode = tmAvailTest.ianfCode
                                                                                        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                        If (ilRet = BTRV_ERR_NONE) Then
                                                                                            slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                                        End If
                                                                                        'Print #hmMsg, slMsg
                                                                                        gAutomationAlertAndLogHandler slMsg
                                                                                        lbcMsg.AddItem slMsg
                                                                                    End If
                                                                                    DoEvents
                                                                                    If ilRet <> 0 Then
                                                                                        mExpPhnxDay = False
                                                                                        Exit Function
                                                                                    End If
                                                                                    Exit Do
                                                                                End If
                                                                            Else
                                                                                If tmDlf.iMnfSubFeed = 0 Then
                                                                                    '6/1/16: Replaced GoSub
                                                                                    'GoSub lProcAdjDate  'Result stored into slAirDate
                                                                                    mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour
                                                                                    If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                                        slMsg = "Unsold: " & Trim$(tmSVef.sName)
                                                                                        gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                                        slMsg = slMsg & " at " & slTime
                                                                                        tmAnfSrchKey.iCode = tmAvailTest.ianfCode
                                                                                        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                        If (ilRet = BTRV_ERR_NONE) Then
                                                                                            slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                                        End If
                                                                                        'Print #hmMsg, slMsg
                                                                                        gAutomationAlertAndLogHandler slMsg
                                                                                        lbcMsg.AddItem slMsg
                                                                                    End If
                                                                                    DoEvents
                                                                                    If ilRet <> 0 Then
                                                                                        mExpPhnxDay = False
                                                                                        Exit Function
                                                                                    End If
                                                                                    Exit Do
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                    Loop
                                                                Else
                                                                    tmDlf.iAirTime(0) = tmAvail.iTime(0)
                                                                    tmDlf.iAirTime(1) = tmAvail.iTime(1)
                                                                    tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                    tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                    tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                    tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                    tmDlf.sZone = ""
                                                                    tmDlf.iEtfCode = 0
                                                                    tmDlf.iEnfCode = 0
                                                                    tmDlf.sProgCode = ""
                                                                    tmDlf.iMnfFeed = 0
                                                                    tmDlf.sBus = ""
                                                                    tmDlf.sSchedule = ""
                                                                    tmDlf.iMnfSubFeed = 0
                                                                    '6/1/16: Replaced GoSub
                                                                    'GoSub lProcAdjDate  'Result stored into slAirDate
                                                                    mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour
                                                                    If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                        slMsg = "Unsold: " & Trim$(tmSVef.sName)
                                                                        gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                        slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                        tmAnfSrchKey.iCode = tmAvailTest.ianfCode
                                                                        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                        If (ilRet = BTRV_ERR_NONE) Then
                                                                            slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                        End If
                                                                        'Print #hmMsg, slMsg
                                                                        gAutomationAlertAndLogHandler slMsg
                                                                        lbcMsg.AddItem slMsg
                                                                    End If
                                                                    DoEvents
                                                                    If ilRet <> 0 Then
                                                                        mExpPhnxDay = False
                                                                        Exit Function
                                                                    End If
                                                                End If
                                                            End If
                                                            For ilSpot = 1 To tmAvailTest.iNoSpotsThis Step 1
                                                               LSet tmSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilSpot + ilSIndex)
                                                                tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                If ilRet = BTRV_ERR_NONE Then
                                                                    tmChfSrchKey.lCode = tmSdf.lChfCode
                                                                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                                    If ilRet = BTRV_ERR_NONE Then
                                                                        tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                                        If ilRet = BTRV_ERR_NONE Then
                                                                            slZone = "EST"  'Use EST as standard, if not found, use OTH
                                                                            ilCopy = mObtainCopy(slZone)
                                                                            If Not ilCopy Then
                                                                                slMsg = "Copy Missing: " & Trim$(tmSVef.sName)
                                                                                gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                                slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                                slMsg = slMsg & " for " & Trim$(str$(tmChf.lCntrNo)) & " " & Trim$(tmAdf.sName)
                                                                                'Print #hmMsg, slMsg
                                                                                'lbcMsg.AddItem slMsg
                                                                            End If
                                                                            If ilDlfFound Then
                                                                                tmDlfSrchKey.iVefCode = ilVehCode
                                                                                tmDlfSrchKey.sAirDay = slDay
                                                                                tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                                tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                                tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                                tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                                ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                                                    ilTerminated = False
                                                                                    If (tmDlf.sFed = "N") Or (StrComp(Trim$(tmDlf.sZone), "CST", 1) <> 0) Then
                                                                                        ilTerminated = True
                                                                                    Else
                                                                                        If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                                            If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                                                ilTerminated = True
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                    If Not ilTerminated Then
                                                                                        If ilSubFeed Then
                                                                                            If tmDlf.iMnfSubFeed > 0 Then
                                                                                                '6/1/16: Replaced GoSub
                                                                                                'GoSub lProcSpot
                                                                                                If Not mProcSpot(ilVehCode, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilLocalHour, slCodeStn, llSpotTime, llPrevTime, ilShowTime, llBChfCode, llBCifCode, llBSifCode, llBBofCode, ilBlackoutDefined, slDate, slTime, ilCopy, slCart, llCopyMissingSdfCode, slShortTitle, llHour, llPrevHour, ilPageNo, ilNoLines, ilLineNo, ilRet) Then
                                                                                                    mExpPhnxDay = False
                                                                                                    Exit Function
                                                                                                End If
                                                                                                DoEvents
                                                                                                If ilRet <> 0 Then
                                                                                                    mExpPhnxDay = False
                                                                                                    Exit Function
                                                                                                End If
                                                                                            End If
                                                                                        Else
                                                                                            If tmDlf.iMnfSubFeed = 0 Then
                                                                                                '6/1/16: Replaced GoSub
                                                                                                'GoSub lProcSpot
                                                                                                If Not mProcSpot(ilVehCode, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilLocalHour, slCodeStn, llSpotTime, llPrevTime, ilShowTime, llBChfCode, llBCifCode, llBSifCode, llBBofCode, ilBlackoutDefined, slDate, slTime, ilCopy, slCart, llCopyMissingSdfCode, slShortTitle, llHour, llPrevHour, ilPageNo, ilNoLines, ilLineNo, ilRet) Then
                                                                                                    mExpPhnxDay = False
                                                                                                    Exit Function
                                                                                                End If
                                                                                                DoEvents
                                                                                                If ilRet <> 0 Then
                                                                                                    mExpPhnxDay = False
                                                                                                    Exit Function
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                    ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                                Loop
                                                                            Else
                                                                                tmDlf.iAirTime(0) = tmAvail.iTime(0)
                                                                                tmDlf.iAirTime(1) = tmAvail.iTime(1)
                                                                                tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                                tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                                tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                                tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                                tmDlf.sZone = ""
                                                                                tmDlf.iEtfCode = 0
                                                                                tmDlf.iEnfCode = 0
                                                                                tmDlf.sProgCode = ""
                                                                                tmDlf.iMnfFeed = 0
                                                                                tmDlf.sBus = ""
                                                                                tmDlf.sSchedule = ""
                                                                                tmDlf.iMnfSubFeed = 0
                                                                                '6/1/16: Replaced GoSub
                                                                                'GoSub lProcSpot
                                                                                If Not mProcSpot(ilVehCode, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilLocalHour, slCodeStn, llSpotTime, llPrevTime, ilShowTime, llBChfCode, llBCifCode, llBSifCode, llBBofCode, ilBlackoutDefined, slDate, slTime, ilCopy, slCart, llCopyMissingSdfCode, slShortTitle, llHour, llPrevHour, ilPageNo, ilNoLines, ilLineNo, ilRet) Then
                                                                                    mExpPhnxDay = False
                                                                                    Exit Function
                                                                                End If
                                                                                If (ilRet <> 0) Or (imTerminate = True) Then
                                                                                    mExpPhnxDay = False
                                                                                    Exit Function
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            Screen.MousePointer = vbDefault
                                                                            ''MsgBox Trim$(tmSVef.sName) & " advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                            gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                            Screen.MousePointer = vbHourglass
                                                                            mExpPhnxDay = False
                                                                            Exit Function
                                                                        End If
                                                                    Else
                                                                        Screen.MousePointer = vbDefault
                                                                        ''MsgBox Trim$(tmSVef.sName) & " contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                        gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                        Screen.MousePointer = vbHourglass
                                                                        mExpPhnxDay = False
                                                                        Exit Function
                                                                    End If
                                                                Else
                                                                    Screen.MousePointer = vbDefault
                                                                    ''MsgBox Trim$(tmSVef.sName) & " spot get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                    gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " spot get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                    Screen.MousePointer = vbHourglass
                                                                    mExpPhnxDay = False
                                                                    Exit Function
                                                                End If
                                                            Next ilSpot
                                                            Exit Do
                                                        End If
                                                    End If
                                                Next ilSIndex
                                                ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                ilRet = gSSFGetNext(hmCTSsf, tmCTSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                            Loop
                                        End If
                                        ilRet = btrGetNext(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                Else
                                    ilAvailUnits = tmAvail.iAvInfo And &H1F
                                    ilAvailLen = tmAvail.iLen
                                    ilSpotLen = 0
                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpot + ilEvt)
                                        If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                            ilSpotLen = ilSpotLen + (tmSpot.iPosLen And &HFFF)
                                            ilSpotUnits = ilSpotUnits + 1
                                        End If
                                    Next ilSpot
                                    If ilAvailLen > ilSpotLen Then
                                        If ilDlfFound Then
                                            'Obtain engineering entry to see is avail is sent
                                            tmDlfSrchKey.iVefCode = ilVehCode
                                            tmDlfSrchKey.sAirDay = slDay
                                            tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                            tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                            tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                            tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                            ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                ilTerminated = False
                                                If (tmDlf.sFed = "N") Or (StrComp(Trim$(tmDlf.sZone), "CST", 1) <> 0) Then
                                                    ilTerminated = True
                                                Else
                                                    If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                        If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                            ilTerminated = True
                                                        End If
                                                    End If
                                                End If
                                                If Not ilTerminated Then
                                                    If ilSubFeed Then
                                                        If tmDlf.iMnfSubFeed > 0 Then
                                                            '6/1/16: Replaced GoSub
                                                            'GoSub lProcAdjDate  'Result stored into slAirDate
                                                            mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour
                                                            If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                slMsg = "Unsold: " & Trim$(tmVef.sName)
                                                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                                slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                tmAnfSrchKey.iCode = tmAvail.ianfCode
                                                                ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                If (ilRet = BTRV_ERR_NONE) Then
                                                                    slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                End If
                                                                'Print #hmMsg, slMsg
                                                                gAutomationAlertAndLogHandler slMsg
                                                                lbcMsg.AddItem slMsg
                                                            End If
                                                            DoEvents
                                                            If ilRet <> 0 Then
                                                                mExpPhnxDay = False
                                                                Exit Function
                                                            End If
                                                            Exit Do
                                                        End If
                                                    Else
                                                        If tmDlf.iMnfSubFeed = 0 Then
                                                            '6/1/16: Replaced GoSub
                                                            'GoSub lProcAdjDate  'Result stored into slAirDate
                                                            mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour
                                                            If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                slMsg = "Unsold: " & Trim$(tmVef.sName)
                                                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                                slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                tmAnfSrchKey.iCode = tmAvail.ianfCode
                                                                ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                If (ilRet = BTRV_ERR_NONE) Then
                                                                    slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                End If
                                                                'Print #hmMsg, slMsg
                                                                gAutomationAlertAndLogHandler slMsg
                                                                lbcMsg.AddItem slMsg
                                                            End If
                                                            DoEvents
                                                            If ilRet <> 0 Then
                                                                mExpPhnxDay = False
                                                                Exit Function
                                                            End If
                                                            Exit Do
                                                        End If
                                                    End If
                                                End If
                                                ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                            Loop
                                        Else
                                            tmDlf.iAirTime(0) = tmAvail.iTime(0)
                                            tmDlf.iAirTime(1) = tmAvail.iTime(1)
                                            tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                            tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                            tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                            tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                            tmDlf.sZone = ""
                                            tmDlf.iEtfCode = 0
                                            tmDlf.iEnfCode = 0
                                            tmDlf.sProgCode = ""
                                            tmDlf.iMnfFeed = 0
                                            tmDlf.sBus = ""
                                            tmDlf.sSchedule = ""
                                            tmDlf.iMnfSubFeed = 0
                                            '6/1/16: Replaced GoSub
                                            'GoSub lProcAdjDate  'Result stored into slAirDate
                                            mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour
                                            If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                slMsg = "Unsold: " & Trim$(tmVef.sName)
                                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                slMsg = slMsg & " on " & slDate & " at " & slTime
                                                tmAnfSrchKey.iCode = tmAvail.ianfCode
                                                ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                If (ilRet = BTRV_ERR_NONE) Then
                                                    slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                End If
                                                'Print #hmMsg, slMsg
                                                gAutomationAlertAndLogHandler slMsg
                                                lbcMsg.AddItem slMsg
                                            End If
                                            DoEvents
                                            If ilRet <> 0 Then
                                                mExpPhnxDay = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                        ilEvt = ilEvt + 1
                                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet = BTRV_ERR_NONE Then
                                            tmChfSrchKey.lCode = tmSdf.lChfCode
                                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet = BTRV_ERR_NONE Then
                                                tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilRet = BTRV_ERR_NONE Then
                                                    slZone = "EST"  'Use EST as standard, if not found, use OTH
                                                    ilCopy = mObtainCopy(slZone)
                                                    If Not ilCopy Then
                                                        slMsg = "Copy Missing: " & Trim$(smVehName)
                                                        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                        slMsg = slMsg & " on " & slDate & " at " & slTime
                                                        slMsg = slMsg & " for " & Trim$(str$(tmChf.lCntrNo)) & " " & Trim$(tmAdf.sName)
                                                        'Print #hmMsg, slMsg
                                                        'lbcMsg.AddItem slMsg
                                                    End If
                                                    If ilDlfFound Then
                                                        'Obtain engineering entry to see is avail is sent
                                                        tmDlfSrchKey.iVefCode = ilVehCode
                                                        tmDlfSrchKey.sAirDay = slDay
                                                        tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                        tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                        tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                        tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                        ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                        Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                            ilTerminated = False
                                                            If (tmDlf.sFed = "N") Or (StrComp(Trim$(tmDlf.sZone), "CST", 1) <> 0) Then
                                                                ilTerminated = True
                                                            Else
                                                                If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                    If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                        ilTerminated = True
                                                                    End If
                                                                End If
                                                            End If
                                                            If Not ilTerminated Then
                                                                If ilSubFeed Then
                                                                    If tmDlf.iMnfSubFeed > 0 Then
                                                                        '6/1/16: Replaced GoSub
                                                                        'GoSub lProcSpot
                                                                        If Not mProcSpot(ilVehCode, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilLocalHour, slCodeStn, llSpotTime, llPrevTime, ilShowTime, llBChfCode, llBCifCode, llBSifCode, llBBofCode, ilBlackoutDefined, slDate, slTime, ilCopy, slCart, llCopyMissingSdfCode, slShortTitle, llHour, llPrevHour, ilPageNo, ilNoLines, ilLineNo, ilRet) Then
                                                                            mExpPhnxDay = False
                                                                            Exit Function
                                                                        End If
                                                                        DoEvents
                                                                        If ilRet <> 0 Then
                                                                            mExpPhnxDay = False
                                                                            Exit Function
                                                                        End If
                                                                    End If
                                                                Else
                                                                    If tmDlf.iMnfSubFeed = 0 Then
                                                                        '6/1/16: Replaced GoSub
                                                                        'GoSub lProcSpot
                                                                        If Not mProcSpot(ilVehCode, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilLocalHour, slCodeStn, llSpotTime, llPrevTime, ilShowTime, llBChfCode, llBCifCode, llBSifCode, llBBofCode, ilBlackoutDefined, slDate, slTime, ilCopy, slCart, llCopyMissingSdfCode, slShortTitle, llHour, llPrevHour, ilPageNo, ilNoLines, ilLineNo, ilRet) Then
                                                                            mExpPhnxDay = False
                                                                            Exit Function
                                                                        End If
                                                                        DoEvents
                                                                        If ilRet <> 0 Then
                                                                            mExpPhnxDay = False
                                                                            Exit Function
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                            ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        Loop
                                                    Else
                                                        tmDlf.iAirTime(0) = tmAvail.iTime(0)
                                                        tmDlf.iAirTime(1) = tmAvail.iTime(1)
                                                        tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                        tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                        tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                        tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                        tmDlf.sZone = ""
                                                        tmDlf.iEtfCode = 0
                                                        tmDlf.iEnfCode = 0
                                                        tmDlf.sProgCode = ""
                                                        tmDlf.iMnfFeed = 0
                                                        tmDlf.sBus = ""
                                                        tmDlf.sSchedule = ""
                                                        tmDlf.iMnfSubFeed = 0
                                                        '6/1/16: Replaced GoSub
                                                        'GoSub lProcSpot
                                                        If Not mProcSpot(ilVehCode, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilLocalHour, slCodeStn, llSpotTime, llPrevTime, ilShowTime, llBChfCode, llBCifCode, llBSifCode, llBBofCode, ilBlackoutDefined, slDate, slTime, ilCopy, slCart, llCopyMissingSdfCode, slShortTitle, llHour, llPrevHour, ilPageNo, ilNoLines, ilLineNo, ilRet) Then
                                                            mExpPhnxDay = False
                                                            Exit Function
                                                        End If
                                                        If (ilRet <> 0) Or (imTerminate = True) Then
                                                            mExpPhnxDay = False
                                                            Exit Function
                                                        End If
                                                    End If
                                                Else
                                                    Screen.MousePointer = vbDefault
                                                    ''MsgBox smVehName & " advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                    gAutomationAlertAndLogHandler smVehName & " advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                    Screen.MousePointer = vbHourglass
                                                    mExpPhnxDay = False
                                                    Exit Function
                                                End If
                                            Else
                                                Screen.MousePointer = vbDefault
                                                ''MsgBox smVehName & " contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                gAutomationAlertAndLogHandler smVehName & " contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                Screen.MousePointer = vbHourglass
                                                mExpPhnxDay = False
                                                Exit Function
                                            End If
                                        Else
                                            Screen.MousePointer = vbDefault
                                            ''MsgBox smVehName & " spot get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                            gAutomationAlertAndLogHandler smVehName & " spot get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                            Screen.MousePointer = vbHourglass
                                            mExpPhnxDay = False
                                            Exit Function
                                        End If
                                    Next ilSpot
                                End If
                            End If
                        End If
                    End If
                    ilEvt = ilEvt + 1
                Loop
                ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    Next llDate
    mExpPhnxDay = True
    Exit Function

'lProcSpot:
'    'Record type
'    'Test if spot is to be sent
'    If tmDlf.iMnfSubFeed > 0 Then
'        tmMnfSrchKey.iCode = tmDlf.iMnfSubFeed
'        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'        slCodeStn = Left$(tmMnf.sCodeStn, 5)
'        'Export only Real Country for Phoenix (all others are ignored)
'        If StrComp(Trim$(slCodeStn), "RC", 1) <> 0 Then
'            Return
'        End If
'    Else
'        slCodeStn = Left$(tmVef.sCodeStn, 5)
'    End If
'    '6/1/16: Replaced GoSub
'    'GoSub lProcAdjDate  'Result stored into slAirDate
'    mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour
'    If gDateValue(slAirDate) > llEDate Then
'        Return
'    End If
'    If gDateValue(slAirDate) < llSDate Then
'        Return
'    End If
'    '6/1/16: Replaced GoSub
'    'GoSub lProcSpotTime 'Result stored into slSpotTime
'    mProcSpotTime slAirDate, llAirDate, llAvailTime, ilVehCode, slCodeStn, llSpotTime, llPrevTime, ilShowTime
'    'Test if Blackout exist
'    If gSDFBlackoutTest(hmRsf, hmBof, tmSdf.lCode, ilVehCode, llBChfCode, llBCifCode, llBSifCode, llBBofCode) Then
'        ilBlackoutDefined = True
'        tmChfSrchKey.lCode = llBChfCode
'        ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        If ilRet = BTRV_ERR_NONE Then
'            tmAdfSrchKey.iCode = tmChf.iAdfCode
'            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'            If ilRet <> BTRV_ERR_NONE Then
'                Screen.MousePointer = vbDefault
'                MsgBox Trim$(tmSVef.sName) & " blackout advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
'                Screen.MousePointer = vbHourglass
'                mExpPhnxDay = False
'                Exit Function
'            End If
'        Else
'            Screen.MousePointer = vbDefault
'            MsgBox Trim$(tmSVef.sName) & " blackout contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
'            Screen.MousePointer = vbHourglass
'            mExpPhnxDay = False
'            Exit Function
'        End If
'        If llBCifCode <= 0 Then
'            slMsg = "Blackout Copy Missing : " & Trim$(tmVef.sName)
'            gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
'            slMsg = slMsg & " on " & slDate & " at " & slTime
'            slMsg = slMsg & " for " & Trim$(str$(tmChf.lCntrNo)) & " " & Trim$(tmAdf.sName)
'            Print #hmMsg, slMsg
'            lbcMsg.AddItem slMsg
'        End If
'    Else
'        ilBlackoutDefined = False
'        'Get Copy for spot
'        'slZone = "EST"  'Trim$(tmDlf.sZone) use EST as standard, if not found, use OTH
'        'If slZone = "" Then
'        '    slZone = "Oth"
'        'End If
'        'ilCopy = mObtainCopy(slZone)
'        'Later if copy missing- output error and stop creation, delete file
'        If Not ilCopy Then
'            If llCopyMissingSdfCode <> tmSdf.lCode Then
'                Print #hmMsg, slMsg
'                lbcMsg.AddItem slMsg
'                llCopyMissingSdfCode = tmSdf.lCode
'            End If
'        End If
'    End If
'
'    'Generate record: Time Cart Length Advertiser,Product
'    llHour = llSpotTime \ 3600
'    If ilShowTime Then
'        slTime = gCurrencyToTime(CCur(llSpotTime))
'    Else
'        slTime = " "
'    End If
'    Do While Len(slTime) < 10
'        slTime = slTime & " "
'    Loop
'    'Cart ID
'    If ilBlackoutDefined Then
'        tmCifSrchKey.lCode = llBCifCode
'        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            ilCopy = True
'            tmMcfSrchKey.iCode = tmCif.iMcfCode
'            ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'            If ilRet <> BTRV_ERR_NONE Then
'                tmMcf.sName = "C"
'                tmMcf.sPrefix = "C"
'            End If
'        Else
'            ilCopy = False
'        End If
'    End If
'    If ilCopy Then
'        If Trim$(tmCif.sCut) = "" Then
'            slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & " "
'        Else
'            slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut) & " "
'        End If
'    Else
'        slCart = "C    "
'        imCopyMissing = True
'    End If
'    slLen = Trim$(str$(tmSdf.iLen))
'    Do While Len(slLen) < 3
'        slLen = " " & slLen
'    Loop
'    'Advertiser name
'    slAdvtName = Trim$(UCase$(Left$(tmAdf.sName, 27)))
'    'If ilCopy And Trim$(tmCpf.sName) <> "" Then
'    '    slShortTitle = Trim$(UCase$(Left$(tmCpf.sName, 15)))
'    'Else
'    '    slShortTitle = Trim$(UCase$(Left$(tmChf.sProduct, 15)))
'    'End If
'    If ilBlackoutDefined Then
'        If tgSpf.sUseProdSptScr = "P" Then
'            If llBSifCode <= 0 Then
'                tmSifSrchKey.lCode = tmChf.lSifCode
'            Else
'                tmSifSrchKey.lCode = llBSifCode
'            End If
'            ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'            If ilRet = BTRV_ERR_NONE Then
'                slShortTitle = Trim$(tmSif.sName)
'            Else
'                slShortTitle = ""
'            End If
'        Else
'            slShortTitle = Trim$(tmAdf.sAbbr) & "," & Trim$(tmChf.sProduct)
'        End If
'    Else
'        slShortTitle = UCase$(Left$(gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf), 15))
'    End If
'    slAdvtProd = slAdvtName & ", " & slShortTitle
'    slLine = slTime & "| " & slCart & " " & slLen & " " & slAdvtProd & "|"
'    ilRet = 0
'    If (llHour <> llPrevHour) And (ilPageNo > 0) Then
'        ilNoLines = 5   'Each cluster can have 4 spots- don't split clusters'1
'    Else
'        ilNoLines = 0
'    End If
'    If ilLineNo + ilNoLines >= 66 Then 'Output header
'        If ilPageNo > 0 Then
'            smNewLines(UBound(smNewLines)) = Chr(12)    'Form Feed
'            ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
'        End If
'        ilPageNo = ilPageNo + 1
'        slHeader = "Commercial Studio Log"
'        Do While Len(slHeader) < 70
'            slHeader = slHeader & " "
'        Loop
'        slHeader = slHeader & "Page"
'        slStr = Trim$(str$(ilPageNo))
'        Do While Len(slStr) < 6
'            slStr = " " & slStr
'        Loop
'        slHeader = slHeader & slStr
'        smNewLines(UBound(smNewLines)) = slHeader    'Form Feed
'        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
'        If tmDlf.iMnfSubFeed > 0 Then
'            slHeader = Trim$(tmMnf.sName)
'        Else
'            slHeader = Trim$(tmVef.sName)
'        End If
'        Select Case gWeekDayStr(slAirDate)
'            Case 0
'                slDayName = "Monday, "
'            Case 1
'                slDayName = "Tuesday, "
'            Case 2
'                slDayName = "Wednesday, "
'            Case 3
'                slDayName = "Thursday, "
'            Case 4
'                slDayName = "Friday, "
'            Case 5
'                slDayName = "Saturday, "
'            Case 6
'                slDayName = "Sunday, "
'        End Select
'        smNewLines(UBound(smNewLines)) = slHeader & " on " & slDayName & slAirDate
'        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
'        smNewLines(UBound(smNewLines)) = "--------------------------------------------------------------------------------"    'Blank line
'        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
'        slHeader = "Time      | Cart   Len Advertiser, Product                          | Notes"
'        smNewLines(UBound(smNewLines)) = slHeader    '
'        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
'        smNewLines(UBound(smNewLines)) = "--------------------------------------------------------------------------------"    'Blank line
'        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
'        ilLineNo = 5
'        llPrevHour = llHour 'Don't start new page with blank line
'    End If
'    If llPrevHour <> llHour Then
'        smNewLines(UBound(smNewLines)) = "--------------------------------------------------------------------------------"    'Blank line
'        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
'        ilLineNo = ilLineNo + 1
'        llPrevHour = llHour
'    End If
'    smNewLines(UBound(smNewLines)) = slLine
'    ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
'    ilLineNo = ilLineNo + 1
'    Return
'lProcAdjDate:
'    'Test if Air time is AM and Feed Time is PM. If so, adjust date
'    slAirDate = Format$(llDate, "m/d/yy")
'    ilAirHour = tmDlf.iAirTime(1) \ 256  'Obtain Hour
'    ilLocalHour = tmDlf.iLocalTime(1) \ 256  'Obtain Hour
'    If (ilAirHour < 6) And (ilLocalHour > 17) Then
'        slAirDate = gDecOneDay(slAirDate)
'    End If
'    Return
'lProcSpotTime:
'    llAirDate = gDateValue(slAirDate)
'    gUnpackTimeLong tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), False, llAvailTime
'    For ilTest = LBound(tmSpotTimes) To UBound(tmSpotTimes) - 1 Step 1
'        If (tmSpotTimes(ilTest).iVefCode = ilVehCode) And (tmSpotTimes(ilTest).sCodeStn = slCodeStn) And (tmSpotTimes(ilTest).lAirDate = llAirDate) And (tmSpotTimes(ilTest).lAvailTime = llAvailTime) Then
'            ilShowTime = False
'            llSpotTime = tmSpotTimes(ilTest).lNextSpotTime
'            tmSpotTimes(ilTest).lNextSpotTime = tmSpotTimes(ilTest).lNextSpotTime + tmSdf.iLen
'            llPrevTime = tmSpotTimes(ilTest).lNextSpotTime
'            Return
'        End If
'    Next ilTest
'    ReDim Preserve tmSpotTimes(0 To UBound(tmSpotTimes) + 1) As SPOTTIMES
'    tmSpotTimes(UBound(tmSpotTimes) - 1).iVefCode = ilVehCode
'    tmSpotTimes(UBound(tmSpotTimes) - 1).sCodeStn = slCodeStn
'    tmSpotTimes(UBound(tmSpotTimes) - 1).lAirDate = llAirDate
'    tmSpotTimes(UBound(tmSpotTimes) - 1).lAvailTime = llAvailTime
'    tmSpotTimes(UBound(tmSpotTimes) - 1).lNextSpotTime = llAvailTime + tmSdf.iLen
'    llSpotTime = llAvailTime
'    If llPrevTime = llAvailTime Then
'        ilShowTime = False
'    Else
'        ilShowTime = True
'    End If
'    llPrevTime = tmSpotTimes(UBound(tmSpotTimes) - 1).lNextSpotTime
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilTest As Integer
    Dim ilVefCode As Integer
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imSifRecLen = Len(tmSif)
    hmRsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imRsfRecLen = Len(tmRsf)
    hmBof = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmBof, "", sgDBPath & "Bof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imBofRecLen = Len(tmBof)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imTzfRecLen = Len(tmTzf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)
    hmDlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmDlf, "", sgDBPath & "Dlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imDlfRecLen = Len(tmDlf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    hmCTSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCTSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpPhnx
    On Error GoTo 0
    'Populate arrays to determine if records exist
    mVehPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        'mTerminate
        Exit Sub
    End If
    'Select Satellite Music; SMN Mix; Tom Joyner
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        'For ilTest = 0 To UBound(tgVpf) Step 1
        '    If ilVefCode = tgVpf(ilTest).iVefKCode Then
            ilTest = gBinarySearchVpf(ilVefCode)
            If ilTest <> -1 Then
                If tgVpf(ilTest).sExpHiPhoenix = "Y" Then
                    lbcVehicle.Selected(ilLoop) = True
                End If
        '        Exit For
            End If
        'Next ilTest
        'If (StrComp(lbcVehicle.List(ilLoop), "Satellite Music", 1) = 0) Or (StrComp(lbcVehicle.List(ilLoop), "24-Hour Formats", 1) = 0) Then
        '    lbcVehicle.Selected(ilLoop) = True
        'ElseIf (StrComp(lbcVehicle.List(ilLoop), "SMN Mix", 1) = 0) Or (StrComp(lbcVehicle.List(ilLoop), "Mix 24-Hour Formats", 1) = 0) Then
        '    'lbcVehicle.Selected(ilLoop) = True
        'ElseIf StrComp(lbcVehicle.List(ilLoop), "Tom Joyner", 1) = 0 Then
        '    'lbcVehicle.Selected(ilLoop) = True
        'End If
    Next ilLoop
    For ilLoop = LBound(imEvtType) To UBound(imEvtType) Step 1
        imEvtType(ilLoop) = True
    Next ilLoop
    imEvtType(0) = False 'Don't include library names
    'plcGauge.Move ExpPhnx.Width / 2 - plcGauge.Width / 2
    'cmcFileConv.Move ExpPhnx.Width / 2 - cmcFileConv.Width / 2
    'cmcCancel.Move ExpPhnx.Width / 2 - cmcCancel.Width / 2 - cmcCancel.Width
    'cmcReport.Move ExpPhnx.Width / 2 - cmcReport.Width / 2 + cmcReport.Width
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    slStr = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slStr)
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    lacDate.Visible = False
    gCenterStdAlone ExpPhnx
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    Exit Sub
mInitErr:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move edcStartDate.Left, edcStartDate.Top + edcStartDate.Height
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCopy                     *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Private Function mObtainCopy(slZone As String) As Integer
'
'   mObtainCopy
'       Where:
'           tmSdf(I)- Spot record
'           tmCif(O)- Inventory
'           tmCpf(O)- Product/ISCI record
'
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilCifFound As Integer
    ilCifFound = False
    tmCpf.sISCI = ""
    tmCpf.sName = ""
    tmCpf.sCreative = ""
    tmMcf.sName = "C"
    tmMcf.sPrefix = "C"
    If tmSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilCifFound = True
        End If
    ElseIf tmSdf.sPtType = "2" Then  '  Combo Copy
    ElseIf tmSdf.sPtType = "3" Then  '  Time Zone Copy
        ' Read TZF using lCopyCode from SDF
        tmTzfSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ' Look for the first positive lZone value
        For ilIndex = 1 To 6 Step 1
            If tmTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                If StrComp(tmTzf.sZone(ilIndex - 1), slZone, 1) = 0 Then
                    ' Read CIF using lCopyCode from SDF
                    tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        ilCifFound = True
                    End If
                    Exit For
                End If
            End If
        Next ilIndex
        If Not ilCifFound Then
            For ilIndex = 1 To 6 Step 1
                If tmTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                    If StrComp(tmTzf.sZone(ilIndex - 1), "Oth", 1) = 0 Then
                        ' Read CIF using lCopyCode from SDF
                        tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            ilCifFound = True
                        End If
                        Exit For
                    End If
                End If
            Next ilIndex
        End If
    End If
    If ilCifFound Then
        ' Read CPF using lCpfCode from CIF
        If tmCif.lcpfCode > 0 Then
            tmCpfSrchKey.lCode = tmCif.lcpfCode
            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmCpf.sISCI = ""
                tmCpf.sName = ""
                tmCpf.sCreative = ""
            End If
        Else
            tmCpf.sISCI = ""
            tmCpf.sName = ""
            tmCpf.sCreative = ""
        End If
        tmMcfSrchKey.iCode = tmCif.iMcfCode
        ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmMcf.sName = "C"
            tmMcf.sPrefix = "C"
            mObtainCopy = False
            Exit Function
        End If
        mObtainCopy = True
        Exit Function
    End If
    mObtainCopy = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    'On Error GoTo mOpenMsgFileErr:
    ''slToFile = sgExportPath & "ExpPhnx.Txt"
    slToFile = sgDBPath & "Messages\" & "ExpPhnx.Txt"
    sgMessageFile = slToFile
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            'ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            'ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        'ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    'Print #hmMsg, ""
    
    'Print #hmMsg, "Export Phoenix: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    gAutomationAlertAndLogHandler "Export Phoenix: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
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
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer

    slCommand = sgCommandStr    'Command$
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            ''MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            gAutomationAlertAndLogHandler "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
'        imShowHelpMsg = True
'        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
'        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
'            imShowHelpMsg = False
'        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone ExpPhnx, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igCmmCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igCmmCallSource = CALLNONE
    'End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'

    Screen.MousePointer = vbDefault

    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpPhnx
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVpf As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    
    'ilRet = gPopUserVehicleBox(ExpPhnx, VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH, lbcVehicle, Traffic!lbcUserVehicle)
    ilRet = gPopUserVehicleBox(ExpPhnx, VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExpPhnx
        On Error GoTo 0
    End If
    'Select on all vehicles that have Clearance as the format
    'For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
    '    slNameCode = Traffic!lbcUserVehicle.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    tmVefSrchKey.iCode = Val(slCode)
    '   ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    '    If (ilRet = BTRV_ERR_NONE) Then
    '        If StrComp(Trim$(tmVef.sFormat), "Clearance", 1) = 0 Then
    '            lbcVehicle.Selected(ilLoop) = True
    '        End If
    '    End If
    'Next ilLoop
    
    For ilLoop = LBound(tgUserVehicle) To UBound(tgUserVehicle) - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        ilVpf = gBinarySearchVpf(ilVefCode)
        If ilVpf <> -1 Then
            If tgVpf(ilVpf).sExpHiPhoenix = "Y" Then
                lbcVehicle.Selected(ilLoop) = True
            End If
        End If
    Next ilLoop

    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcStartDate.Text = Format$(llDate, "m/d/yy")
                edcStartDate.SelStart = 0
                edcStartDate.SelLength = Len(edcStartDate.Text)
                imBypassFocus = True
                edcStartDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcStartDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Export Phoenix Log"
End Sub

Private Sub mProcAdjDate(slAirDate As String, llDate As Long, ilAirHour As Integer, ilLocalHour As Integer)
    'Test if Air time is AM and Feed Time is PM. If so, adjust date
    slAirDate = Format$(llDate, "m/d/yy")
    ilAirHour = tmDlf.iAirTime(1) \ 256  'Obtain Hour
    ilLocalHour = tmDlf.iLocalTime(1) \ 256  'Obtain Hour
    If (ilAirHour < 6) And (ilLocalHour > 17) Then
        slAirDate = gDecOneDay(slAirDate)
    End If
End Sub

Private Sub mProcSpotTime(slAirDate As String, llAirDate As Long, llAvailTime As Long, ilVehCode As Integer, slCodeStn As String, llSpotTime As Long, llPrevTime As Long, ilShowTime As Integer)
    Dim ilTest As Integer
    
    llAirDate = gDateValue(slAirDate)
    gUnpackTimeLong tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), False, llAvailTime
    For ilTest = LBound(tmSpotTimes) To UBound(tmSpotTimes) - 1 Step 1
        If (tmSpotTimes(ilTest).iVefCode = ilVehCode) And (tmSpotTimes(ilTest).sCodeStn = slCodeStn) And (tmSpotTimes(ilTest).lAirDate = llAirDate) And (tmSpotTimes(ilTest).lAvailTime = llAvailTime) Then
            ilShowTime = False
            llSpotTime = tmSpotTimes(ilTest).lNextSpotTime
            tmSpotTimes(ilTest).lNextSpotTime = tmSpotTimes(ilTest).lNextSpotTime + tmSdf.iLen
            llPrevTime = tmSpotTimes(ilTest).lNextSpotTime
            'Return
            Exit Sub
        End If
    Next ilTest
    ReDim Preserve tmSpotTimes(0 To UBound(tmSpotTimes) + 1) As SPOTTIMES
    tmSpotTimes(UBound(tmSpotTimes) - 1).iVefCode = ilVehCode
    tmSpotTimes(UBound(tmSpotTimes) - 1).sCodeStn = slCodeStn
    tmSpotTimes(UBound(tmSpotTimes) - 1).lAirDate = llAirDate
    tmSpotTimes(UBound(tmSpotTimes) - 1).lAvailTime = llAvailTime
    tmSpotTimes(UBound(tmSpotTimes) - 1).lNextSpotTime = llAvailTime + tmSdf.iLen
    llSpotTime = llAvailTime
    If llPrevTime = llAvailTime Then
        ilShowTime = False
    Else
        ilShowTime = True
    End If
    llPrevTime = tmSpotTimes(UBound(tmSpotTimes) - 1).lNextSpotTime
End Sub

Private Function mProcSpot(ilVehCode As Integer, llSDate As Long, llEDate As Long, slAirDate As String, llAirDate As Long, llAvailTime As Long, llDate As Long, ilAirHour As Integer, ilLocalHour As Integer, slCodeStn As String, llSpotTime As Long, llPrevTime As Long, ilShowTime As Integer, llBChfCode As Long, llBCifCode As Long, llBSifCode As Long, llBBofCode As Long, ilBlackoutDefined As Integer, slDate As String, slTime As String, ilCopy As Integer, slCart As String, llCopyMissingSdfCode As Long, slShortTitle As String, llHour As Long, llPrevHour As Long, ilPageNo As Integer, ilNoLines As Integer, ilLineNo As Integer, ilRet As Integer) As Integer
    Dim slMsg As String
    Dim slStr As String
    Dim slAdvtName As String
    Dim slAdvtProd As String * 45   '52
    Dim slLen As String
    Dim slLine As String
    Dim slHeader As String
    Dim slDayName As String
    
    mProcSpot = True
    ilRet = 0
    'Record type
    'Test if spot is to be sent
    If tmDlf.iMnfSubFeed > 0 Then
        tmMnfSrchKey.iCode = tmDlf.iMnfSubFeed
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        slCodeStn = Left$(tmMnf.sCodeStn, 5)
        'Export only Real Country for Phoenix (all others are ignored)
        If StrComp(Trim$(slCodeStn), "RC", 1) <> 0 Then
            'Return
            mProcSpot = True
            Exit Function
        End If
    Else
        slCodeStn = Left$(tmVef.sCodeStn, 5)
    End If
    '6/1/16: Replaced GoSub
    'GoSub lProcAdjDate  'Result stored into slAirDate
    mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour
    If gDateValue(slAirDate) > llEDate Then
        'Return
        mProcSpot = True
        Exit Function
    End If
    If gDateValue(slAirDate) < llSDate Then
        'Return
        mProcSpot = True
        Exit Function
    End If
    '6/1/16: Replaced GoSub
    'GoSub lProcSpotTime 'Result stored into slSpotTime
    mProcSpotTime slAirDate, llAirDate, llAvailTime, ilVehCode, slCodeStn, llSpotTime, llPrevTime, ilShowTime
    'Test if Blackout exist
    If gSDFBlackoutTest(hmRsf, hmBof, tmSdf.lCode, ilVehCode, llBChfCode, llBCifCode, llBSifCode, llBBofCode) Then
        ilBlackoutDefined = True
        tmChfSrchKey.lCode = llBChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tmAdfSrchKey.iCode = tmChf.iAdfCode
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                ''MsgBox Trim$(tmSVef.sName) & " blackout advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " blackout advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                Screen.MousePointer = vbHourglass
                mProcSpot = False
                Exit Function
            End If
        Else
            Screen.MousePointer = vbDefault
            ''MsgBox Trim$(tmSVef.sName) & " blackout contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " blackout contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            Screen.MousePointer = vbHourglass
            mProcSpot = False
            Exit Function
        End If
        If llBCifCode <= 0 Then
            slMsg = "Blackout Copy Missing : " & Trim$(tmVef.sName)
            gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
            slMsg = slMsg & " on " & slDate & " at " & slTime
            slMsg = slMsg & " for " & Trim$(str$(tmChf.lCntrNo)) & " " & Trim$(tmAdf.sName)
            'Print #hmMsg, slMsg
            gAutomationAlertAndLogHandler slMsg
            lbcMsg.AddItem slMsg
        End If
    Else
        ilBlackoutDefined = False
        'Get Copy for spot
        'slZone = "EST"  'Trim$(tmDlf.sZone) use EST as standard, if not found, use OTH
        'If slZone = "" Then
        '    slZone = "Oth"
        'End If
        'ilCopy = mObtainCopy(slZone)
        'Later if copy missing- output error and stop creation, delete file
        If Not ilCopy Then
            If llCopyMissingSdfCode <> tmSdf.lCode Then
                'Print #hmMsg, slMsg
                gAutomationAlertAndLogHandler slMsg
                lbcMsg.AddItem slMsg
                llCopyMissingSdfCode = tmSdf.lCode
            End If
        End If
    End If

    'Generate record: Time Cart Length Advertiser,Product
    llHour = llSpotTime \ 3600
    If ilShowTime Then
        slTime = gCurrencyToTime(CCur(llSpotTime))
    Else
        slTime = " "
    End If
    Do While Len(slTime) < 10
        slTime = slTime & " "
    Loop
    'Cart ID
    If ilBlackoutDefined Then
        tmCifSrchKey.lCode = llBCifCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilCopy = True
            tmMcfSrchKey.iCode = tmCif.iMcfCode
            ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmMcf.sName = "C"
                tmMcf.sPrefix = "C"
            End If
        Else
            ilCopy = False
        End If
    End If
    If ilCopy Then
        If Trim$(tmCif.sCut) = "" Then
            slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & " "
        Else
            slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut) & " "
        End If
    Else
        slCart = "C    "
        imCopyMissing = True
    End If
    slLen = Trim$(str$(tmSdf.iLen))
    Do While Len(slLen) < 3
        slLen = " " & slLen
    Loop
    'Advertiser name
    slAdvtName = Trim$(UCase$(Left$(tmAdf.sName, 27)))
    'If ilCopy And Trim$(tmCpf.sName) <> "" Then
    '    slShortTitle = Trim$(UCase$(Left$(tmCpf.sName, 15)))
    'Else
    '    slShortTitle = Trim$(UCase$(Left$(tmChf.sProduct, 15)))
    'End If
    If ilBlackoutDefined Then
        If tgSpf.sUseProdSptScr = "P" Then
            If llBSifCode <= 0 Then
                tmSifSrchKey.lCode = tmChf.lSifCode
            Else
                tmSifSrchKey.lCode = llBSifCode
            End If
            ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slShortTitle = Trim$(tmSif.sName)
            Else
                slShortTitle = ""
            End If
        Else
            slShortTitle = Trim$(tmAdf.sAbbr) & "," & Trim$(tmChf.sProduct)
        End If
    Else
        slShortTitle = UCase$(Left$(gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf), 15))
    End If
    slAdvtProd = slAdvtName & ", " & slShortTitle
    slLine = slTime & "| " & slCart & " " & slLen & " " & slAdvtProd & "|"
    ilRet = 0
    If (llHour <> llPrevHour) And (ilPageNo > 0) Then
        ilNoLines = 5   'Each cluster can have 4 spots- don't split clusters'1
    Else
        ilNoLines = 0
    End If
    If ilLineNo + ilNoLines >= 66 Then 'Output header
        If ilPageNo > 0 Then
            smNewLines(UBound(smNewLines)) = Chr(12)    'Form Feed
            ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
        End If
        ilPageNo = ilPageNo + 1
        slHeader = "Commercial Studio Log"
        Do While Len(slHeader) < 70
            slHeader = slHeader & " "
        Loop
        slHeader = slHeader & "Page"
        slStr = Trim$(str$(ilPageNo))
        Do While Len(slStr) < 6
            slStr = " " & slStr
        Loop
        slHeader = slHeader & slStr
        smNewLines(UBound(smNewLines)) = slHeader    'Form Feed
        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
        If tmDlf.iMnfSubFeed > 0 Then
            slHeader = Trim$(tmMnf.sName)
        Else
            slHeader = Trim$(tmVef.sName)
        End If
        Select Case gWeekDayStr(slAirDate)
            Case 0
                slDayName = "Monday, "
            Case 1
                slDayName = "Tuesday, "
            Case 2
                slDayName = "Wednesday, "
            Case 3
                slDayName = "Thursday, "
            Case 4
                slDayName = "Friday, "
            Case 5
                slDayName = "Saturday, "
            Case 6
                slDayName = "Sunday, "
        End Select
        smNewLines(UBound(smNewLines)) = slHeader & " on " & slDayName & slAirDate
        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
        smNewLines(UBound(smNewLines)) = "--------------------------------------------------------------------------------"    'Blank line
        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
        slHeader = "Time      | Cart   Len Advertiser, Product                          | Notes"
        smNewLines(UBound(smNewLines)) = slHeader    '
        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
        smNewLines(UBound(smNewLines)) = "--------------------------------------------------------------------------------"    'Blank line
        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
        ilLineNo = 5
        llPrevHour = llHour 'Don't start new page with blank line
    End If
    If llPrevHour <> llHour Then
        smNewLines(UBound(smNewLines)) = "--------------------------------------------------------------------------------"    'Blank line
        ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
        ilLineNo = ilLineNo + 1
        llPrevHour = llHour
    End If
    smNewLines(UBound(smNewLines)) = slLine
    ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String
    ilLineNo = ilLineNo + 1
    mProcSpot = True
    Exit Function
End Function
