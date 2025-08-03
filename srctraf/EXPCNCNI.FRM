VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExpCnCNI 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4815
   ClientLeft      =   795
   ClientTop       =   1335
   ClientWidth     =   7050
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
   ScaleHeight     =   4815
   ScaleWidth      =   7050
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   5880
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox edcNoWks 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3795
      MaxLength       =   10
      TabIndex        =   7
      Top             =   840
      Width           =   555
   End
   Begin VB.PictureBox plcTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   990
      ScaleHeight     =   345
      ScaleWidth      =   4275
      TabIndex        =   23
      Top             =   345
      Width           =   4305
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   60
         TabIndex        =   24
         Top             =   75
         Width           =   4200
      End
   End
   Begin VB.CommandButton cmcTo 
      Appearance      =   0  'Flat
      Caption         =   "&Browse..."
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
      Left            =   5370
      TabIndex        =   22
      Top             =   375
      Width           =   1485
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
      Left            =   2280
      Picture         =   "EXPCNCNI.frx":0000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   855
      Width           =   195
   End
   Begin VB.TextBox edcStartDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   2
      Top             =   840
      Width           =   1110
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1185
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   17
      Top             =   1050
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
         Top             =   30
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   30
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
         Picture         =   "EXPCNCNI.frx":00FA
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   19
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
         Top             =   30
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
      ScaleWidth      =   2250
      TabIndex        =   0
      Top             =   45
      Width           =   2250
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4185
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4455
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4605
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   1500
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   1200
      Width           =   3915
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
      Left            =   2265
      TabIndex        =   9
      Top             =   4395
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
      Left            =   3705
      TabIndex        =   10
      Top             =   4395
      Width           =   1050
   End
   Begin VB.Label lacBlk 
      Appearance      =   0  'Flat
      Caption         =   "(Leave Blank for all weeks)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   4605
      TabIndex        =   26
      Top             =   855
      Width           =   2085
   End
   Begin VB.Label lacNoWks 
      Appearance      =   0  'Flat
      Caption         =   "# of Weeks"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2655
      TabIndex        =   6
      Top             =   825
      Width           =   1065
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      Caption         =   "To File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   105
      TabIndex        =   25
      Top             =   420
      Width           =   810
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   180
      Top             =   4350
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   435
      TabIndex        =   21
      Top             =   4080
      Width           =   6330
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Export Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label lacErrors 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5055
      TabIndex        =   16
      Top             =   1815
      Width           =   1725
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   5175
      TabIndex        =   14
      Top             =   1395
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3255
      TabIndex        =   12
      Top             =   1845
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "ExpCnCNI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EXPCNCNI.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpCncNI.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export feed (for RCS Linker) input screen code
Option Explicit
Option Compare Text
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmTo As Integer   'From file hanle
Dim hmMsg As Integer   'From file hanle
Dim lmNowDate As Long   'Todays date
'Log Calendar file
Dim hmLcf As Integer
Dim tmLcf As LCF
Dim imLcfRecLen As Integer  'LCF record length
'Event Types
Dim hmEtf As Integer
Dim tmEtf As ETF
Dim imEtfRecLen As Integer  'ETF record length
'Event name
Dim hmEnf As Integer
Dim tmEnf As ENF
Dim tmEnfSrchKey As INTKEY0 'MNF key record image
Dim imEnfRecLen As Integer  'MNF record length
' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmSVef As VEF           'Selling Vehicle
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer     'VEF record length
Dim smVehName As String
Dim tmCef As CEF
Dim hmCef As Integer
Dim imCefRecLen As Integer
Dim tmCefSrchKey As LONGKEY0
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
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smDateTime As String
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBtmCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imEvtType(0 To 14) As Integer
Dim tmPrgInfo() As PRGINFO
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
    Dim llDate As Long
    Dim ilNoWks As Integer
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
    If Trim$(edcNoWks.Text) <> "" Then
        ilNoWks = Val(edcNoWks.Text)
    Else
        ilNoWks = 0
    End If
    'ilDays = 7  'Val(edcNoDays.Text)
    'slEndDate = Format$(gDateValue(slStartDate) + ilDays - 1, "m/d/yy")
    slEndDate = ""
    If Not mOpenMsgFile() Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    slToFile = Trim$(edcTo.Text)
    If Len(slToFile) = 0 Then
        Beep
        edcTo.SetFocus
        Exit Sub
    End If
    'If InStr(slToFile, ":") = 0 Then
    If (InStr(slToFile, ":") = 0) And (Left$(slToFile, 2) <> "\\") Then
        slToFile = sgExportPath & slToFile
    End If
    gAutomationAlertAndLogHandler "* Storing Output into " & slToFile
    
    gAutomationAlertAndLogHandler "* StartDate = " & edcStartDate.Text
    If edcNoWks.Text = "" Then
        gAutomationAlertAndLogHandler "* Weeks = All"
    Else
        gAutomationAlertAndLogHandler "* Weeks = " & edcNoWks.Text
    End If
    
    ilRet = 0
    'On Error GoTo cmcExportErr:
    'hmTo = FreeFile
    'Open slToFile For Output As hmTo
    ilRet = gFileOpen(slToFile, "Output", hmTo)
    If ilRet <> 0 Then
        ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        edcTo.SetFocus
        Exit Sub
    End If
    'Print #hmMsg, "** Storing Output into " & slToFile & " **"
    gAutomationAlertAndLogHandler "* Storing Output into " & slToFile
    'Determine file name from date: YYMMDDn.NY
    'where n= version letter
    Screen.MousePointer = vbHourglass
    smDateTime = Format$(gNow(), "mmddyyyy")
    smDateTime = smDateTime & Format$(gNow(), "hhmmss")
    imExporting = True
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            smVehName = Trim$(slName)
            'Print #hmMsg, "** Generating Data for " & Trim$(slName) & " **"
            gAutomationAlertAndLogHandler "** Generating Data for " & Trim$(slName) & " **"
            lacProcessing.Caption = "Generating Data for " & Trim$(slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            tmVefSrchKey.iCode = ilVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then
                'Obtain last date
                llDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
                slEndDate = Format$(llDate, "m/d/yy")
                If (tmVef.sType = "A") And (gDateValue(slStartDate) > gDateValue(slEndDate)) Then
                    slEndDate = gObtainNextSunday(slStartDate)
                Else
                    If ilNoWks > 0 Then
                        If gDateValue(slStartDate) + 7 * ilNoWks - 1 < gDateValue(slEndDate) Then
                            slEndDate = Format$(gDateValue(slStartDate) + 7 * ilNoWks - 1, "m/d/yy")
                        End If
                    End If
                End If
                If Not mExpRecs(0, "C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType()) Then
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
                'Print #hmMsg, "** Completed " & Trim$(tmVef.sName) & " **"
                gAutomationAlertAndLogHandler "** Completed " & Trim$(tmVef.sName) & " **"
            Else
            End If
        End If
    Next ilLoop
    Close #hmTo
    'Print #hmMsg, "** Completed Export CNC Network Inventory: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** Completed Export CNC Network Inventory: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    Close #hmMsg
    On Error GoTo 0

    lacProcessing.Caption = ""
    Screen.MousePointer = vbDefault
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
Private Sub cmcTo_Click()
    CMDialogBox.DialogTitle = "Export To File"
    CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
    CMDialogBox.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
    CMDialogBox.DefaultExt = ".Txt"
    CMDialogBox.flags = cdlOFNCreatePrompt
    CMDialogBox.Action = 1 'Open dialog
    edcTo.Text = CMDialogBox.fileName
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcNoWks_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus edcNoWks
End Sub
Private Sub edcNoWks_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
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
Private Sub edcTo_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
        'mInitDDE
        'mSendHelpMsg "BT"
    End If
    gCtrlGotFocus ActiveControl
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
    On Error Resume Next
    
    Erase tmPrgInfo
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmEtf)
    btrDestroy hmEtf
    ilRet = btrClose(hmEnf)
    btrDestroy hmEnf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    ilRet = btrClose(hmDlf)
    btrDestroy hmDlf
    
    Set ExpCnCNI = Nothing   'Remove data segment
    
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
'*      Procedure Name:mExpRecs                        *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Export Recs                    *
'*                                                     *
'*******************************************************
Private Function mExpRecs(ilLCFType As Integer, sLCP As String, ilCallCode As Integer, slSDate As String, slEDate As String, slStartTime As String, slEndTime As String, ilEvtType() As Integer) As Integer
'
'   iRet = mExpRecs()
'   Where:
'       iRet (O)- True if record exported,
'                 False if error
'
    Dim ilLoop As Integer
    Dim ilRet As Integer    'Return status
    Dim slStr As String
    Dim slTime As String
    Dim llTime As Long
    Dim slRecord As String
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim slDate As String
    Dim slMoDate As String
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilTerminated As Integer
    Dim ilStartTime0 As Integer
    Dim ilStartTime1 As Integer
    Dim ilVpfIndex As Integer
    Dim ilVehCode As Integer
    Dim ilVeh As Integer
    Dim ilBreakNo As Integer    'Reset to zero for each program
    Dim ilPositionNo As Integer 'Reset to zero for each avail
    Dim llEvtTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilWithinTime As Integer
    Dim llCefCode As Long
    Dim llEvtIDCefCode As Long
    Dim slPrgC As String
    'ReDim slPrgComm(0 To 4) As String
    'ReDim slPrgCommZ(0 To 4) As String
    Dim iZ As Integer
    Dim slAvailComm As String
    Dim ilType As Integer   '1=Prog; 2=Comment; 3=Avail; 4=Spot
    Dim ilDlfDate0 As Integer
    Dim ilDlfDate1 As Integer
    Dim ilDlfFound As Integer
    Dim ilVlfDate0 As Integer
    Dim ilVlfDate1 As Integer
    Dim ilEvtRet As Integer
    Dim ilAirHour As Integer
    Dim ilLocalHour As Integer
    Dim slAdjDate As String
    Dim ilAdjDay As Integer
    Dim ilUpper As Integer
    Dim ilPass As Integer
    Dim ilVlfFd As Integer
    Dim llSTime As Long
    Dim ilSLLC As Integer
    Dim llSEvtTime As Long
    Dim llSCefCode As Long
    ReDim tlLLC(0 To 0) As LLC  'Image
    ReDim tlSLLC(0 To 0) As LLC
    ReDim ilVehicle(0 To 1) As Integer
    Dim doe As Integer
    ilVehicle(0) = ilCallCode
    llStartTime = CLng(gTimeToCurrency(slStartTime, False))
    llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
    'slPrgComm(0) = ""
    'slPrgComm(1) = ""
    'slPrgComm(2) = ""
    'slPrgComm(3) = ""
    'slPrgComm(4) = ""
    'slPrgCommZ(0) = ""
    'slPrgCommZ(1) = "EST"
    'slPrgCommZ(2) = "CST"
    'slPrgCommZ(3) = "PST"
    'slPrgCommZ(4) = "MST"

    For ilPass = 0 To 1 Step 1
        tmVef.iCode = 0
        If ilPass = 0 Then
            llSDate = gDateValue(slSDate)
            llEDate = gDateValue(slEDate)
        Else
            llSDate = gDateValue(gObtainNextMonday(slEDate))
            llEDate = llSDate + 6
        End If
        For llDate = llSDate To llEDate Step 1
            slDate = Format$(llDate, "m/d/yy")
            ilDay = gWeekDayStr(slDate)
            gPackDate slDate, ilLogDate0, ilLogDate1
            For ilVeh = LBound(ilVehicle) To UBound(ilVehicle) - 1 Step 1
                ReDim tmPrgInfo(0 To 0) As PRGINFO
                ilVehCode = ilVehicle(ilVeh)
                If ilVehCode <> tmVef.iCode Then
                    tmVefSrchKey.iCode = ilVehCode
                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                    If (ilRet <> BTRV_ERR_NONE) Then
                        mExpRecs = False
                        Exit Function
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
                        mExpRecs = False
                        Exit Function
                    End If
                End If
                ilDlfFound = False
                'If (((tmVef.sType = "A") Or ((tmVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(1) <> 0)))) Then
                If (((tmVef.sType = "A") Or ((tmVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(0) <> 0)))) Then
                    'Obtain delivery records for date
                    If (ilDay >= 0) And (ilDay <= 4) Then
                        slDay = "0"
                    ElseIf ilDay = 5 Then
                        slDay = "6"
                    Else
                        slDay = "7"
                    End If
                    'Obtain the start date of DLF
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
                        Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) 'And (tmVlf.iAirDay = Val(slDay))
                            If (tmVlf.iAirDay = Val(slDay)) Then
                                ilTerminated = False
                                If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                    If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                        ilTerminated = True
                                    End If
                                End If
                                If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                    ilVlfDate0 = tmVlf.iEffDate(0)
                                    ilVlfDate1 = tmVlf.iEffDate(1)
                                    Exit Do
                                End If
                            End If
                            ilRet = btrGetPrevious(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                End If
                igDOE = igDOE + 1
                If igDOE > 1000 Then
                    DoEvents
                    igDOE = 0
                End If
                
                'gObtainVlf hlVlf, ilVehCode, llDate, tlVlf0(), tlVlf5(), tlVlf6()
                ilDay = gWeekDayStr(slDate)
                gPackDate slDate, ilLogDate0, ilLogDate1
                ReDim tlLLC(0 To 0) As LLC  'Image
                If ilPass = 0 Then
                    ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, slDate, slStartTime, slEndTime, ilEvtType(), tlLLC())
                Else
                    Select Case llDate - llSDate
                        Case 0
                            ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, "TFNMO", slStartTime, slEndTime, ilEvtType(), tlLLC())
                        Case 1
                            ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, "TFNTU", slStartTime, slEndTime, ilEvtType(), tlLLC())
                        Case 2
                            ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, "TFNWE", slStartTime, slEndTime, ilEvtType(), tlLLC())
                        Case 3
                            ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, "TFNTH", slStartTime, slEndTime, ilEvtType(), tlLLC())
                        Case 4
                            ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, "TFNFR", slStartTime, slEndTime, ilEvtType(), tlLLC())
                        Case 5
                            ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, "TFNSA", slStartTime, slEndTime, ilEvtType(), tlLLC())
                        Case 6
                            ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, "TFNSU", slStartTime, slEndTime, ilEvtType(), tlLLC())
                    End Select
                End If
                ilBreakNo = 0
                ilPositionNo = 0
                For ilLoop = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                    gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                    llEvtTime = gTimeToCurrency(tlLLC(ilLoop).sStartTime, False)
                    If llEvtTime > llEndTime Then
                        ilWithinTime = False
                        Exit For
                    End If
                    If llEvtTime >= llStartTime Then
                        ilWithinTime = True
                    End If
                    llCefCode = 0
                    llEvtIDCefCode = 0
                    'tmEnfSrchKey.iCode = tlLLC(ilLoop).iEnfCode
                    'ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    '6/13/06:  Added test for zero
                    If tlLLC(ilLoop).iEnfCode > 0 Then
                        tmEnfSrchKey.iCode = tlLLC(ilLoop).iEnfCode
                        ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Else
                        tmEnf.sName = ""
                    End If
                    If tlLLC(ilLoop).iEtfCode = 1 Then
                        ilBreakNo = 0
                        ilPositionNo = 0
                    ElseIf (tlLLC(ilLoop).iEtfCode >= 2) And (tlLLC(ilLoop).iEtfCode <= 9) Then
                        ilBreakNo = ilBreakNo + 1
                        ilPositionNo = 0
                    End If
                    If ilWithinTime Then
                        ilVlfFd = False
                        If tlLLC(ilLoop).iEtfCode = 1 Then    'Program
                            llCefCode = tlLLC(ilLoop).lCefCode
                            llEvtIDCefCode = tlLLC(ilLoop).lEvtIDCefCode
                            ilType = 1
                            If (tmVef.sType = "A") Or (ilDlfFound) Then
                                'Obtain delivery entry to see is prog is sent
                                tmDlfSrchKey.iVefCode = ilVehCode
                                tmDlfSrchKey.sAirDay = slDay
                                tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                gPackTime tlLLC(ilLoop).sStartTime, tmDlfSrchKey.iAirTime(0), tmDlfSrchKey.iAirTime(1)
                                ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = ilStartTime0) And (tmDlf.iAirTime(1) = ilStartTime1)
                                    ilTerminated = False
                                    If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                        ilTerminated = True
                                    Else
                                        If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                            If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                ilTerminated = True
                                            End If
                                        End If
                                    End If
                                    If Not ilTerminated Then
                                        If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                            'If slFor = "C" Then
                                                If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                                                    tmDlf.iMnfFeed = 0
                                                    '6/1/16: Replace GoSub
                                                    'GoSub lProcProg
                                                    mProcProg ilDay, ilAirHour, ilLocalHour, ilLogDate0, ilLogDate1, slAdjDate, ilAdjDay, ilLoop, slDate, tlLLC(), ilBreakNo, ilPass, ilVpfIndex, slRecord
                                                    DoEvents
                                                End If
                                            'End If
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            Else
                                tmDlf.iVefCode = ilVehCode
                                tmDlf.iLocalTime(0) = ilStartTime0
                                tmDlf.iLocalTime(1) = ilStartTime1
                                tmDlf.iFeedTime(0) = ilStartTime0
                                tmDlf.iFeedTime(1) = ilStartTime1
                                tmDlf.iAirTime(0) = ilStartTime0
                                tmDlf.iAirTime(1) = ilStartTime1
                                tmDlf.sZone = ""
                                tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
                                tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
                                tmDlf.sProgCode = ""
                                tmDlf.iMnfFeed = 0
                                tmDlf.sBus = ""
                                tmDlf.sSchedule = ""
                                tmDlf.iMnfSubFeed = 0
                                '6/1/16: Replace GoSub
                                'GoSub lProcProg
                                mProcProg ilDay, ilAirHour, ilLocalHour, ilLogDate0, ilLogDate1, slAdjDate, ilAdjDay, ilLoop, slDate, tlLLC(), ilBreakNo, ilPass, ilVpfIndex, slRecord
                                'DoEvents
                                igDOE = igDOE + 1
                                If igDOE > 500 Then
                                    DoEvents
                                    igDOE = 0
                                End If
                            End If
                        ElseIf (tlLLC(ilLoop).iEtfCode >= 2) And (tlLLC(ilLoop).iEtfCode <= 9) Then 'Avail
                            llCefCode = tlLLC(ilLoop).lCefCode
                            llEvtIDCefCode = tlLLC(ilLoop).lEvtIDCefCode
                            tmCefSrchKey.lCode = llEvtIDCefCode 'llCefCode
                            imCefRecLen = Len(tmCef)
                            ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                'slAvailComm = Trim$(Left$(tmCef.sComment, tmCef.iStrLen)) 'tlCef.sComment
                                slAvailComm = gStripChr0(tmCef.sComment)
                            Else
                                slAvailComm = ""
                            End If
                            ilType = 3
                            If tmVef.sType = "A" Then
                                tmVlfSrchKey1.iAirCode = ilVehCode
                                tmVlfSrchKey1.iAirDay = Val(slDay)
                                tmVlfSrchKey1.iEffDate(0) = ilVlfDate0
                                tmVlfSrchKey1.iEffDate(1) = ilVlfDate1
                                tmVlfSrchKey1.iAirTime(0) = ilStartTime0
                                tmVlfSrchKey1.iAirTime(1) = ilStartTime1
                                tmVlfSrchKey1.iAirPosNo = 0
                                tmVlfSrchKey1.iAirSeq = 1
                                ilRet = btrGetGreaterOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) And (tmVlf.iAirDay = Val(slDay)) And (tmVlf.iEffDate(0) = ilVlfDate0) And (tmVlf.iEffDate(1) = ilVlfDate1) And (tmVlf.iAirTime(0) = ilStartTime0) And (tmVlf.iAirTime(1) = ilStartTime1)
                                    ilTerminated = False
                                    If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                        If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                            ilTerminated = True
                                        End If
                                    End If
                                    If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                        ilVlfFd = True
                                        ReDim tlSLLC(0 To 0) As LLC  'Image
                                        If ilPass = 0 Then
                                            ilEvtRet = gBuildEventDay(ilLCFType, sLCP, tmVlf.iSellCode, slDate, slStartTime, slEndTime, ilEvtType(), tlSLLC())
                                        Else
                                            Select Case llDate - llSDate
                                                Case 0
                                                    ilEvtRet = gBuildEventDay(ilLCFType, sLCP, tmVlf.iSellCode, "TFNMO", slStartTime, slEndTime, ilEvtType(), tlSLLC())
                                                Case 1
                                                    ilEvtRet = gBuildEventDay(ilLCFType, sLCP, tmVlf.iSellCode, "TFNTU", slStartTime, slEndTime, ilEvtType(), tlSLLC())
                                                Case 2
                                                    ilEvtRet = gBuildEventDay(ilLCFType, sLCP, tmVlf.iSellCode, "TFNWE", slStartTime, slEndTime, ilEvtType(), tlSLLC())
                                                Case 3
                                                    ilEvtRet = gBuildEventDay(ilLCFType, sLCP, tmVlf.iSellCode, "TFNTH", slStartTime, slEndTime, ilEvtType(), tlSLLC())
                                                Case 4
                                                    ilEvtRet = gBuildEventDay(ilLCFType, sLCP, tmVlf.iSellCode, "TFNFR", slStartTime, slEndTime, ilEvtType(), tlSLLC())
                                                Case 5
                                                    ilEvtRet = gBuildEventDay(ilLCFType, sLCP, tmVlf.iSellCode, "TFNSA", slStartTime, slEndTime, ilEvtType(), tlSLLC())
                                                Case 6
                                                    ilEvtRet = gBuildEventDay(ilLCFType, sLCP, tmVlf.iSellCode, "TFNSU", slStartTime, slEndTime, ilEvtType(), tlSLLC())
                                            End Select
                                        End If
                                        Exit Do
                                    End If
                                    ilRet = btrGetNext(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            Else
                            End If
                            If (tmVef.sType = "A") Or (ilDlfFound) Then

                                'Obtain delivery entry to see is avail is sent
                                tmDlfSrchKey.iVefCode = ilVehCode
                                tmDlfSrchKey.sAirDay = slDay
                                tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                tmDlfSrchKey.iAirTime(0) = ilStartTime0
                                tmDlfSrchKey.iAirTime(1) = ilStartTime1
                                ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = ilStartTime0) And (tmDlf.iAirTime(1) = ilStartTime1)
                                    ilTerminated = False
                                    If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                        ilTerminated = True
                                    Else
                                        If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                            If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                ilTerminated = True
                                            End If
                                        End If
                                    End If
                                    If Not ilTerminated Then
                                        If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                            'If slFor = "C" Then
                                                If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                                                    tmDlf.iMnfFeed = 0
                                                    '6/1/16: Replaced GoSub
                                                    'GoSub lProcAvail
                                                    mProcAvail ilDay, ilAirHour, ilLocalHour, ilLogDate0, ilLogDate1, slAdjDate, ilAdjDay, tlLLC(), tlSLLC(), slDate, ilBreakNo, slAvailComm, ilPass, ilVlfFd, ilVpfIndex, ilLoop, slRecord
                                                End If
                                            'End If
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            Else
                                tmDlf.iVefCode = ilVehCode
                                tmDlf.iLocalTime(0) = ilStartTime0
                                tmDlf.iLocalTime(1) = ilStartTime1
                                tmDlf.iFeedTime(0) = ilStartTime0
                                tmDlf.iFeedTime(1) = ilStartTime1
                                tmDlf.iAirTime(0) = ilStartTime0
                                tmDlf.iAirTime(1) = ilStartTime1
                                tmDlf.sZone = ""
                                tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
                                tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
                                tmDlf.sProgCode = ""
                                tmDlf.iMnfFeed = 0
                                tmDlf.sBus = ""
                                tmDlf.sSchedule = ""
                                tmDlf.iMnfSubFeed = 0
                                '6/1/16: Replaced GoSub
                                'GoSub lProcAvail
                                mProcAvail ilDay, ilAirHour, ilLocalHour, ilLogDate0, ilLogDate1, slAdjDate, ilAdjDay, tlLLC(), tlSLLC(), slDate, ilBreakNo, slAvailComm, ilPass, ilVlfFd, ilVpfIndex, ilLoop, slRecord
                            End If
                        ElseIf (tlLLC(ilLoop).iEtfCode > 9) Then
                            'Handle like program
                            ilType = 2
                            llCefCode = tlLLC(ilLoop).lCefCode
                            If (tmVef.sType = "A") Or (ilDlfFound) Then
                                'Obtain delivery entry to see is prog is sent
                                tmDlfSrchKey.iVefCode = ilVehCode
                                tmDlfSrchKey.sAirDay = slDay
                                tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                tmDlfSrchKey.iAirTime(0) = ilStartTime0
                                tmDlfSrchKey.iAirTime(1) = ilStartTime1
                                ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = ilStartTime0) And (tmDlf.iAirTime(1) = ilStartTime1)
                                    ilTerminated = False
                                    If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                        ilTerminated = True
                                    Else
                                        If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                            If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                ilTerminated = True
                                            End If
                                        End If
                                    End If
                                    If Not ilTerminated Then
                                        If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                            'If slFor = "C" Then
                                                If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                                                    tmDlf.iMnfFeed = 0
                                                    '6/1/16: Replace GoSub
                                                    'GoSub lProcProg
                                                    mProcProg ilDay, ilAirHour, ilLocalHour, ilLogDate0, ilLogDate1, slAdjDate, ilAdjDay, ilLoop, slDate, tlLLC(), ilBreakNo, ilPass, ilVpfIndex, slRecord
                                                    DoEvents
                                                End If
                                            'End If
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            Else
                                tmDlf.iVefCode = ilVehCode
                                gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                tmDlf.iLocalTime(0) = ilStartTime0
                                tmDlf.iLocalTime(1) = ilStartTime1
                                tmDlf.iFeedTime(0) = ilStartTime0
                                tmDlf.iFeedTime(1) = ilStartTime1
                                tmDlf.iAirTime(0) = ilStartTime0
                                tmDlf.iAirTime(1) = ilStartTime1
                                tmDlf.sZone = ""
                                tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
                                tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
                                tmDlf.sProgCode = ""
                                tmDlf.iMnfFeed = 0
                                tmDlf.sBus = ""
                                tmDlf.sSchedule = ""
                                tmDlf.iMnfSubFeed = 0
                                '6/1/16: Replace GoSub
                                'GoSub lProcProg
                                mProcProg ilDay, ilAirHour, ilLocalHour, ilLogDate0, ilLogDate1, slAdjDate, ilAdjDay, ilLoop, slDate, tlLLC(), ilBreakNo, ilPass, ilVpfIndex, slRecord
                                DoEvents
                            End If
                        End If
                    End If
                Next ilLoop
            Next ilVeh
        Next llDate
    Next ilPass
    mExpRecs = True
    Exit Function
'lProcProg:
'    '6/1/16: Replaced GoSub
'    'GoSub lProcAdjDate
'    mProcAdjDate ilDay, ilAirHour, ilLocalHour, ilLogDate0, ilLogDate1, slAdjDate, ilAdjDay
'    slRecord = smDateTime & sgCnCDelimiter
'    'Vehicle Type
'    slRecord = slRecord & tmVef.sType & sgCnCDelimiter
'    'Event Type
'    If tlLLC(ilLoop).iEtfCode = 1 Then
'        slRecord = slRecord & "PR" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "CM" & sgCnCDelimiter
'    End If
'    If tlLLC(ilLoop).iEtfCode = 1 Then
'        'llCefCode = tlLLC(ilLoop).lCefCode
'        llCefCode = tlLLC(ilLoop).lEvtIDCefCode
'        tmCefSrchKey.lCode = llCefCode
'        imCefRecLen = Len(tmCef)
'        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'        '    slPrgComm = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))'tlCef.sComment
'        '    slPrgComm = Trim$(slPrgComm) & Trim$(tmDlf.sBus)
'        'Else
'        '    slPrgComm = ""
'            'slPrgC = Trim$(Left$(tmCef.sComment, tmCef.iStrLen)) 'tlCef.sComment
'            slPrgC = gStripChr0(tmCef.sComment)
'            slPrgC = Trim$(slPrgC) & Trim$(tmDlf.sBus)
'        Else
'            slPrgC = ""
'        End If
'        'If Trim$(tmDlf.sZone) <> "" Then
'        '    For iZ = 1 To 4 Step 1
'        '       If StrComp(Trim$(tmDlf.sZone), slPrgCommZ(iZ), 1) = 0 Then
'        '           slPrgComm(iZ) = slPrgC
'        '           Exit For
'        '       End If
'        '    Next iZ
'        'Else
'        '    slPrgComm(0) = slPrgC
'        'End If
'        ilUpper = UBound(tmPrgInfo)
'        gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
'        tmPrgInfo(ilUpper).lLocalSTime = gTimeToLong(slTime, False)
'        tmPrgInfo(ilUpper).lLocalETime = tmPrgInfo(ilUpper).lLocalSTime + CLng(gLengthToCurrency(tlLLC(ilLoop).sLength))
'        tmPrgInfo(ilUpper).sZone = tmDlf.sZone
'        tmPrgInfo(ilUpper).sPrgID = slPrgC
'        ReDim Preserve tmPrgInfo(0 To ilUpper + 1) As PRGINFO
'    Else
'        llCefCode = tlLLC(ilLoop).lEvtIDCefCode
'        tmCefSrchKey.lCode = llCefCode
'        imCefRecLen = Len(tmCef)
'        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'        '    slPrgComm = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))'tlCef.sComment
'        '    slPrgComm = Trim$(slPrgComm) & Trim$(tmDlf.sBus)
'        'Else
'        '    slPrgComm = ""
'            'slPrgC = Trim$(Left$(tmCef.sComment, tmCef.iStrLen)) 'tlCef.sComment
'            slPrgC = gStripChr0(tmCef.sComment)
'            slPrgC = Trim$(slPrgC) & Trim$(tmDlf.sBus)
'        Else
'            slPrgC = ""
'        End If
'    End If
'    'Network code
'    slRecord = slRecord & Trim$(tmVef.sCodeStn) & sgCnCDelimiter
'    'Network ID
'    slRecord = slRecord & Trim$(str$(tmVef.iCode)) & sgCnCDelimiter
'    'Vehicle Name
'    slRecord = slRecord & Trim$(tmVef.sName) & sgCnCDelimiter
'    'Effective date
'    slMoDate = gObtainPrevMonday(slDate)
'    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
'    If ilPass = 0 Then
'        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
'    Else
'        slRecord = slRecord & sgCnCDelimiter
'    End If
'    'Program Dates
'    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
'    If ilPass = 0 Then
'        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
'    Else
'        slRecord = slRecord & sgCnCDelimiter
'    End If
'    'Event Names
'    slRecord = slRecord & Trim$(tmEnf.sName) & sgCnCDelimiter
'    'Program Event ID
'    'slRecord = slRecord & Left$(slPrgComm, 5) & sgCnCDelimiter
'    slRecord = slRecord & Trim$(slPrgC) & sgCnCDelimiter
'    'Program Category for RADAR
'    slRecord = slRecord & sgCnCDelimiter
'    'Length in seconds
'    slRecord = slRecord & Trim$(str$(CLng(gLengthToCurrency(tlLLC(ilLoop).sLength)))) & sgCnCDelimiter
'    'Feed Time
'    'slTime = Format$(gConvertTime(tlLLC(ilLoop).sStartTime), "hhmmss")
'    gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
'    slTime = Format$(gConvertTime(slTime), "hhmmss")
'    slRecord = slRecord & slTime & sgCnCDelimiter
'    'Feed Zone
'    slRecord = slRecord & Trim$(tmDlf.sZone) & sgCnCDelimiter
'    ''Time Zone
'    'slStr = ""
'    'If Trim$(tmDlf.sZone) <> "" Then
'    '    For ilZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
'    '        If StrComp(Trim$(tmDlf.sZone), Trim$(tgVpf(ilVpfIndex).sGZone(ilZone)), 1) = 0 Then
'    '            If tgVpf(ilVpfIndex).sGFed(ilZone) = "Y" Then
'    '                slStr = Trim$(tmDlf.sZone)
'    '            End If
'    '            Exit For
'    '        End If
'    '    Next ilZone
'    'End If
'    'slRecord = slRecord & slStr & sgCnCDelimiter
'    'Break No
'    slRecord = slRecord & Trim$(str$(ilBreakNo)) & sgCnCDelimiter
'    'Program Event ID
'    'slRecord = slRecord & Left$(slPrgComm, 5) & sgCnCDelimiter
'    slRecord = slRecord & Trim$(slPrgC) & sgCnCDelimiter
'    'Units
'    slRecord = slRecord & sgCnCDelimiter
'    'Network basis
'    If tgVpf(ilVpfIndex).sSSellOut = "U" Then
'        slRecord = slRecord & "U" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "S" & sgCnCDelimiter
'    End If
'    'Network basis change date
'    slRecord = slRecord & sgCnCDelimiter
'    'Total Network Units/Length
'    slRecord = slRecord & sgCnCDelimiter
'    'airday
'    If ilAdjDay = 0 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 1 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 2 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 3 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 4 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 5 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 6 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    'Relative time
'    'slRecord = slRecord & Trim$(Str$(CLng(gLengthToCurrency(tlLLC(ilLoop).sLength)))) & sgCnCDelimiter
'    slTime = Format$(gConvertTime(tlLLC(ilLoop).sStartTime), "hhmmss")
'    slRecord = slRecord & slTime & sgCnCDelimiter
'    'Network Length
'    slRecord = slRecord & sgCnCDelimiter
'    'Event Effective date
'    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
'    If ilPass = 0 Then
'        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
'    Else
'        slRecord = slRecord & sgCnCDelimiter
'    End If
'    'Avail Type
'    slRecord = slRecord & sgCnCDelimiter
'    'Selling Network code
'    slRecord = slRecord & sgCnCDelimiter
'    'Selling Network ID
'    slRecord = slRecord & sgCnCDelimiter
'    'Selling Vehicle Name
'    slRecord = slRecord & sgCnCDelimiter
'    'Selling Avail Time
'    slRecord = slRecord & sgCnCDelimiter
'    'Selling Day
'    slRecord = slRecord & sgCnCDelimiter
'    'Link Effective Start Date
'    slRecord = slRecord & sgCnCDelimiter
'    'Link Effective End Date
'    slRecord = slRecord & sgCnCDelimiter
'    'Link Internal ID
'    slRecord = slRecord & sgCnCDelimiter
'    'Selling Event ID
'    slRecord = slRecord
'    Print #hmTo, slRecord
'    Return
'lProcAvail:
'    '6/1/16: Replaced GoSub
'    'GoSub lProcAdjDate
'    mProcAdjDate ilDay, ilAirHour, ilLocalHour, ilLogDate0, ilLogDate1, slAdjDate, ilAdjDay
'    slRecord = smDateTime & sgCnCDelimiter
'    'Vehicle Type
'    slRecord = slRecord & tmVef.sType & sgCnCDelimiter
'    'Event Type
'    Select Case tlLLC(ilLoop).iEtfCode
'        Case 2
'            'Test Event Name
'            If Left$(tmEnf.sName, 1) = "N" Then
'                slRecord = slRecord & "NA" & sgCnCDelimiter
'            Else
'                slRecord = slRecord & "LA" & sgCnCDelimiter
'            End If
'        Case 3
'            slRecord = slRecord & "OB" & sgCnCDelimiter
'        'Case 4
'        '    slRecord = slRecord & "CB" & sgCnCDelimiter
'        Case 5
'            slRecord = slRecord & "CB" & sgCnCDelimiter
'        Case 8
'            slRecord = slRecord & "PS" & sgCnCDelimiter
'    End Select
'    'Network code
'    slRecord = slRecord & Trim$(tmVef.sCodeStn) & sgCnCDelimiter
'    'Network ID
'    slRecord = slRecord & Trim$(str$(tmVef.iCode)) & sgCnCDelimiter
'    'Vehicle Name
'    slRecord = slRecord & Trim$(tmVef.sName) & sgCnCDelimiter
'    'Effective date
'    slMoDate = gObtainPrevMonday(slDate)
'    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
'    If ilPass = 0 Then
'        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
'    Else
'        slRecord = slRecord & sgCnCDelimiter
'    End If
'    'Program Dates
'    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
'    If ilPass = 0 Then
'        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
'    Else
'        slRecord = slRecord & sgCnCDelimiter
'    End If
'    'Event Names
'    slRecord = slRecord & Trim$(tmEnf.sName) & sgCnCDelimiter
'    'Program Event ID
'    ''If slAvailComm = "" Then
'    '    'slRecord = slRecord & Left$(slPrgComm, 5) & sgCnCDelimiter
'    '    'slRecord = slRecord & Trim$(slPrgComm) & sgCnCDelimiter
'    'If Trim$(tmDlf.sZone) <> "" Then
'    '    For iZ = 1 To 4 Step 1
'    '        If StrComp(Trim$(tmDlf.sZone), slPrgCommZ(iZ), 1) = 0 Then
'    '            slRecord = slRecord & Trim$(slPrgComm(iZ)) & sgCnCDelimiter
'    '            Exit For
'    '        End If
'    '    Next iZ
'    'Else
'    '    slRecord = slRecord & Trim$(slPrgComm(0)) & sgCnCDelimiter
'    'End If
'    ''Else
'    ''    slRecord = slRecord & Left$(slAvailComm, 5) & sgCnCDelimiter
'    ''End If
'    slPrgC = ""
'    gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
'    llTime = gTimeToLong(slTime, False)
'    For iZ = 0 To UBound(tmPrgInfo) - 1 Step 1
'        If (llTime >= tmPrgInfo(iZ).lLocalSTime) And (llTime < tmPrgInfo(iZ).lLocalETime) And (Trim$(tmDlf.sZone) = Trim$(tmPrgInfo(iZ).sZone)) Then
'            slPrgC = tmPrgInfo(iZ).sPrgID
'            Exit For
'        End If
'    Next iZ
'    slRecord = slRecord & Trim$(slPrgC) & sgCnCDelimiter
'    'Prgram Category
'    slRecord = slRecord & sgCnCDelimiter
'    'Length in seconds
'    slRecord = slRecord & Trim$(str$(CLng(gLengthToCurrency(tlLLC(ilLoop).sLength)))) & sgCnCDelimiter
'    'Feed Time
'    'slTime = Format$(gConvertTime(tlLLC(ilLoop).sStartTime), "hhmmss")
'    gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
'    slTime = Format$(gConvertTime(slTime), "hhmmss")
'    slRecord = slRecord & slTime & sgCnCDelimiter
'    'Feed Zone
'    slRecord = slRecord & Trim$(tmDlf.sZone) & sgCnCDelimiter
'    ''Time Zone
'    'slStr = ""
'    'If Trim$(tmDlf.sZone) <> "" Then
'    '    For ilZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
'    '        If StrComp(Trim$(tmDlf.sZone), Trim$(tgVpf(ilVpfIndex).sGZone(ilZone)), 1) = 0 Then
'    '            If tgVpf(ilVpfIndex).sGFed(ilZone) = "Y" Then
'    '                slStr = Trim$(tmDlf.sZone)
'    '            End If
'    '            Exit For
'    '        End If
'    '    Next ilZone
'    'End If
'    'slRecord = slRecord & slStr & sgCnCDelimiter
'    'Break No
'    slRecord = slRecord & Trim$(str$(ilBreakNo)) & sgCnCDelimiter
'    'Avail Event ID
'    'slRecord = slRecord & Left$(slAvailComm, 5) & sgCnCDelimiter
'    slRecord = slRecord & Trim$(slAvailComm) & Trim$(tmDlf.sBus) & sgCnCDelimiter
'    'Units
'    slRecord = slRecord & Trim$(str$(tlLLC(ilLoop).iUnits)) & sgCnCDelimiter
'    'Network basis
'    If tgVpf(ilVpfIndex).sSSellOut = "U" Then
'        slRecord = slRecord & "U" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "S" & sgCnCDelimiter
'    End If
'    'Network basis change date
'    slRecord = slRecord & sgCnCDelimiter
'    'Total Network Units/Length
'    slRecord = slRecord & sgCnCDelimiter
'    'airday
'    If ilAdjDay = 0 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 1 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 2 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 3 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 4 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 5 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    If ilAdjDay = 6 Then
'        slRecord = slRecord & "Y" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & "N" & sgCnCDelimiter
'    End If
'    'Relative time
'    'slRecord = slRecord & Trim$(Str$(CLng(gLengthToCurrency(tlLLC(ilLoop).sLength)))) & sgCnCDelimiter
'    slTime = Format$(gConvertTime(tlLLC(ilLoop).sStartTime), "hhmmss")
'    slRecord = slRecord & slTime & sgCnCDelimiter
'    'Network Length
'    slRecord = slRecord & sgCnCDelimiter
'    'Event Effective date
'    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
'    If ilPass = 0 Then
'        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
'    Else
'        slRecord = slRecord & sgCnCDelimiter
'    End If
'    If tlLLC(ilLoop).iEtfCode = 2 Then
'        If Left$(tmEnf.sName, 1) = "N" Then
'            slRecord = slRecord & "N" & sgCnCDelimiter
'        Else
'            slRecord = slRecord & "L" & sgCnCDelimiter
'        End If
'    Else
'        slRecord = slRecord & sgCnCDelimiter
'    End If
'    If ilVlfFd Then
'        If tmVlf.iSellCode <> tmSVef.iCode Then
'            tmVefSrchKey.iCode = tmVlf.iSellCode
'            ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'        End If
'        'Selling Network code
'        slRecord = slRecord & Trim$(tmSVef.sCodeStn) & sgCnCDelimiter
'        'Selling Network ID
'        slRecord = slRecord & Trim$(str$(tmSVef.iCode)) & sgCnCDelimiter
'        'Selling Vehicle Name
'        slRecord = slRecord & Trim$(tmSVef.sName) & sgCnCDelimiter
'        'Selling Avail Time
'        gUnpackTime tmVlf.iSellTime(0), tmVlf.iSellTime(1), "A", "1", slTime
'        slTime = Format$(gConvertTime(slTime), "hhmmss")
'        slRecord = slRecord & slTime & sgCnCDelimiter
'        'Selling Day (0=Mon-Fri; 6=Sa; 7= Su)
'        slRecord = slRecord & Trim$(str$(tmVlf.iSellDay)) & sgCnCDelimiter
'        'Link Effective Start Date
'        gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slStr
'        slStr = Format$(gDateValue(slStr), "mmddyyyy")
'        slRecord = slRecord & slStr & sgCnCDelimiter
'        'Link Effective End Date
'        gUnpackDate tmVlf.iTermDate(0), tmVlf.iTermDate(1), slStr
'        If slStr <> "" Then
'            slStr = Format$(gDateValue(slStr), "mmddyyyy")
'        End If
'        slRecord = slRecord & slStr & sgCnCDelimiter
'        'Link Internal ID
'        slRecord = slRecord & Trim$(str$(tmVlf.lCode)) & sgCnCDelimiter
'        'Program Number
'        llSCefCode = 0
'        llSTime = gTimeToCurrency(slTime, False)
'        For ilSLLC = LBound(tlSLLC) To UBound(tlSLLC) - 1 Step 1
'            llSEvtTime = gTimeToCurrency(tlSLLC(ilSLLC).sStartTime, False)
'            If llSEvtTime > llSTime Then
'                Exit For
'            ElseIf llSEvtTime = llSTime Then
'                If (tlSLLC(ilSLLC).iEtfCode >= 2) And (tlSLLC(ilSLLC).iEtfCode <= 9) Then 'Avail
'                    'If tlSLLC(ilSLLC).lCefCode > 0 Then
'                    '    llSCefCode = tlSLLC(ilSLLC).lCefCode
'                    'End If
'                    If tlSLLC(ilSLLC).lEvtIDCefCode > 0 Then
'                        llSCefCode = tlSLLC(ilSLLC).lEvtIDCefCode
'                    End If
'                    Exit For
'                End If
'            End If
'            'If tlSLLC(ilSLLC).iEtfCode = 1 Then    'Program
'            '    llSCefCode = tlSLLC(ilSLLC).lCefCode
'            'End If
'        Next ilSLLC
'        If llSCefCode > 0 Then
'            tmCefSrchKey.lCode = llSCefCode
'            imCefRecLen = Len(tmCef)
'            ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'            If ilRet = BTRV_ERR_NONE Then
'                'slStr = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
'                slStr = gStripChr0(tmCef.sComment)
'                'slStr = Left$(slStr, 5)
'            Else
'                slStr = ""
'            End If
'        Else
'            slStr = ""
'        End If
'        slRecord = slRecord & slStr
'    Else
'        'Selling Network code
'        slRecord = slRecord & sgCnCDelimiter
'        'Selling Network ID
'        slRecord = slRecord & sgCnCDelimiter
'        'Selling Vehicle Name
'        slRecord = slRecord & sgCnCDelimiter
'        'Selling Avail Time
'        slRecord = slRecord & sgCnCDelimiter
'        'Selling Day
'        slRecord = slRecord & sgCnCDelimiter
'        'Link Effective Start Date
'        slRecord = slRecord & sgCnCDelimiter
'        'Link Effective End Date
'        slRecord = slRecord & sgCnCDelimiter
'        'Link Internal ID
'        slRecord = slRecord & sgCnCDelimiter
'        'Selling Event ID
'        slRecord = slRecord
'    End If
'    Print #hmTo, slRecord
'    Return
'lProcAdjDate:
'    ilAdjDay = ilDay
'    'Test if Air time is AM and Local Time is PM. If so, adjust date
'    ilAirHour = tmDlf.iAirTime(1) \ 256  'Obtain month
'    ilLocalHour = tmDlf.iLocalTime(1) \ 256  'Obtain month
'    If (ilAirHour < 6) And (ilLocalHour > 17) Then
'        'If monday convert to next sunday- this is wrong but the same spot
'        'runs each sunday (the spot should have show on the previous week sunday)
'        'If not monday, then subtract one day
'        gUnpackDate ilLogDate0, ilLogDate1, slAdjDate
'        If gWeekDayStr(slAdjDate) = 0 Then
'            slAdjDate = gObtainNextSunday(slAdjDate)
'        Else
'            slAdjDate = gDecOneDay(slAdjDate)
'        End If
'        ilAdjDay = gWeekDayStr(slAdjDate)
'    End If
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

    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imLBtmCDCtrls = 1
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0

    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCNI
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)
    hmEtf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmEtf, "", sgDBPath & "Etf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCNI
    On Error GoTo 0
    imEtfRecLen = Len(tmEtf)
    hmEnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmEnf, "", sgDBPath & "Enf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCNI
    On Error GoTo 0
    imEnfRecLen = Len(tmEnf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCNI
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmCef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCNI
    On Error GoTo 0
    imCefRecLen = Len(tmCef)
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCNI
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)
    hmDlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmDlf, "", sgDBPath & "Dlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCNI
    On Error GoTo 0
    imDlfRecLen = Len(tmDlf)
    'Populate arrays to determine if records exist
    mVehPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        'mTerminate
        Exit Sub
    End If
    For ilLoop = LBound(imEvtType) To UBound(imEvtType) Step 1
        imEvtType(ilLoop) = True
    Next ilLoop
    imEvtType(0) = False 'Don't include library names
    imEvtType(10) = False
    imEvtType(11) = False
    imEvtType(12) = False
    imEvtType(13) = False
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    slStr = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slStr)
    slStr = gObtainPrevMonday(slStr)
    slStr = gObtainPrevMonday(slStr)
    edcStartDate.Text = slStr
    'gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    'pbcCalendar_Paint   'mBoxCalDate called within paint
    'lacDate.Visible = False
    gCenterStdAlone ExpCnCNI
    Screen.MousePointer = vbDefault
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    'imcHelp.Picture = Traffic!imcHelp.Picture
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
    ilRet = 0
    'slToFile = sgExportPath & "ExpCncNI.Txt"
    slToFile = sgDBPath & "Messages\" & "ExpCncNI.Txt"
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
    
    'Print #hmMsg, "** Export CNC Network Inventory: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** Export CNC Network Inventory: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
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
    Dim ilPos As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer

    slCommand = sgCommandStr  'Command$
    ilPos = InStr(1, slCommand, "/CS", 1)
'    If (StrComp(slCommand, "Debug", 1) = 0) Or (ilPos > 0) Then
'        igStdAloneMode = True 'False  'Switch from/to stand alone mode
'        sgCallAppName = ""
'        slStr = "Guide"
'        ilTestSystem = False
'        ilPos = InStr(1, slCommand, "/S:Test", 1)
'        If ilPos > 0 Then
'            ilTestSystem = True
'        End If
'        ilPos = InStr(1, slCommand, "/S:Prod", 1)
'        If ilPos > 0 Then
'            ilTestSystem = False
'        End If
'        slCommand = "Traffic\Guide"
'    Else
'        igStdAloneMode = False  'Switch from/to stand alone mode
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
        'imShowHelpMsg = True
        'ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
        'If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
        '    imShowHelpMsg = False
        'End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone ExpCnCNI, slStr, ilTestSystem
    igCmmCallSource = Val(slStr)
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

    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpCnCNI
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
    Dim ilVff As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    
    'ilRet = gPopUserVehicleBox(ExpCncNI, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    ilRet = gPopUserVehicleBox(ExpCnCNI, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExpCnCNI
        On Error GoTo 0
    End If

    For ilLoop = LBound(tgUserVehicle) To UBound(tgUserVehicle) - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If ilVefCode = tgVff(ilVff).iVefCode Then
                If tgVff(ilVff).sExportCnCNetInv = "Y" Then
                    lbcVehicle.Selected(ilLoop) = True
                End If
                Exit For
            End If
        Next ilVff
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
    plcScreen.Print "Export Network Inventory"
End Sub

Private Sub mProcAdjDate(ilDay As Integer, ilAirHour As Integer, ilLocalHour As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, slAdjDate As String, ilAdjDay As Integer)
    ilAdjDay = ilDay
    'Test if Air time is AM and Local Time is PM. If so, adjust date
    ilAirHour = tmDlf.iAirTime(1) \ 256  'Obtain month
    ilLocalHour = tmDlf.iLocalTime(1) \ 256  'Obtain month
    If (ilAirHour < 6) And (ilLocalHour > 17) Then
        'If monday convert to next sunday- this is wrong but the same spot
        'runs each sunday (the spot should have show on the previous week sunday)
        'If not monday, then subtract one day
        gUnpackDate ilLogDate0, ilLogDate1, slAdjDate
        If gWeekDayStr(slAdjDate) = 0 Then
            slAdjDate = gObtainNextSunday(slAdjDate)
        Else
            slAdjDate = gDecOneDay(slAdjDate)
        End If
        ilAdjDay = gWeekDayStr(slAdjDate)
    End If
End Sub

Private Sub mProcAvail(ilDay As Integer, ilAirHour As Integer, ilLocalHour As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, slAdjDate As String, ilAdjDay As Integer, tlLLC() As LLC, tlSLLC() As LLC, slDate As String, ilBreakNo As Integer, slAvailComm As String, ilPass As Integer, ilVlfFd As Integer, ilVpfIndex As Integer, ilLoop As Integer, slRecord As String)
    Dim slMoDate As String
    Dim slPrgC As String
    Dim slTime As String
    Dim llTime As Long
    Dim iZ As Integer
    Dim ilSLLC As Integer
    Dim llSCefCode As Long
    Dim llSTime As Long
    Dim llSEvtTime As Long
    Dim slStr As String
    Dim ilRet As Integer
    
    '6/1/16: Replaced GoSub
    'GoSub lProcAdjDate
    mProcAdjDate ilDay, ilAirHour, ilLocalHour, ilLogDate0, ilLogDate1, slAdjDate, ilAdjDay
    slRecord = smDateTime & sgCnCDelimiter
    'Vehicle Type
    slRecord = slRecord & tmVef.sType & sgCnCDelimiter
    'Event Type
    Select Case tlLLC(ilLoop).iEtfCode
        Case 2
            'Test Event Name
            If Left$(tmEnf.sName, 1) = "N" Then
                slRecord = slRecord & "NA" & sgCnCDelimiter
            Else
                slRecord = slRecord & "LA" & sgCnCDelimiter
            End If
        Case 3
            slRecord = slRecord & "OB" & sgCnCDelimiter
        'Case 4
        '    slRecord = slRecord & "CB" & sgCnCDelimiter
        Case 5
            slRecord = slRecord & "CB" & sgCnCDelimiter
        Case 8
            slRecord = slRecord & "PS" & sgCnCDelimiter
    End Select
    'Network code
    slRecord = slRecord & Trim$(tmVef.sCodeStn) & sgCnCDelimiter
    'Network ID
    slRecord = slRecord & Trim$(str$(tmVef.iCode)) & sgCnCDelimiter
    'Vehicle Name
    slRecord = slRecord & Trim$(tmVef.sName) & sgCnCDelimiter
    'Effective date
    slMoDate = gObtainPrevMonday(slDate)
    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
    If ilPass = 0 Then
        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
    Else
        slRecord = slRecord & sgCnCDelimiter
    End If
    'Program Dates
    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
    If ilPass = 0 Then
        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
    Else
        slRecord = slRecord & sgCnCDelimiter
    End If
    'Event Names
    slRecord = slRecord & Trim$(tmEnf.sName) & sgCnCDelimiter
    'Program Event ID
    ''If slAvailComm = "" Then
    '    'slRecord = slRecord & Left$(slPrgComm, 5) & sgCnCDelimiter
    '    'slRecord = slRecord & Trim$(slPrgComm) & sgCnCDelimiter
    'If Trim$(tmDlf.sZone) <> "" Then
    '    For iZ = 1 To 4 Step 1
    '        If StrComp(Trim$(tmDlf.sZone), slPrgCommZ(iZ), 1) = 0 Then
    '            slRecord = slRecord & Trim$(slPrgComm(iZ)) & sgCnCDelimiter
    '            Exit For
    '        End If
    '    Next iZ
    'Else
    '    slRecord = slRecord & Trim$(slPrgComm(0)) & sgCnCDelimiter
    'End If
    ''Else
    ''    slRecord = slRecord & Left$(slAvailComm, 5) & sgCnCDelimiter
    ''End If
    slPrgC = ""
    gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
    llTime = gTimeToLong(slTime, False)
    For iZ = 0 To UBound(tmPrgInfo) - 1 Step 1
        If (llTime >= tmPrgInfo(iZ).lLocalSTime) And (llTime < tmPrgInfo(iZ).lLocalETime) And (Trim$(tmDlf.sZone) = Trim$(tmPrgInfo(iZ).sZone)) Then
            slPrgC = tmPrgInfo(iZ).sPrgID
            Exit For
        End If
    Next iZ
    slRecord = slRecord & Trim$(slPrgC) & sgCnCDelimiter
    'Prgram Category
    slRecord = slRecord & sgCnCDelimiter
    'Length in seconds
    slRecord = slRecord & Trim$(str$(CLng(gLengthToCurrency(tlLLC(ilLoop).sLength)))) & sgCnCDelimiter
    'Feed Time
    'slTime = Format$(gConvertTime(tlLLC(ilLoop).sStartTime), "hhmmss")
    gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
    slTime = Format$(gConvertTime(slTime), "hhmmss")
    slRecord = slRecord & slTime & sgCnCDelimiter
    'Feed Zone
    slRecord = slRecord & Trim$(tmDlf.sZone) & sgCnCDelimiter
    ''Time Zone
    'slStr = ""
    'If Trim$(tmDlf.sZone) <> "" Then
    '    For ilZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
    '        If StrComp(Trim$(tmDlf.sZone), Trim$(tgVpf(ilVpfIndex).sGZone(ilZone)), 1) = 0 Then
    '            If tgVpf(ilVpfIndex).sGFed(ilZone) = "Y" Then
    '                slStr = Trim$(tmDlf.sZone)
    '            End If
    '            Exit For
    '        End If
    '    Next ilZone
    'End If
    'slRecord = slRecord & slStr & sgCnCDelimiter
    'Break No
    slRecord = slRecord & Trim$(str$(ilBreakNo)) & sgCnCDelimiter
    'Avail Event ID
    'slRecord = slRecord & Left$(slAvailComm, 5) & sgCnCDelimiter
    slRecord = slRecord & Trim$(slAvailComm) & Trim$(tmDlf.sBus) & sgCnCDelimiter
    'Units
    slRecord = slRecord & Trim$(str$(tlLLC(ilLoop).iUnits)) & sgCnCDelimiter
    'Network basis
    If tgVpf(ilVpfIndex).sSSellOut = "U" Then
        slRecord = slRecord & "U" & sgCnCDelimiter
    Else
        slRecord = slRecord & "S" & sgCnCDelimiter
    End If
    'Network basis change date
    slRecord = slRecord & sgCnCDelimiter
    'Total Network Units/Length
    slRecord = slRecord & sgCnCDelimiter
    'airday
    If ilAdjDay = 0 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 1 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 2 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 3 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 4 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 5 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 6 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    'Relative time
    'slRecord = slRecord & Trim$(Str$(CLng(gLengthToCurrency(tlLLC(ilLoop).sLength)))) & sgCnCDelimiter
    slTime = Format$(gConvertTime(tlLLC(ilLoop).sStartTime), "hhmmss")
    slRecord = slRecord & slTime & sgCnCDelimiter
    'Network Length
    slRecord = slRecord & sgCnCDelimiter
    'Event Effective date
    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
    If ilPass = 0 Then
        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
    Else
        slRecord = slRecord & sgCnCDelimiter
    End If
    If tlLLC(ilLoop).iEtfCode = 2 Then
        If Left$(tmEnf.sName, 1) = "N" Then
            slRecord = slRecord & "N" & sgCnCDelimiter
        Else
            slRecord = slRecord & "L" & sgCnCDelimiter
        End If
    Else
        slRecord = slRecord & sgCnCDelimiter
    End If
    If ilVlfFd Then
        If tmVlf.iSellCode <> tmSVef.iCode Then
            tmVefSrchKey.iCode = tmVlf.iSellCode
            ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        End If
        'Selling Network code
        slRecord = slRecord & Trim$(tmSVef.sCodeStn) & sgCnCDelimiter
        'Selling Network ID
        slRecord = slRecord & Trim$(str$(tmSVef.iCode)) & sgCnCDelimiter
        'Selling Vehicle Name
        slRecord = slRecord & Trim$(tmSVef.sName) & sgCnCDelimiter
        'Selling Avail Time
        gUnpackTime tmVlf.iSellTime(0), tmVlf.iSellTime(1), "A", "1", slTime
        slTime = Format$(gConvertTime(slTime), "hhmmss")
        slRecord = slRecord & slTime & sgCnCDelimiter
        'Selling Day (0=Mon-Fri; 6=Sa; 7= Su)
        slRecord = slRecord & Trim$(str$(tmVlf.iSellDay)) & sgCnCDelimiter
        'Link Effective Start Date
        gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slStr
        slStr = Format$(gDateValue(slStr), "mmddyyyy")
        slRecord = slRecord & slStr & sgCnCDelimiter
        'Link Effective End Date
        gUnpackDate tmVlf.iTermDate(0), tmVlf.iTermDate(1), slStr
        If slStr <> "" Then
            slStr = Format$(gDateValue(slStr), "mmddyyyy")
        End If
        slRecord = slRecord & slStr & sgCnCDelimiter
        'Link Internal ID
        slRecord = slRecord & Trim$(str$(tmVlf.lCode)) & sgCnCDelimiter
        'Program Number
        llSCefCode = 0
        llSTime = gTimeToCurrency(slTime, False)
        For ilSLLC = LBound(tlSLLC) To UBound(tlSLLC) - 1 Step 1
            llSEvtTime = gTimeToCurrency(tlSLLC(ilSLLC).sStartTime, False)
            If llSEvtTime > llSTime Then
                Exit For
            ElseIf llSEvtTime = llSTime Then
                If (tlSLLC(ilSLLC).iEtfCode >= 2) And (tlSLLC(ilSLLC).iEtfCode <= 9) Then 'Avail
                    'If tlSLLC(ilSLLC).lCefCode > 0 Then
                    '    llSCefCode = tlSLLC(ilSLLC).lCefCode
                    'End If
                    If tlSLLC(ilSLLC).lEvtIDCefCode > 0 Then
                        llSCefCode = tlSLLC(ilSLLC).lEvtIDCefCode
                    End If
                    Exit For
                End If
            End If
            'If tlSLLC(ilSLLC).iEtfCode = 1 Then    'Program
            '    llSCefCode = tlSLLC(ilSLLC).lCefCode
            'End If
        Next ilSLLC
        If llSCefCode > 0 Then
            tmCefSrchKey.lCode = llSCefCode
            imCefRecLen = Len(tmCef)
            ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                'slStr = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                slStr = gStripChr0(tmCef.sComment)
                'slStr = Left$(slStr, 5)
            Else
                slStr = ""
            End If
        Else
            slStr = ""
        End If
        slRecord = slRecord & slStr
    Else
        'Selling Network code
        slRecord = slRecord & sgCnCDelimiter
        'Selling Network ID
        slRecord = slRecord & sgCnCDelimiter
        'Selling Vehicle Name
        slRecord = slRecord & sgCnCDelimiter
        'Selling Avail Time
        slRecord = slRecord & sgCnCDelimiter
        'Selling Day
        slRecord = slRecord & sgCnCDelimiter
        'Link Effective Start Date
        slRecord = slRecord & sgCnCDelimiter
        'Link Effective End Date
        slRecord = slRecord & sgCnCDelimiter
        'Link Internal ID
        slRecord = slRecord & sgCnCDelimiter
        'Selling Event ID
        slRecord = slRecord
    End If
    Print #hmTo, slRecord
End Sub

Private Sub mProcProg(ilDay As Integer, ilAirHour As Integer, ilLocalHour As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, slAdjDate As String, ilAdjDay As Integer, ilLoop As Integer, slDate As String, tlLLC() As LLC, ilBreakNo As Integer, ilPass As Integer, ilVpfIndex As Integer, slRecord As String)
    Dim llCefCode As Long
    Dim slPrgC As String
    Dim ilUpper As Integer
    Dim slTime As String
    Dim slMoDate As String
    Dim ilRet As Integer
    
    '6/1/16: Replaced GoSub
    'GoSub lProcAdjDate
    mProcAdjDate ilDay, ilAirHour, ilLocalHour, ilLogDate0, ilLogDate1, slAdjDate, ilAdjDay
    slRecord = smDateTime & sgCnCDelimiter
    'Vehicle Type
    slRecord = slRecord & tmVef.sType & sgCnCDelimiter
    'Event Type
    If tlLLC(ilLoop).iEtfCode = 1 Then
        slRecord = slRecord & "PR" & sgCnCDelimiter
    Else
        slRecord = slRecord & "CM" & sgCnCDelimiter
    End If
    If tlLLC(ilLoop).iEtfCode = 1 Then
        'llCefCode = tlLLC(ilLoop).lCefCode
        llCefCode = tlLLC(ilLoop).lEvtIDCefCode
        tmCefSrchKey.lCode = llCefCode
        imCefRecLen = Len(tmCef)
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
        '    slPrgComm = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))'tlCef.sComment
        '    slPrgComm = Trim$(slPrgComm) & Trim$(tmDlf.sBus)
        'Else
        '    slPrgComm = ""
            'slPrgC = Trim$(Left$(tmCef.sComment, tmCef.iStrLen)) 'tlCef.sComment
            slPrgC = gStripChr0(tmCef.sComment)
            slPrgC = Trim$(slPrgC) & Trim$(tmDlf.sBus)
        Else
            slPrgC = ""
        End If
        'If Trim$(tmDlf.sZone) <> "" Then
        '    For iZ = 1 To 4 Step 1
        '       If StrComp(Trim$(tmDlf.sZone), slPrgCommZ(iZ), 1) = 0 Then
        '           slPrgComm(iZ) = slPrgC
        '           Exit For
        '       End If
        '    Next iZ
        'Else
        '    slPrgComm(0) = slPrgC
        'End If
        ilUpper = UBound(tmPrgInfo)
        gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
        tmPrgInfo(ilUpper).lLocalSTime = gTimeToLong(slTime, False)
        tmPrgInfo(ilUpper).lLocalETime = tmPrgInfo(ilUpper).lLocalSTime + CLng(gLengthToCurrency(tlLLC(ilLoop).sLength))
        tmPrgInfo(ilUpper).sZone = tmDlf.sZone
        tmPrgInfo(ilUpper).sPrgID = slPrgC
        ReDim Preserve tmPrgInfo(0 To ilUpper + 1) As PRGINFO
    Else
        llCefCode = tlLLC(ilLoop).lEvtIDCefCode
        tmCefSrchKey.lCode = llCefCode
        imCefRecLen = Len(tmCef)
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
        '    slPrgComm = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))'tlCef.sComment
        '    slPrgComm = Trim$(slPrgComm) & Trim$(tmDlf.sBus)
        'Else
        '    slPrgComm = ""
            'slPrgC = Trim$(Left$(tmCef.sComment, tmCef.iStrLen)) 'tlCef.sComment
            slPrgC = gStripChr0(tmCef.sComment)
            slPrgC = Trim$(slPrgC) & Trim$(tmDlf.sBus)
        Else
            slPrgC = ""
        End If
    End If
    'Network code
    slRecord = slRecord & Trim$(tmVef.sCodeStn) & sgCnCDelimiter
    'Network ID
    slRecord = slRecord & Trim$(str$(tmVef.iCode)) & sgCnCDelimiter
    'Vehicle Name
    slRecord = slRecord & Trim$(tmVef.sName) & sgCnCDelimiter
    'Effective date
    slMoDate = gObtainPrevMonday(slDate)
    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
    If ilPass = 0 Then
        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
    Else
        slRecord = slRecord & sgCnCDelimiter
    End If
    'Program Dates
    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
    If ilPass = 0 Then
        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
    Else
        slRecord = slRecord & sgCnCDelimiter
    End If
    'Event Names
    slRecord = slRecord & Trim$(tmEnf.sName) & sgCnCDelimiter
    'Program Event ID
    'slRecord = slRecord & Left$(slPrgComm, 5) & sgCnCDelimiter
    slRecord = slRecord & Trim$(slPrgC) & sgCnCDelimiter
    'Program Category for RADAR
    slRecord = slRecord & sgCnCDelimiter
    'Length in seconds
    slRecord = slRecord & Trim$(str$(CLng(gLengthToCurrency(tlLLC(ilLoop).sLength)))) & sgCnCDelimiter
    'Feed Time
    'slTime = Format$(gConvertTime(tlLLC(ilLoop).sStartTime), "hhmmss")
    gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
    slTime = Format$(gConvertTime(slTime), "hhmmss")
    slRecord = slRecord & slTime & sgCnCDelimiter
    'Feed Zone
    slRecord = slRecord & Trim$(tmDlf.sZone) & sgCnCDelimiter
    ''Time Zone
    'slStr = ""
    'If Trim$(tmDlf.sZone) <> "" Then
    '    For ilZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
    '        If StrComp(Trim$(tmDlf.sZone), Trim$(tgVpf(ilVpfIndex).sGZone(ilZone)), 1) = 0 Then
    '            If tgVpf(ilVpfIndex).sGFed(ilZone) = "Y" Then
    '                slStr = Trim$(tmDlf.sZone)
    '            End If
    '            Exit For
    '        End If
    '    Next ilZone
    'End If
    'slRecord = slRecord & slStr & sgCnCDelimiter
    'Break No
    slRecord = slRecord & Trim$(str$(ilBreakNo)) & sgCnCDelimiter
    'Program Event ID
    'slRecord = slRecord & Left$(slPrgComm, 5) & sgCnCDelimiter
    slRecord = slRecord & Trim$(slPrgC) & sgCnCDelimiter
    'Units
    slRecord = slRecord & sgCnCDelimiter
    'Network basis
    If tgVpf(ilVpfIndex).sSSellOut = "U" Then
        slRecord = slRecord & "U" & sgCnCDelimiter
    Else
        slRecord = slRecord & "S" & sgCnCDelimiter
    End If
    'Network basis change date
    slRecord = slRecord & sgCnCDelimiter
    'Total Network Units/Length
    slRecord = slRecord & sgCnCDelimiter
    'airday
    If ilAdjDay = 0 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 1 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 2 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 3 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 4 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 5 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    If ilAdjDay = 6 Then
        slRecord = slRecord & "Y" & sgCnCDelimiter
    Else
        slRecord = slRecord & "N" & sgCnCDelimiter
    End If
    'Relative time
    'slRecord = slRecord & Trim$(Str$(CLng(gLengthToCurrency(tlLLC(ilLoop).sLength)))) & sgCnCDelimiter
    slTime = Format$(gConvertTime(tlLLC(ilLoop).sStartTime), "hhmmss")
    slRecord = slRecord & slTime & sgCnCDelimiter
    'Network Length
    slRecord = slRecord & sgCnCDelimiter
    'Event Effective date
    slRecord = slRecord & Format$(gDateValue(slMoDate), "mmddyyyy") & sgCnCDelimiter
    If ilPass = 0 Then
        slRecord = slRecord & Format$(gDateValue(slMoDate) + 6, "mmddyyyy") & sgCnCDelimiter
    Else
        slRecord = slRecord & sgCnCDelimiter
    End If
    'Avail Type
    slRecord = slRecord & sgCnCDelimiter
    'Selling Network code
    slRecord = slRecord & sgCnCDelimiter
    'Selling Network ID
    slRecord = slRecord & sgCnCDelimiter
    'Selling Vehicle Name
    slRecord = slRecord & sgCnCDelimiter
    'Selling Avail Time
    slRecord = slRecord & sgCnCDelimiter
    'Selling Day
    slRecord = slRecord & sgCnCDelimiter
    'Link Effective Start Date
    slRecord = slRecord & sgCnCDelimiter
    'Link Effective End Date
    slRecord = slRecord & sgCnCDelimiter
    'Link Internal ID
    slRecord = slRecord & sgCnCDelimiter
    'Selling Event ID
    slRecord = slRecord
    Print #hmTo, slRecord
End Sub
