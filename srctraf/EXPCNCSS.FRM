VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExpCnCSS 
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
      Left            =   6000
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox plcTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   990
      ScaleHeight     =   345
      ScaleWidth      =   4275
      TabIndex        =   21
      Top             =   345
      Width           =   4305
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   60
         TabIndex        =   22
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
      TabIndex        =   20
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
      Left            =   2115
      Picture         =   "EXPCNCSS.frx":0000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
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
      Width           =   930
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1185
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   15
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
         Picture         =   "EXPCNCSS.frx":00FA
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
      ScaleWidth      =   3600
      TabIndex        =   0
      Top             =   45
      Width           =   3600
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4185
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4455
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4605
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   1515
      MultiSelect     =   2  'Extended
      TabIndex        =   6
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
      Left            =   2250
      TabIndex        =   7
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
      TabIndex        =   8
      Top             =   4395
      Width           =   1050
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      Caption         =   "To File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   105
      TabIndex        =   23
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
      TabIndex        =   19
      Top             =   4095
      Width           =   6330
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Export Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   825
      Width           =   1065
   End
   Begin VB.Label lacErrors 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5055
      TabIndex        =   14
      Top             =   1800
      Width           =   1725
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
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
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3255
      TabIndex        =   10
      Top             =   1845
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "ExpCnCSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EXPCNCSS.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpCncSS.Frm
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
'Rotation
Dim hmCrf As Integer
Dim tmCrf As CRF
Dim imCrfRecLen As Integer  'CRF record length
'Avail name
Dim hmAnf As Integer
Dim tmAnf As ANF
Dim imAnfRecLen As Integer  'ANF record length
'Book name
Dim hmDnf As Integer
Dim tmDnf As DNF
Dim tmDnfSrchKey As INTKEY0 'ANF key record image
Dim imDnfRecLen As Integer  'ANF record length
'Contract record information
Dim hmCHF As Integer        'Contract header file handle
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
'Contract Line record information
Dim hmClf As Integer        'Contract Line file handle
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image
Dim tmClfSrchKey As CLFKEY0            'CLF record image
'Contract Line Flight record information
Dim hmCff As Integer        'Contract Line file handle
Dim imCffRecLen As Integer  'CLF record length
Dim tmCff As CFF            'CLF record image
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
Dim tmBofSrchKey As LONGKEY0
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
Dim tmSmf As SMF
Dim hmSmf As Integer
Dim imSmfRecLen As Integer
Dim tmSmfSrchKey2 As LONGKEY0
Dim tmCef As CEF
Dim hmCef As Integer
Dim imCefRecLen As Integer
Dim tmCefSrchKey As LONGKEY0
Dim tmCyf As CYF
Dim hmCyf As Integer
Dim imCyfRecLen As Integer
'Daypart
Dim hmRdf As Integer
Dim tmRdf As RDF
Dim tmRdfSrchKey As INTKEY0 'RDF key record image
Dim imRdfRecLen As Integer  'RDF record length
Dim tmOdf As ODF
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
    ilDays = 7  'Val(edcNoDays.Text)
    'If ilDays = 0 Then
    '    Beep
    '    edcNoDays.SetFocus
    '    Exit Sub
    'End If
    slEndDate = Format$(gDateValue(slStartDate) + ilDays - 1, "m/d/yy")
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
    imExporting = True
    slMissingCopyNames = ""
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            imCopyMissing = False
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
    'Print #hmMsg, "** Completed Export CNC Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** Completed Export CNC Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
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
    
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    ilRet = btrClose(hmCyf)
    btrDestroy hmCyf
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
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
    ilRet = btrClose(hmDnf)
    btrDestroy hmDnf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
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
    
    Set ExpCnCSS = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
'    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
'    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
'    Traffic!cdcSetup.Action = 6
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
'Private Function mExpRecs(slType As String, slCP As String, ilCallCode As Integer, slSDate As String, slEDate As String, slStartTime As String, slEndTime As String, ilEvtType() As Integer) As Integer
Private Function mExpRecs(ilSSFType As Integer, sLCP As String, ilCallCode As Integer, slSDate As String, slEDate As String, slStartTime As String, slEndTime As String, ilEvtType() As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slType                                                                                *
'******************************************************************************************


'
'   iRet = mExpRecs()
'   Where:
'       iRet (O)- True if record exported,
'                 False if error
'
    Dim ilLCFType As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer    'Return status
    Dim slTime As String
    Dim slRecord As String
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilSsfDate0 As Integer
    Dim ilSsfDate1 As Integer
    Dim ilEvt As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim ilSpot As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilTerminated As Integer
    Dim ilStartTime0 As Integer
    Dim ilStartTime1 As Integer
    Dim ilEndTime0 As Integer
    Dim ilEndTime1 As Integer
    Dim ilSeqNo As Integer
    Dim slLength As String
    Dim ilVpfIndex As Integer
    Dim ilVehCode As Integer
    Dim ilVeh As Integer
    Dim ilDlfDate0 As Integer
    Dim ilDlfDate1 As Integer
    Dim ilDlfFound As Integer
    Dim ilVlfDate0 As Integer
    Dim ilVlfDate1 As Integer
    Dim ilSIndex As Integer
    Dim slSsfDate As String
    Dim ilBreakNo As Integer    'Reset to zero for each program
    Dim ilPositionNo As Integer 'Reset to zero for each avail
    Dim llADate As Long
    Dim llAWDate As Long
    'Spot summary
    Dim tlSsfSrchKey As SSFKEY0 'SSF key record image
    Dim ilSsfRecLen As Integer  'SSF record length
    Dim llEvtTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilWithinTime As Integer
    Dim ilEvtFdIndex As Integer
    Dim llCefCode As Long
    Dim llEvtIDCefCode As Long
    Dim ilAirHour As Integer
    Dim ilLocalHour As Integer
    Dim slAdjDate As String
    Dim ilType As Integer   '1=Prog; 2=Comment; 3=Avail; 4=Spot
    'Spot detail record information
    Dim ilEvtRet As Integer
    'Copy rotation record information
    Dim ilCrfVefCode As Integer
    Dim ilPkgVefCode As Integer
    Dim ilLnVefCode As Integer
    Dim slLive As String
    Dim ilRdfCode As Integer
    'Copy inventory record information
    'Dim tlSifSrchKey As LONGKEY0 'SIF key record image
    'Dim ilSifRecLen As Integer  'SIF record length
    'Dim tlSif As SIF            'SIF record image
    'Dim llSifCode As Long
    'Advertiser record information
    Dim llCifCode As Long
    Dim slZone As String
    'Dim ilWkNo As Integer   'Week number from 1/5/70
    Dim ilOther As Integer
    Dim ilZone As Integer
    Dim ilNoZones As Integer
    Dim llCrfCsfCode As Long
    Dim ilRdf As Integer
    Dim ilRdfDay As Integer
    Dim slPrice As String
    Dim ilPos As Integer
    Dim slAvailComm As String
    Dim ilBlackoutDefined As Integer
    Dim llBChfCode As Long
    Dim llBCifCode As Long
    Dim llBSifCode As Long
    Dim llBBofCode As Long
    Dim slXMid As String

    ReDim tlLLC(0 To 0) As LLC  'Image
    ReDim ilVehicle(0 To 1) As Integer
    Dim ilVff As Integer
    Dim blBypassZeroUnits As Boolean

    'If ilSSFType <> 0 Then
    '    slType = "A"
    'Else
    '    slType = "O"
    'End If
    ilLCFType = 0
    ilVehicle(0) = ilCallCode
    llStartTime = CLng(gTimeToCurrency(slStartTime, False))
    llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1

    tmVef.iCode = 0
    llSDate = gDateValue(slSDate)
    llEDate = gDateValue(slEDate)
    For llDate = llSDate To llEDate Step 1
        If imTerminate Then
            mExpRecs = False
            Exit Function
        End If
        ilWithinTime = False
        slDate = Format$(llDate, "m/d/yy")
        ilDay = gWeekDayStr(slDate)
        gPackDate slDate, ilLogDate0, ilLogDate1
        For ilVeh = LBound(ilVehicle) To UBound(ilVehicle) - 1 Step 1
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
            
            blBypassZeroUnits = False
            If (tmVef.sType = "A") Then
                ilVff = gBinarySearchVff(tmVef.iCode)
                If ilVff <> -1 Then
                    If tgVff(ilVff).sHonorZeroUnits = "Y" Then
                        blBypassZeroUnits = True
                    End If
                End If
            End If
            
            'DoEvents
            igDOE = igDOE + 1
            If igDOE > 100 Then
                DoEvents
                igDOE = 0
            End If
            'gObtainVlf hlVlf, ilVehCode, llDate, tlVlf0(), tlVlf5(), tlVlf6()
            ilDay = gWeekDayStr(slDate)
            gPackDate slDate, ilLogDate0, ilLogDate1
            ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilSsfDate0 = ilLogDate0
            ilSsfDate1 = ilLogDate1
            tlSsfSrchKey.iType = ilSSFType
            tlSsfSrchKey.iVefCode = ilVehCode
            tlSsfSrchKey.iDate(0) = ilSsfDate0
            tlSsfSrchKey.iDate(1) = ilSsfDate1
            tlSsfSrchKey.iStartTime(0) = 0
            tlSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetEqual(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilSSFType) Or (tmSsf.iVefCode <> ilVehCode) Or (tmSsf.iDate(0) <> ilSsfDate0) Or (tmSsf.iDate(1) <> ilSsfDate1) Then
                'If airing- then use first Ssf prior to date defined
                If tmVef.sType = "A" Then
                    ilSsfDate0 = 0
                    ilSsfDate1 = 0
                    ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    tlSsfSrchKey.iType = ilSSFType
                    tlSsfSrchKey.iVefCode = ilVehCode
                    tlSsfSrchKey.iDate(0) = ilLogDate0
                    tlSsfSrchKey.iDate(1) = ilLogDate1
                    tlSsfSrchKey.iStartTime(0) = 0
                    tlSsfSrchKey.iStartTime(1) = 6144   '24*256
                    ilRet = gSSFGetLessOrEqual(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = ilVehCode)
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
            'DoEvents
            igDOE = igDOE + 1
            If igDOE > 100 Then
                DoEvents
                igDOE = 0
            End If
            If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = ilVehCode) Then
                gUnpackDate ilSsfDate0, ilSsfDate1, slSsfDate
                If (ilEvtType(0) = True) Or (ilEvtType(10) = True) Or (ilEvtType(11) = True) Or (ilEvtType(12) = True) Or (ilEvtType(13) = True) Or (ilEvtType(14) = True) Then
                    ReDim tlLLC(0 To 0) As LLC  'Image
                    ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, slSsfDate, slStartTime, slEndTime, ilEvtType(), tlLLC())
                Else
                    If ilEvtType(1) = True Then
                        ReDim tlLLC(0 To 0) As LLC  'Image
                        ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, slSsfDate, slStartTime, slEndTime, ilEvtType(), tlLLC())
                    End If
                End If
                ilBreakNo = 0
                ilPositionNo = 0
                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = ilVehCode) And (tmSsf.iDate(0) = ilSsfDate0) And (tmSsf.iDate(1) = ilSsfDate1)
                    'Loop thru Ssf and move records to tmOdf
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
                        llCefCode = 0
                        llEvtIDCefCode = 0
                        If tmProg.iRecType = 1 Then
                            ilBreakNo = 0
                            ilPositionNo = 0
                        ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then
                            ilBreakNo = ilBreakNo + 1
                            ilPositionNo = 0
                        End If
                        ilEvtFdIndex = -1
                        If ilWithinTime Then
                            If tmProg.iRecType = 1 Then    'Program
                                For ilLoop = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                                    'Match start time and length
                                    If tlLLC(ilLoop).iEtfCode = 1 Then
                                        gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                        If (ilStartTime0 = tmProg.iStartTime(0)) And (ilStartTime1 = tmProg.iStartTime(1)) Then
                                            gAddTimeLength tlLLC(ilLoop).sStartTime, tlLLC(ilLoop).sLength, "A", "1", slTime, slXMid
                                            gPackTime slTime, ilEndTime0, ilEndTime1
                                            If (ilEndTime0 = tmProg.iEndTime(0)) And (ilEndTime1 = tmProg.iEndTime(1)) Then
                                                ilEvtFdIndex = ilLoop
                                                llCefCode = tlLLC(ilLoop).lCefCode
                                                llEvtIDCefCode = tlLLC(ilLoop).lEvtIDCefCode
                                                ilType = 1
                                                    'Only spot are sent to delivery
                                                If (tmVef.sType = "A") Or (ilDlfFound) Then
                                                    'Obtain delivery entry to see is prog is sent
                                                    tmDlfSrchKey.iVefCode = ilVehCode
                                                    tmDlfSrchKey.sAirDay = slDay
                                                    tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                    tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                    tmDlfSrchKey.iAirTime(0) = tmProg.iStartTime(0)
                                                    tmDlfSrchKey.iAirTime(1) = tmProg.iStartTime(1)
                                                    ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                    Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmProg.iStartTime(0)) And (tmDlf.iAirTime(1) = tmProg.iStartTime(1))
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
                                                                        '6/2/16: Replaced GoSub
                                                                        'GoSub lProcProg
                                                                        mProcProg
                                                                        DoEvents
                                                                    End If
                                                                'End If
                                                            End If
                                                        End If
                                                        ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    Loop
                                                Else
                                                    tmDlf.iVefCode = ilVehCode
                                                    tmDlf.iLocalTime(0) = tmProg.iStartTime(0)
                                                    tmDlf.iLocalTime(1) = tmProg.iStartTime(1)
                                                    tmDlf.iFeedTime(0) = tmProg.iStartTime(0)
                                                    tmDlf.iFeedTime(1) = tmProg.iStartTime(1)
                                                    tmDlf.sZone = ""
                                                    tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
                                                    tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
                                                    tmDlf.sProgCode = ""
                                                    tmDlf.iMnfFeed = 0
                                                    tmDlf.sBus = ""
                                                    tmDlf.sSchedule = ""
                                                    tmDlf.iMnfSubFeed = 0
                                                    '6/2/16: Replaced GoSub
                                                    'GoSub lProcProg
                                                    mProcProg
                                                    'DoEvents
                                                    igDOE = igDOE + 1
                                                    If igDOE > 100 Then
                                                        DoEvents
                                                        igDOE = 0
                                                    End If
                                                End If
                                                'Determine if any events that are
                                                tlLLC(ilLoop).iEtfCode = -1 'Remove event
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next ilLoop
                            ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then 'Avail
                               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                For ilLoop = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                                    'Match start time and length
                                    If (tlLLC(ilLoop).iEtfCode >= 2) And (tlLLC(ilLoop).iEtfCode <= 9) Then
                                        gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                        If (ilStartTime0 = tmAvail.iTime(0)) And (ilStartTime1 = tmAvail.iTime(1)) Then
                                            ilEvtFdIndex = ilLoop
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
                                                'Only spot are sent to delivery
                                            If (tmVef.sType = "A") Or (ilDlfFound) Then
                                                'Don't include avails

                                                'Obtain delivery entry to see is avail is sent
                                                'tmDlfSrchKey.iVefCode = ilVehCode
                                                'tmDlfSrchKey.sAirDay = slDay
                                                'tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                'tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                'tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                'tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                'ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                                'Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                '        ilTerminated = False
                                                '        If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                                '            ilTerminated = True
                                                '        Else
                                                '           If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                '               If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                '                   ilTerminated = True
                                                '               End If
                                                '           End If
                                                '        End If
                                                '        If Not ilTerminated Then
                                                '        If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                                '            'If slFor = "C" Then
                                                '                If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0)  Then
                                                '                    tmDlf.iMnfFeed = 0
                                                '                    GoSub lProcAvail
                                                '                End If
                                                '            'End If
                                                '        End If
                                                '    End If
                                                '   ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE)
                                                'Loop
                                            Else
                                                tmDlf.iVefCode = ilVehCode
                                                tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                tmDlf.sZone = ""
                                                tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
                                                tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
                                                tmDlf.sProgCode = ""
                                                tmDlf.iMnfFeed = 0
                                                tmDlf.sBus = ""
                                                tmDlf.sSchedule = ""
                                                tmDlf.iMnfSubFeed = 0
                                                'GoSub lProcAvail
                                            End If
                                            'Loop on spots, then add conflicting spots
                                            If (tmVef.sType = "A") Then
                                                ilType = 4
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
                                                        If (tmCTSsf.iType <> ilSSFType) Or (tmCTSsf.iVefCode <> tmVlf.iSellCode) Or (tmCTSsf.iDate(0) <> ilLogDate0) Or (tmCTSsf.iDate(1) <> ilLogDate1) Then
                                                            ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                            tlSsfSrchKey.iType = ilSSFType
                                                            tlSsfSrchKey.iVefCode = tmVlf.iSellCode
                                                            tlSsfSrchKey.iDate(0) = ilLogDate0
                                                            tlSsfSrchKey.iDate(1) = ilLogDate1
                                                            tlSsfSrchKey.iStartTime(0) = 0
                                                            tlSsfSrchKey.iStartTime(1) = 0
                                                            ilRet = gSSFGetEqual(hmCTSsf, tmCTSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                        End If
                                                        Do While (ilRet = BTRV_ERR_NONE) And (tmCTSsf.iType = ilSSFType) And (tmCTSsf.iVefCode = tmVlf.iSellCode) And (tmCTSsf.iDate(0) = ilLogDate0) And (tmCTSsf.iDate(1) = ilLogDate1)
                                                            For ilSIndex = 1 To tmCTSsf.iCount Step 1
                                                                tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSIndex)
                                                                If ((tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9)) Then
                                                                    If (tmAvailTest.iTime(0) = tmVlf.iSellTime(0)) And (tmAvailTest.iTime(1) = tmVlf.iSellTime(1)) Then
                                                                        For ilSpot = 1 To tmAvailTest.iNoSpotsThis Step 1
                                                                            ilType = 4
                                                                           LSet tmSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilSpot + ilSIndex)
                                                                            tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                            If ilRet = BTRV_ERR_NONE Then
                                                                                tmChfSrchKey.lCode = tmSdf.lChfCode
                                                                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                                                If ilRet = BTRV_ERR_NONE Then
                                                                                    ilPositionNo = ilPositionNo + 1
                                                                                    slZone = "EST"  'Use EST as standard, if not found, use OTH
                                                                                    llCifCode = mObtainCifCodeSS(tmSdf, slZone, hmTzf, ilOther)
                                                                                    ilRet = gGetCrfVefCode(hmClf, tmSdf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                                                                                    llCrfCsfCode = mObtainCrfCsfCodeSS(tmSdf, slZone, hmCrf, hmTzf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, tmCrf)
                                                                                    'Remove comment
                                                                                    llCrfCsfCode = 0
                                                                                    tmDlfSrchKey.iVefCode = ilVehCode
                                                                                    tmDlfSrchKey.sAirDay = slDay
                                                                                    tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                                    tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                                    tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                                    tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                                    ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                    Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
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
                                                                                                If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                                                                                                    tmDlf.iMnfFeed = 0
                                                                                                    '6/2/16:Replaced GoSub
                                                                                                    'GoSub lProcSpot
                                                                                                    mProcSpot ilLogDate0, ilLogDate1, slLength, ilVehCode, ilType, llBChfCode, llBCifCode, llBSifCode, llBBofCode, ilBlackoutDefined, llCifCode, ilBreakNo, ilPositionNo, llCrfCsfCode, ilAirHour, ilLocalHour, slAdjDate, ilSeqNo, slAvailComm, slRecord
                                                                                                    DoEvents
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                        ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                                    Loop
                                                                                End If
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
                                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                                    ilType = 4
                                                    ilEvt = ilEvt + 1
                                                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                                    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    If ilRet = BTRV_ERR_NONE Then
                                                        tmChfSrchKey.lCode = tmSdf.lChfCode
                                                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            ilPositionNo = ilPositionNo + 1
                                                            slZone = "EST"  'Use EST as standard, if not found, use OTH
                                                            If ilDlfFound Then
                                                                llCifCode = mObtainCifCodeSS(tmSdf, slZone, hmTzf, ilOther)
                                                                ilRet = gGetCrfVefCode(hmClf, tmSdf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                                                                llCrfCsfCode = mObtainCrfCsfCodeSS(tmSdf, slZone, hmCrf, hmTzf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, tmCrf)
                                                                'Remove Comment
                                                                llCrfCsfCode = 0
                                                                'Obtain delivery entry to see is avail is sent
                                                                tmDlfSrchKey.iVefCode = ilVehCode
                                                                tmDlfSrchKey.sAirDay = slDay
                                                                tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
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
                                                                            If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                                                                                tmDlf.iMnfFeed = 0
                                                                                '6/2/16:Replaced GoSub
                                                                                'GoSub lProcSpot
                                                                                mProcSpot ilLogDate0, ilLogDate1, slLength, ilVehCode, ilType, llBChfCode, llBCifCode, llBSifCode, llBBofCode, ilBlackoutDefined, llCifCode, ilBreakNo, ilPositionNo, llCrfCsfCode, ilAirHour, ilLocalHour, slAdjDate, ilSeqNo, slAvailComm, slRecord
                                                                                DoEvents
                                                                            End If
                                                                        End If
                                                                    End If
                                                                    ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                Loop
                                                            Else
                                                                'Assign copy for all zones
                                                                ilNoZones = 0
                                                                For ilZone = 1 To 5 Step 1
                                                                    Select Case ilZone
                                                                        Case 1
                                                                            slZone = "EST"
                                                                        Case 2
                                                                            slZone = "MST"
                                                                        Case 3
                                                                            slZone = "CST"
                                                                        Case 4
                                                                            slZone = "PST"
                                                                        Case 5
                                                                            slZone = "Oth"
                                                                    End Select
                                                                    llCifCode = mObtainCifCodeSS(tmSdf, slZone, hmTzf, ilOther)
                                                                    ilRet = gGetCrfVefCode(hmClf, tmSdf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                                                                    llCrfCsfCode = mObtainCrfCsfCodeSS(tmSdf, slZone, hmCrf, hmTzf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, tmCrf)
                                                                    tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                    tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                    tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                    tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                    If (ilOther) And (ilNoZones = 0) Then
                                                                        tmDlf.sZone = ""
                                                                    Else
                                                                        tmDlf.sZone = slZone
                                                                        If ilNoZones = 0 Then
                                                                            ilType = 4
                                                                            ilNoZones = 1
                                                                        Else
                                                                            ilType = 5
                                                                            ilNoZones = ilNoZones + 1
                                                                        End If
                                                                    End If
                                                                    tmDlf.iEtfCode = 0
                                                                    tmDlf.iEnfCode = 0
                                                                    'tmDlf.sProgCode = ""
                                                                    '1/10/98 AMFM wants comment in all spots in break
                                                                    'If ilPositionNo = 1 Then        '1st position of spot has the avail spot code
                                                                        tmDlf.sProgCode = slAvailComm
                                                                    'Else
                                                                    '    tmDlf.sProgCode = ""
                                                                    'End If
                                                                    tmDlf.iMnfFeed = 0
                                                                    tmDlf.sBus = ""
                                                                    tmDlf.sSchedule = ""
                                                                    tmDlf.iMnfSubFeed = 0
                                                                    '6/2/16:Replaced GoSub
                                                                    'GoSub lProcSpot
                                                                    mProcSpot ilLogDate0, ilLogDate1, slLength, ilVehCode, ilType, llBChfCode, llBCifCode, llBSifCode, llBBofCode, ilBlackoutDefined, llCifCode, ilBreakNo, ilPositionNo, llCrfCsfCode, ilAirHour, ilLocalHour, slAdjDate, ilSeqNo, slAvailComm, slRecord
                                                                    igDOE = igDOE + 1
                                                                    If igDOE > 100 Then
                                                                        DoEvents
                                                                        igDOE = 0
                                                                    End If
                
                                                                    If (ilOther) Or (ilNoZones = 4) Then
                                                                        Exit For
                                                                    End If
                                                                Next ilZone
                                                            End If
                                                        End If
                                                    End If
                                                Next ilSpot
                                            End If
                                            tlLLC(ilLoop).iEtfCode = -1 'Remove event
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                            End If
                        End If
                        'Any other events to be sent out
                        If ilEvtFdIndex >= 0 Then
                            'Output all event until Program or Avail found
                            For ilLoop = ilEvtFdIndex + 1 To UBound(tlLLC) - 1 Step 1
                                'Match start time and length
                                If (tlLLC(ilLoop).iEtfCode > 9) Then
                                    'Handle like program
                                    ilType = 2
                                    llEvtTime = CLng(gTimeToCurrency(tlLLC(ilLoop).sStartTime, False))
                                    If llEvtTime > llEndTime Then
                                        ilWithinTime = False
                                        Exit Do
                                    End If
                                    If llEvtTime >= llStartTime Then
                                        ilWithinTime = True
                                    End If
                                    If ilWithinTime Then
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
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmProg.iStartTime(0)) And (tmDlf.iAirTime(1) = tmProg.iStartTime(1))
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
                                                                '6/2/16: Replaced GoSub
                                                                'GoSub lProcProg
                                                                mProcProg
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
                                            tmDlf.sZone = ""
                                            tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
                                            tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
                                            tmDlf.sProgCode = ""
                                            tmDlf.iMnfFeed = 0
                                            tmDlf.sBus = ""
                                            tmDlf.sSchedule = ""
                                            tmDlf.iMnfSubFeed = 0
                                            '6/2/16: Replaced GoSub
                                            'GoSub lProcProg
                                            mProcProg
                                            DoEvents
                                        End If
                                    End If
                                    tlLLC(ilLoop).iEtfCode = -1 'Remove event
                                Else
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                        ilEvt = ilEvt + 1
                    Loop
                    ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetNext(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
        Next ilVeh
    Next llDate
    mExpRecs = True
    Exit Function
'lProcProg:
'    Return
'lProcSpot:
'    tmOdf.iUrfCode = tgUrf(0).iCode
'    If tmVef.iVefCode > 0 Then  'Log vehicle defined
'        tmOdf.iVefCode = tmVef.iVefCode
'    Else
'        tmOdf.iVefCode = ilVehCode
'    End If
'    tmOdf.iAirDate(0) = ilLogDate0
'    tmOdf.iAirDate(1) = ilLogDate1
'    tmOdf.iAirTime(0) = tmAvail.iTime(0)
'    tmOdf.iAirTime(1) = tmAvail.iTime(1)
'    tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
'    tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
'    tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
'    tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
'    tmOdf.sZone = tmDlf.sZone
'    tmOdf.iEtfCode = 0
'    tmOdf.iEnfCode = tmDlf.iEnfCode 'Required for Commercial schedule
'    tmOdf.sProgCode = tmDlf.sProgCode
'    tmOdf.iMnfFeed = tmDlf.iMnfFeed
'    'tmOdf.sUnused1 = "" 'tmDlf.sBus
'    tmOdf.iWkNo = 0    'tmDlf.sSchedule
'    tmOdf.ianfCode = tmAvail.ianfCode
'    tmOdf.iUnits = 0
'    slLength = Trim$(str$(tmSdf.iLen)) & "s"
'    gPackLength slLength, tmOdf.iLen(0), tmOdf.iLen(1)
'    'Test if Blackout exist
'    If gSDFBlackoutTest(hmRsf, hmBof, tmSdf.lCode, ilVehCode, llBChfCode, llBCifCode, llBSifCode, llBBofCode) Then
'        ilBlackoutDefined = True
'        tmChfSrchKey.lCode = llBChfCode
'        ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        If ilRet = BTRV_ERR_NONE Then
'            tmAdfSrchKey.iCode = tmChf.iAdfCode
'            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        End If
'        tmOdf.iAdfCode = tmChf.iAdfCode
'        tmOdf.lCifCode = llBCifCode
'        tmBofSrchKey.lCode = llBBofCode
'        ilRet = btrGetEqual(hmBof, tmBof, imBofRecLen, tmBofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    Else
'        ilBlackoutDefined = False
'        tmOdf.iAdfCode = tmChf.iAdfCode
'        'Test tmSdf.sPtType
'        tmOdf.lCifCode = llCifCode
'    End If
'    If Trim$(tmChf.sProduct) <> "" Then
'        tmOdf.sProduct = tmChf.sProduct
'    Else
'        tmOdf.sProduct = "" '"???? Name ????"
'    End If
'    tmOdf.iMnfSubFeed = tmDlf.iMnfSubFeed
'    tmOdf.lCntrNo = tmChf.lCntrNo
'    tmOdf.iBreakNo = ilBreakNo
'    tmOdf.iPositionNo = ilPositionNo
'    tmOdf.iType = ilType
'    If ilBlackoutDefined Then
'        tmOdf.lCefCode = 0
'        If tgSpf.sUseProdSptScr = "P" Then
'            If llBSifCode <= 0 Then
'                tmSifSrchKey.lCode = tmChf.lSifCode
'            Else
'                tmSifSrchKey.lCode = llBSifCode
'            End If
'            ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'            If ilRet = BTRV_ERR_NONE Then
'                tmOdf.sShortTitle = Trim$(tmSif.sName)
'            Else
'                tmOdf.sShortTitle = ""
'            End If
'        Else
'            tmOdf.sShortTitle = Trim$(tmAdf.sAbbr) & "," & Trim$(tmChf.sProduct)
'        End If
'    Else
'        tmOdf.lCefCode = llCrfCsfCode   '0
'        If tmAdf.iCode <> tmChf.iAdfCode Then
'            tmAdfSrchKey.iCode = tmChf.iAdfCode
'            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        End If
'        tmOdf.sShortTitle = gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)
'    End If
'    '6/2/16: Replaced GoSub
'    'GoSub lProcAdjDate  'Adjust dates prior to adjusting seq numbers
'    mProcAdjDate ilAirHour, ilLocalHour, slAdjDate
'    '6/2/16: Replaced GoSub
'    'GoSub lProcSeqNo
'    mProcSeqNo ilSeqNo
'    tmOdf.iSeqNo = ilSeqNo + 1
'    'ilRet = btrInsert(hmOdf, tmOdf, imOdfRecLen, INDEXKEY0)
'    slRecord = Trim$(str$(tmSdf.lCode)) & sgCnCDelimiter  '"SPOT,"
'    slRecord = slRecord & Trim$(str$(tmChf.lCntrNo)) & sgCnCDelimiter
'    If ilBlackoutDefined Then
'        slRecord = slRecord & "0" & sgCnCDelimiter
'    Else
'        slRecord = slRecord & Trim$(str$(tmSdf.iLineNo)) & sgCnCDelimiter
'    End If
'    slRecord = slRecord & Trim$(str$(tmSdf.iLen)) & sgCnCDelimiter
'    slRecord = slRecord & Trim$(str$(tmChf.iAdfCode)) & sgCnCDelimiter
'    slRecord = slRecord & sgCnCNameBracket & Trim$(tmAdf.sName) & sgCnCNameBracket & sgCnCDelimiter
'    slRecord = slRecord & sgCnCNameBracket & Trim$(tmChf.sProduct) & sgCnCNameBracket & sgCnCDelimiter
'    'Selling Vehicle
'    If ilBlackoutDefined Then
'        tmVefSrchKey.iCode = tmSdf.iVefCode
'        ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'        slRecord = slRecord & Trim$(tmSVef.sName) & sgCnCDelimiter
'    Else
'        tmClfSrchKey.lChfCode = tmChf.lCode
'        tmClfSrchKey.iLine = tmSdf.iLineNo
'        tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
'        tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
'        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) And (tmClf.iLine = tmSdf.iLineNo) Then
'            tmVefSrchKey.iCode = tmClf.iVefCode
'            ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'            slRecord = slRecord & Trim$(tmSVef.sName) & sgCnCDelimiter
'        Else
'            slRecord = slRecord & "Sell Veh Missing" & sgCnCDelimiter
'        End If
'    End If
'    'Airing Vehicle
'    slRecord = slRecord & Trim$(tmVef.sName) & sgCnCDelimiter
'    'Post Log Vehicle
'    If ilBlackoutDefined Then
'        slRecord = slRecord & Trim$(tmSVef.sName) & sgCnCDelimiter
'        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
'        llAWDate = llADate
'        slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
'        'Aired Date
'        gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), llADate
'        slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
'        'Post Log Date
'        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
'        slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
'        For ilRdfDay = 1 To 7 Step 1
'            If gWeekDayLong(llADate) = ilRdfDay - 1 Then
'                slRecord = slRecord & "Y" & sgCnCDelimiter
'            Else
'                slRecord = slRecord & "N" & sgCnCDelimiter
'            End If
'        Next ilRdfDay
'        gUnpackTime tmBof.iStartTime(0), tmBof.iStartTime(1), "A", "1", slTime
'        slTime = Format$(gConvertTime(slTime), "hhmmss")
'        slRecord = slRecord & slTime & sgCnCDelimiter
'        gUnpackTime tmBof.iEndTime(0), tmBof.iEndTime(1), "A", "1", slTime
'        slTime = Format$(gConvertTime(slTime), "hhmmss")
'        slRecord = slRecord & slTime & sgCnCDelimiter
'        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
'    Else
'        tmVefSrchKey.iCode = tmSdf.iVefCode
'        ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'        slRecord = slRecord & Trim$(tmSVef.sName) & sgCnCDelimiter
'        'Sold Date
'        If (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
'            tmSmfSrchKey2.lCode = tmSdf.lCode
'            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'            If ilRet = BTRV_ERR_NONE Then
'                gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llADate
'            Else
'                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
'            End If
'            llAWDate = llADate
'            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
'            'Aired Date
'            gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), llADate
'            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
'            'Post Log Date
'            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
'            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
'        Else
'            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
'            llAWDate = llADate
'            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
'            'Aired Date
'            gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), llADate
'            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
'            'Post Log Date
'            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
'            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
'        End If
'        'Daypart days sold
'        tmRdfSrchKey.iCode = tmClf.iRdfcode
'        ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'        For ilRdf = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1
'            If (tmRdf.iStartTime(0, ilRdf) <> 1) Or (tmRdf.iStartTime(1, ilRdf) <> 0) Then
'                If gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff) Then
'                    For ilRdfDay = LBound(tmCff.iDay) To UBound(tmCff.iDay) Step 1
'                        If (tmCff.iDay(ilRdfDay) = 1) Or (tmCff.sXDay(ilRdfDay) = "1") Then
'                            slRecord = slRecord & "Y" & sgCnCDelimiter
'                        Else
'                            slRecord = slRecord & "N" & sgCnCDelimiter
'                        End If
'                    Next ilRdfDay
'                Else
'                    For ilRdfDay = 1 To 7 Step 1
'                        slRecord = slRecord & tmRdf.sWkDays(ilRdf, ilRdfDay) & sgCnCDelimiter
'                    Next ilRdfDay
'                End If
'                If (tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0) Then
'                    gUnpackTime tmRdf.iStartTime(0, ilRdf), tmRdf.iStartTime(1, ilRdf), "A", "1", slTime
'                    slTime = Format$(gConvertTime(slTime), "hhmmss")
'                    slRecord = slRecord & slTime & sgCnCDelimiter
'                    gUnpackTime tmRdf.iEndTime(0, ilRdf), tmRdf.iEndTime(1, ilRdf), "A", "1", slTime
'                    slTime = Format$(gConvertTime(slTime), "hhmmss")
'                    slRecord = slRecord & slTime & sgCnCDelimiter
'                Else
'                    gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slTime
'                    slTime = Format$(gConvertTime(slTime), "hhmmss")
'                    slRecord = slRecord & slTime & sgCnCDelimiter
'                    gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slTime
'                    slTime = Format$(gConvertTime(slTime), "hhmmss")
'                    slRecord = slRecord & slTime & sgCnCDelimiter
'                End If
'                Exit For
'            End If
'        Next ilRdf
'        'Break Times
'        If (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
'            tmSmfSrchKey2.lCode = tmSdf.lCode
'            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'            If ilRet = BTRV_ERR_NONE Then
'                gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slTime
'            Else
'                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
'            End If
'        Else
'            gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
'        End If
'    End If
'    slTime = Format$(gConvertTime(slTime), "hhmmss")
'    slRecord = slRecord & slTime & sgCnCDelimiter
'    'Air Time
'    gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
'    slTime = Format$(gConvertTime(slTime), "hhmmss")
'    slRecord = slRecord & slTime & sgCnCDelimiter
'    'Post Log Time
'    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
'    slTime = Format$(gConvertTime(slTime), "hhmmss")
'    slRecord = slRecord & slTime & sgCnCDelimiter
'    'Program #
'    ''slRecord = slRecord & Trim$(tmDlf.sProgCode) & sgCnCDelimiter
'    ''slRecord = slRecord & Trim$(tmDlf.sProgCode) & sgCnCDelimiter
'    ''slRecord = slRecord & Trim$(tmDlf.sProgCode) & sgCnCDelimiter
'    'slRecord = slRecord & Trim$(Left$(slAvailComm, 5)) & sgCnCDelimiter
'    'slRecord = slRecord & Trim$(Left$(slAvailComm, 5)) & sgCnCDelimiter
'    'slRecord = slRecord & Trim$(Left$(slAvailComm, 5)) & sgCnCDelimiter
'    slRecord = slRecord & Trim$(slAvailComm) & Trim$(tmDlf.sBus) & sgCnCDelimiter
'    slRecord = slRecord & Trim$(slAvailComm) & Trim$(tmDlf.sBus) & sgCnCDelimiter
'    slRecord = slRecord & Trim$(slAvailComm) & Trim$(tmDlf.sBus) & sgCnCDelimiter
'    'Post Log Break
'    slRecord = slRecord & Trim$(str$(ilBreakNo)) & sgCnCDelimiter
'    slRecord = slRecord & Trim$(str$(ilPositionNo)) & sgCnCDelimiter
'    If ilBlackoutDefined Then
'        slRecord = slRecord & "N" & sgCnCDelimiter
'        slRecord = slRecord & "0" & sgCnCDelimiter
'        slRecord = slRecord & "0" & sgCnCDelimiter
'        slPrice = "Bonus"
'    Else
'        If (tmSdf.sSchStatus = "G") Then
'            slRecord = slRecord & "Y" & sgCnCDelimiter
'            slRecord = slRecord & "1" & sgCnCDelimiter
'            slRecord = slRecord & "1" & sgCnCDelimiter
'        Else
'            slRecord = slRecord & "N" & sgCnCDelimiter
'            slRecord = slRecord & "0" & sgCnCDelimiter
'            slRecord = slRecord & "0" & sgCnCDelimiter
'        End If
'        'Spot Price
'        ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
'        ilPos = InStr(slPrice, ".")
'        If ilPos > 0 Then
'            slPrice = Left$(slPrice, ilPos - 1) & right$(slPrice, 2)
'        End If
'    End If
'    slRecord = slRecord & Trim$(slPrice) & sgCnCDelimiter
'    'Airing Network ID
'    slRecord = slRecord & Trim$(str$(tmVef.iCode)) & sgCnCDelimiter
'    'Zone
'    slRecord = slRecord & Trim$(tmDlf.sZone) & sgCnCDelimiter
'    'Survey
'    If ilBlackoutDefined Then
'        slRecord = slRecord & "" & sgCnCDelimiter
'    Else
'        If tmClf.iDnfCode > 0 Then
'            tmDnfSrchKey.iCode = tmClf.iDnfCode
'            ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'            If ilRet = BTRV_ERR_NONE Then
'                slRecord = slRecord & Trim$(tmDnf.sBookName) & sgCnCDelimiter
'            Else
'                slRecord = slRecord & "" & sgCnCDelimiter
'            End If
'        Else
'            slRecord = slRecord & "" & sgCnCDelimiter
'        End If
'    End If
'    'Competitive
'    If tmChf.iMnfComp(0) > 0 Then
'        tmMnfSrchKey.iCode = tmChf.iMnfComp(0)
'        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'        If ilRet = BTRV_ERR_NONE Then
'            slRecord = slRecord & sgCnCNameBracket & Trim$(tmMnf.sName) & sgCnCNameBracket & sgCnCDelimiter
'        Else
'            slRecord = slRecord & "" & sgCnCDelimiter
'        End If
'    Else
'        slRecord = slRecord & "" & sgCnCDelimiter
'    End If
'    'Short Title
'    slRecord = slRecord & sgCnCNameBracket & Trim$(tmOdf.sShortTitle) & sgCnCNameBracket & sgCnCDelimiter
'    If ilBlackoutDefined Then
'        tmCifSrchKey.lCode = llBCifCode
'        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            If tmCif.lcpfCode > 0 Then
'                tmCpfSrchKey.lCode = tmCif.lcpfCode
'                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                If ilRet <> BTRV_ERR_NONE Then
'                    tmCpf.sISCI = ""
'                    tmCpf.sName = ""
'                    tmCpf.sCreative = ""
'                End If
'            Else
'                tmCpf.sISCI = ""
'                tmCpf.sName = ""
'                tmCpf.sCreative = ""
'            End If
'        Else
'            tmCpf.sISCI = ""
'            tmCpf.sName = ""
'            tmCpf.sCreative = ""
'        End If
'        'ISCI Code
'        slRecord = slRecord & Trim$(tmCpf.sISCI) & sgCnCDelimiter
'        'Spot Sold
'        slRecord = slRecord & sgCnCDelimiter
'    Else
'        ilRet = mObtainCopy(tmDlf.sZone)
'        'ISCI Code
'        slRecord = slRecord & Trim$(tmCpf.sISCI) & sgCnCDelimiter
'        'Spot Sold
'        If gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff) Then
'            slRecord = slRecord & Trim$(str$(tmCff.iSpotsWk)) & sgCnCDelimiter
'        Else
'            slRecord = slRecord & sgCnCDelimiter
'        End If
'    End If
'    'Sales Network ID
'    slRecord = slRecord & Trim$(str$(tmSdf.iVefCode))
'    Print #hmTo, slRecord
'    Return
'lProcSeqNo:
'    ilSeqNo = 0
'    'If tmVef.iVefCode > 0 Then  'Log vehicle defined
'    '    tmOdfSrchKey0.iVefCode = tmVef.iVefCode
'    '    ilTestVefCode = tmVef.iVefCode
'    'Else
'    '    tmOdfSrchKey0.iVefCode = ilVehCode
'    '    ilTestVefCode = ilVehCode
'    'End If
'    'tmOdfSrchKey0.iAirDate(0) = tmOdf.iAirDate(0)   'ilLogDate0
'    'tmOdfSrchKey0.iAirDate(1) = tmOdf.iAirDate(1)   'ilLogDate1
'    'tmOdfSrchKey0.iLocalTime(0) = tmOdf.iLocalTime(0)   'tmDlf.iLocalTime(0)
'    'tmOdfSrchKey0.iLocalTime(1) = tmOdf.iLocalTime(1)   'tmDlf.iLocalTime(1)
'    'tmOdfSrchKey0.sZone = tmOdf.sZone   'tmDlf.sZone
'    'tmOdfSrchKey0.iSeqNo = 32000
'    'ilRet = btrGetLessOrEqual(hmOdf, tlOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'    'If (ilRet = BTRV_ERR_NONE) And (tlOdf.iVefCode = ilTestVefCode) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1)) Then
'    '    If (tlOdf.iLocalTime(0) = tmOdf.iLocalTime(0)) And (tlOdf.iLocalTime(1) = tmOdf.iLocalTime(1)) And (tlOdf.sZone = tmOdf.sZone) Then
'    '        ilSeqNo = tlOdf.iSeqNo
'    '    End If
'    'End If
'    Return
'lProcAdjDate:
'    'Test if Air time is AM and Local Time is PM. If so, adjust date
'    ilAirHour = tmOdf.iAirTime(1) \ 256  'Obtain month
'    ilLocalHour = tmOdf.iLocalTime(1) \ 256  'Obtain month
'    If (ilAirHour < 6) And (ilLocalHour > 17) Then
'        'If monday convert to next sunday- this is wrong but the same spot
'        'runs each sunday (the spot should have show on the previous week sunday)
'        'If not monday, then subtract one day
'        gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAdjDate
'        If gWeekDayStr(slAdjDate) = 0 Then
'            slAdjDate = gObtainNextSunday(slAdjDate)
'        Else
'            slAdjDate = gDecOneDay(slAdjDate)
'        End If
'        gPackDate slAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
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
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)
    hmDnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imDnfRecLen = Len(tmDnf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imSifRecLen = Len(tmSif)
    hmRsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imRsfRecLen = Len(tmRsf)
    hmBof = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmBof, "", sgDBPath & "Bof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imBofRecLen = Len(tmBof)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imTzfRecLen = Len(tmTzf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)
    hmDlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmDlf, "", sgDBPath & "Dlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imDlfRecLen = Len(tmDlf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    hmCTSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCTSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    hmCef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imCefRecLen = Len(tmCef)
    hmCyf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCyf, "", sgDBPath & "Cyf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imCyfRecLen = Len(tmCyf)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imRdfRecLen = Len(tmRdf)
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSS
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)
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
    gCenterStdAlone ExpCnCSS
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
'*******************************************************
End Sub
'*******************************************************
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
    ''slToFile = sgExportPath & "ExpCncSS.Txt"
    slToFile = sgDBPath & "Messages\" & "ExpCncSS.Txt"
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
    
    'Print #hmMsg, "** Export CNC Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** Export CNC Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
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
'        imShowHelpMsg = True
'        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
'        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
'            imShowHelpMsg = False
'        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
        igCmmCallSource = Val(slStr)
'    End If
'    gInitStdAlone ExpCnCSS, slStr, ilTestSystem
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
    Unload ExpCnCSS
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
    
    'ilRet = gPopUserVehicleBox(ExpCnCSS, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    ilRet = gPopUserVehicleBox(ExpCnCSS, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExpCnCSS
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
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If ilVefCode = tgVff(ilVff).iVefCode Then
                If tgVff(ilVff).sExportCnCSpot = "Y" Then
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
    plcScreen.Print "Export Schedule Spots by Week"
End Sub

Private Sub mProcAdjDate(ilAirHour As Integer, ilLocalHour As Integer, slAdjDate As String)
    'Test if Air time is AM and Local Time is PM. If so, adjust date
    ilAirHour = tmOdf.iAirTime(1) \ 256  'Obtain month
    ilLocalHour = tmOdf.iLocalTime(1) \ 256  'Obtain month
    If (ilAirHour < 6) And (ilLocalHour > 17) Then
        'If monday convert to next sunday- this is wrong but the same spot
        'runs each sunday (the spot should have show on the previous week sunday)
        'If not monday, then subtract one day
        gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAdjDate
        If gWeekDayStr(slAdjDate) = 0 Then
            slAdjDate = gObtainNextSunday(slAdjDate)
        Else
            slAdjDate = gDecOneDay(slAdjDate)
        End If
        gPackDate slAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
    End If
End Sub

Private Sub mProcSeqNo(ilSeqNo As Integer)
    ilSeqNo = 0
End Sub

Private Sub mProcProg()
    Exit Sub
End Sub

Private Sub mProcSpot(ilLogDate0 As Integer, ilLogDate1 As Integer, slLength As String, ilVehCode As Integer, ilType As Integer, llBChfCode As Long, llBCifCode As Long, llBSifCode As Long, llBBofCode As Long, ilBlackoutDefined As Integer, llCifCode As Long, ilBreakNo As Integer, ilPositionNo As Integer, llCrfCsfCode As Long, ilAirHour As Integer, ilLocalHour As Integer, slAdjDate As String, ilSeqNo As Integer, slAvailComm As String, slRecord As String)
    Dim ilRet As Integer
    Dim llADate As Long
    Dim llAWDate As Long
    Dim ilRdf As Integer
    Dim ilRdfDay As Integer
    Dim slTime As String
    Dim slPrice As String
    Dim ilPos As Integer
    
    tmOdf.iUrfCode = tgUrf(0).iCode
    If tmVef.iVefCode > 0 Then  'Log vehicle defined
        tmOdf.iVefCode = tmVef.iVefCode
    Else
        tmOdf.iVefCode = ilVehCode
    End If
    tmOdf.iAirDate(0) = ilLogDate0
    tmOdf.iAirDate(1) = ilLogDate1
    tmOdf.iAirTime(0) = tmAvail.iTime(0)
    tmOdf.iAirTime(1) = tmAvail.iTime(1)
    tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
    tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
    tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
    tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
    tmOdf.sZone = tmDlf.sZone
    tmOdf.iEtfCode = 0
    tmOdf.iEnfCode = tmDlf.iEnfCode 'Required for Commercial schedule
    tmOdf.sProgCode = tmDlf.sProgCode
    tmOdf.iMnfFeed = tmDlf.iMnfFeed
    'tmOdf.sUnused1 = "" 'tmDlf.sBus
    tmOdf.iWkNo = 0    'tmDlf.sSchedule
    tmOdf.ianfCode = tmAvail.ianfCode
    tmOdf.iUnits = 0
    slLength = Trim$(str$(tmSdf.iLen)) & "s"
    gPackLength slLength, tmOdf.iLen(0), tmOdf.iLen(1)
    'Test if Blackout exist
    If gSDFBlackoutTest(hmRsf, hmBof, tmSdf.lCode, ilVehCode, llBChfCode, llBCifCode, llBSifCode, llBBofCode) Then
        ilBlackoutDefined = True
        tmChfSrchKey.lCode = llBChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tmAdfSrchKey.iCode = tmChf.iAdfCode
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        End If
        tmOdf.iAdfCode = tmChf.iAdfCode
        tmOdf.lCifCode = llBCifCode
        tmBofSrchKey.lCode = llBBofCode
        ilRet = btrGetEqual(hmBof, tmBof, imBofRecLen, tmBofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Else
        ilBlackoutDefined = False
        tmOdf.iAdfCode = tmChf.iAdfCode
        'Test tmSdf.sPtType
        tmOdf.lCifCode = llCifCode
    End If
    If Trim$(tmChf.sProduct) <> "" Then
        tmOdf.sProduct = tmChf.sProduct
    Else
        tmOdf.sProduct = "" '"???? Name ????"
    End If
    tmOdf.iMnfSubFeed = tmDlf.iMnfSubFeed
    tmOdf.lCntrNo = tmChf.lCntrNo
    tmOdf.iBreakNo = ilBreakNo
    tmOdf.iPositionNo = ilPositionNo
    tmOdf.iType = ilType
    If ilBlackoutDefined Then
        tmOdf.lCefCode = 0
        If tgSpf.sUseProdSptScr = "P" Then
            If llBSifCode <= 0 Then
                tmSifSrchKey.lCode = tmChf.lSifCode
            Else
                tmSifSrchKey.lCode = llBSifCode
            End If
            ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                tmOdf.sShortTitle = Trim$(tmSif.sName)
            Else
                tmOdf.sShortTitle = ""
            End If
        Else
            tmOdf.sShortTitle = Trim$(tmAdf.sAbbr) & "," & Trim$(tmChf.sProduct)
        End If
    Else
        tmOdf.lCefCode = llCrfCsfCode   '0
        If tmAdf.iCode <> tmChf.iAdfCode Then
            tmAdfSrchKey.iCode = tmChf.iAdfCode
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
        tmOdf.sShortTitle = gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)
    End If
    '6/2/16: Replaced GoSub
    'GoSub lProcAdjDate  'Adjust dates prior to adjusting seq numbers
    mProcAdjDate ilAirHour, ilLocalHour, slAdjDate
    '6/2/16: Replaced GoSub
    'GoSub lProcSeqNo
    mProcSeqNo ilSeqNo
    tmOdf.iSeqNo = ilSeqNo + 1
    'ilRet = btrInsert(hmOdf, tmOdf, imOdfRecLen, INDEXKEY0)
    slRecord = Trim$(str$(tmSdf.lCode)) & sgCnCDelimiter  '"SPOT,"
    slRecord = slRecord & Trim$(str$(tmChf.lCntrNo)) & sgCnCDelimiter
    If ilBlackoutDefined Then
        slRecord = slRecord & "0" & sgCnCDelimiter
    Else
        slRecord = slRecord & Trim$(str$(tmSdf.iLineNo)) & sgCnCDelimiter
    End If
    slRecord = slRecord & Trim$(str$(tmSdf.iLen)) & sgCnCDelimiter
    slRecord = slRecord & Trim$(str$(tmChf.iAdfCode)) & sgCnCDelimiter
    slRecord = slRecord & sgCnCNameBracket & Trim$(tmAdf.sName) & sgCnCNameBracket & sgCnCDelimiter
    slRecord = slRecord & sgCnCNameBracket & Trim$(tmChf.sProduct) & sgCnCNameBracket & sgCnCDelimiter
    'Selling Vehicle
    If ilBlackoutDefined Then
        tmVefSrchKey.iCode = tmSdf.iVefCode
        ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        slRecord = slRecord & Trim$(tmSVef.sName) & sgCnCDelimiter
    Else
        tmClfSrchKey.lChfCode = tmChf.lCode
        tmClfSrchKey.iLine = tmSdf.iLineNo
        tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
        tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) And (tmClf.iLine = tmSdf.iLineNo) Then
            tmVefSrchKey.iCode = tmClf.iVefCode
            ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            slRecord = slRecord & Trim$(tmSVef.sName) & sgCnCDelimiter
        Else
            slRecord = slRecord & "Sell Veh Missing" & sgCnCDelimiter
        End If
    End If
    'Airing Vehicle
    slRecord = slRecord & Trim$(tmVef.sName) & sgCnCDelimiter
    'Post Log Vehicle
    If ilBlackoutDefined Then
        slRecord = slRecord & Trim$(tmSVef.sName) & sgCnCDelimiter
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
        llAWDate = llADate
        slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
        'Aired Date
        gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), llADate
        slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
        'Post Log Date
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
        slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
        For ilRdfDay = 1 To 7 Step 1
            If gWeekDayLong(llADate) = ilRdfDay - 1 Then
                slRecord = slRecord & "Y" & sgCnCDelimiter
            Else
                slRecord = slRecord & "N" & sgCnCDelimiter
            End If
        Next ilRdfDay
        gUnpackTime tmBof.iStartTime(0), tmBof.iStartTime(1), "A", "1", slTime
        slTime = Format$(gConvertTime(slTime), "hhmmss")
        slRecord = slRecord & slTime & sgCnCDelimiter
        gUnpackTime tmBof.iEndTime(0), tmBof.iEndTime(1), "A", "1", slTime
        slTime = Format$(gConvertTime(slTime), "hhmmss")
        slRecord = slRecord & slTime & sgCnCDelimiter
        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
    Else
        tmVefSrchKey.iCode = tmSdf.iVefCode
        ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        slRecord = slRecord & Trim$(tmSVef.sName) & sgCnCDelimiter
        'Sold Date
        If (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
            tmSmfSrchKey2.lCode = tmSdf.lCode
            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llADate
            Else
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
            End If
            llAWDate = llADate
            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
            'Aired Date
            gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), llADate
            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
            'Post Log Date
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
        Else
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
            llAWDate = llADate
            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
            'Aired Date
            gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), llADate
            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
            'Post Log Date
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llADate
            slRecord = slRecord & Format$(llADate, "mmddyyyy") & sgCnCDelimiter
        End If
        'Daypart days sold
        tmRdfSrchKey.iCode = tmClf.iRdfCode
        ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        For ilRdf = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1
            If (tmRdf.iStartTime(0, ilRdf) <> 1) Or (tmRdf.iStartTime(1, ilRdf) <> 0) Then
                If gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff) Then
                    For ilRdfDay = LBound(tmCff.iDay) To UBound(tmCff.iDay) Step 1
                        If (tmCff.iDay(ilRdfDay) = 1) Or (tmCff.sXDay(ilRdfDay) = "1") Then
                            slRecord = slRecord & "Y" & sgCnCDelimiter
                        Else
                            slRecord = slRecord & "N" & sgCnCDelimiter
                        End If
                    Next ilRdfDay
                Else
                    For ilRdfDay = 1 To 7 Step 1
                        slRecord = slRecord & tmRdf.sWkDays(ilRdf, ilRdfDay - 1) & sgCnCDelimiter
                    Next ilRdfDay
                End If
                If (tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0) Then
                    gUnpackTime tmRdf.iStartTime(0, ilRdf), tmRdf.iStartTime(1, ilRdf), "A", "1", slTime
                    slTime = Format$(gConvertTime(slTime), "hhmmss")
                    slRecord = slRecord & slTime & sgCnCDelimiter
                    gUnpackTime tmRdf.iEndTime(0, ilRdf), tmRdf.iEndTime(1, ilRdf), "A", "1", slTime
                    slTime = Format$(gConvertTime(slTime), "hhmmss")
                    slRecord = slRecord & slTime & sgCnCDelimiter
                Else
                    gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slTime
                    slTime = Format$(gConvertTime(slTime), "hhmmss")
                    slRecord = slRecord & slTime & sgCnCDelimiter
                    gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slTime
                    slTime = Format$(gConvertTime(slTime), "hhmmss")
                    slRecord = slRecord & slTime & sgCnCDelimiter
                End If
                Exit For
            End If
        Next ilRdf
        'Break Times
        If (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
            tmSmfSrchKey2.lCode = tmSdf.lCode
            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then
                gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slTime
            Else
                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
            End If
        Else
            gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
        End If
    End If
    slTime = Format$(gConvertTime(slTime), "hhmmss")
    slRecord = slRecord & slTime & sgCnCDelimiter
    'Air Time
    gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
    slTime = Format$(gConvertTime(slTime), "hhmmss")
    slRecord = slRecord & slTime & sgCnCDelimiter
    'Post Log Time
    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
    slTime = Format$(gConvertTime(slTime), "hhmmss")
    slRecord = slRecord & slTime & sgCnCDelimiter
    'Program #
    ''slRecord = slRecord & Trim$(tmDlf.sProgCode) & sgCnCDelimiter
    ''slRecord = slRecord & Trim$(tmDlf.sProgCode) & sgCnCDelimiter
    ''slRecord = slRecord & Trim$(tmDlf.sProgCode) & sgCnCDelimiter
    'slRecord = slRecord & Trim$(Left$(slAvailComm, 5)) & sgCnCDelimiter
    'slRecord = slRecord & Trim$(Left$(slAvailComm, 5)) & sgCnCDelimiter
    'slRecord = slRecord & Trim$(Left$(slAvailComm, 5)) & sgCnCDelimiter
    slRecord = slRecord & Trim$(slAvailComm) & Trim$(tmDlf.sBus) & sgCnCDelimiter
    slRecord = slRecord & Trim$(slAvailComm) & Trim$(tmDlf.sBus) & sgCnCDelimiter
    slRecord = slRecord & Trim$(slAvailComm) & Trim$(tmDlf.sBus) & sgCnCDelimiter
    'Post Log Break
    slRecord = slRecord & Trim$(str$(ilBreakNo)) & sgCnCDelimiter
    slRecord = slRecord & Trim$(str$(ilPositionNo)) & sgCnCDelimiter
    If ilBlackoutDefined Then
        slRecord = slRecord & "N" & sgCnCDelimiter
        slRecord = slRecord & "0" & sgCnCDelimiter
        slRecord = slRecord & "0" & sgCnCDelimiter
        slPrice = "Bonus"
    Else
        If (tmSdf.sSchStatus = "G") Then
            slRecord = slRecord & "Y" & sgCnCDelimiter
            slRecord = slRecord & "1" & sgCnCDelimiter
            slRecord = slRecord & "1" & sgCnCDelimiter
        Else
            slRecord = slRecord & "N" & sgCnCDelimiter
            slRecord = slRecord & "0" & sgCnCDelimiter
            slRecord = slRecord & "0" & sgCnCDelimiter
        End If
        'Spot Price
        ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
        ilPos = InStr(slPrice, ".")
        If ilPos > 0 Then
            slPrice = Left$(slPrice, ilPos - 1) & right$(slPrice, 2)
        End If
    End If
    slRecord = slRecord & Trim$(slPrice) & sgCnCDelimiter
    'Airing Network ID
    slRecord = slRecord & Trim$(str$(tmVef.iCode)) & sgCnCDelimiter
    'Zone
    slRecord = slRecord & Trim$(tmDlf.sZone) & sgCnCDelimiter
    'Survey
    If ilBlackoutDefined Then
        slRecord = slRecord & "" & sgCnCDelimiter
    Else
        If tmClf.iDnfCode > 0 Then
            tmDnfSrchKey.iCode = tmClf.iDnfCode
            ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then
                slRecord = slRecord & Trim$(tmDnf.sBookName) & sgCnCDelimiter
            Else
                slRecord = slRecord & "" & sgCnCDelimiter
            End If
        Else
            slRecord = slRecord & "" & sgCnCDelimiter
        End If
    End If
    'Competitive
    If tmChf.iMnfComp(0) > 0 Then
        tmMnfSrchKey.iCode = tmChf.iMnfComp(0)
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If ilRet = BTRV_ERR_NONE Then
            slRecord = slRecord & sgCnCNameBracket & Trim$(tmMnf.sName) & sgCnCNameBracket & sgCnCDelimiter
        Else
            slRecord = slRecord & "" & sgCnCDelimiter
        End If
    Else
        slRecord = slRecord & "" & sgCnCDelimiter
    End If
    'Short Title
    slRecord = slRecord & sgCnCNameBracket & Trim$(tmOdf.sShortTitle) & sgCnCNameBracket & sgCnCDelimiter
    If ilBlackoutDefined Then
        tmCifSrchKey.lCode = llBCifCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
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
        Else
            tmCpf.sISCI = ""
            tmCpf.sName = ""
            tmCpf.sCreative = ""
        End If
        'ISCI Code
        slRecord = slRecord & Trim$(tmCpf.sISCI) & sgCnCDelimiter
        'Spot Sold
        slRecord = slRecord & sgCnCDelimiter
    Else
        ilRet = mObtainCopy(tmDlf.sZone)
        'ISCI Code
        slRecord = slRecord & Trim$(tmCpf.sISCI) & sgCnCDelimiter
        'Spot Sold
        If gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff) Then
            slRecord = slRecord & Trim$(str$(tmCff.iSpotsWk)) & sgCnCDelimiter
        Else
            slRecord = slRecord & sgCnCDelimiter
        End If
    End If
    'Sales Network ID
    slRecord = slRecord & Trim$(str$(tmSdf.iVefCode))
    Print #hmTo, slRecord

End Sub
