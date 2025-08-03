VERSION 5.00
Begin VB.Form ExptMP2 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5130
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
   ScaleHeight     =   5130
   ScaleWidth      =   9135
   Begin VB.PictureBox plcCalendarTo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   2880
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendarTo 
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
         Picture         =   "ExptMP2.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDateTo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   16
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDnTo 
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalUpTo 
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalNameTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   25
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmcEndDate 
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
      Left            =   4680
      Picture         =   "ExptMP2.frx":2E1A
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   390
      Width           =   195
   End
   Begin VB.TextBox edcEndDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   5
      Top             =   390
      Width           =   930
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   840
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   17
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
         Picture         =   "ExptMP2.frx":2F14
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   14
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
            TabIndex        =   15
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
         TabIndex        =   12
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   165
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3570
      Width           =   1410
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2475
      Top             =   4545
   End
   Begin VB.ListBox lbcMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   3720
      MultiSelect     =   2  'Extended
      TabIndex        =   19
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
      Left            =   2685
      Picture         =   "ExptMP2.frx":5D2E
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
      Left            =   1755
      MaxLength       =   10
      TabIndex        =   2
      Top             =   390
      Width           =   930
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "ExptMP2.frx":5E28
      Left            =   165
      List            =   "ExptMP2.frx":5E2A
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   780
      Width           =   3375
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
      Enabled         =   0   'False
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
      Left            =   3240
      TabIndex        =   9
      Top             =   4575
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
      Left            =   4680
      TabIndex        =   10
      Top             =   4575
      Width           =   1050
   End
   Begin VB.Label lacScreen 
      Caption         =   "Export Audio MP2"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   2070
   End
   Begin VB.Label lacEndDate 
      Appearance      =   0  'Flat
      Caption         =   "To"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3480
      TabIndex        =   4
      Top             =   375
      Width           =   465
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   180
      TabIndex        =   20
      Top             =   4185
      Width           =   8730
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   4500
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   195
      TabIndex        =   18
      Top             =   3975
      Width           =   8730
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Export Date- From"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   75
      TabIndex        =   1
      Top             =   375
      Width           =   1785
   End
End
Attribute VB_Name = "ExptMP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExptMP2.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export feed (for Dalet, Scott, Drake & Prophet) input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim hmMsg As Integer   'From file hanle
Dim lmNowDate As Long   'Todays date
'Required by gMakeSsf
Dim tmSsf As SSF                'SSF record image
Dim hmSsf As Integer
'Dim tmSsfOld As SSF
Dim tmSpot As CSPOTSS
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0 'ANF key record image
Dim imAdfRecLen As Integer  'ANF record length
'Copy Rotation
Dim tmCrfSrchKey As LONGKEY0  'CRF key record image
Dim tmCrfSrchKey1 As CRFKEY1  'CRF key record image
Dim hmCrf As Integer        'CRF Handle
Dim imCrfRecLen As Integer      'CRF record length
Dim tmCrf As CRF
'Copy/Product
Dim hmCpf As Integer
Dim tmCpfSrchKey As LONGKEY0  'CPF key record image
Dim tmCpf As CPF
Dim imCpfRecLen As Integer  'CPF record length
'Copy instruction record information
Dim hmCnf As Integer        'Copy instruction file handle
Dim tmCnfSrchKey As CNFKEY0 'CNF key record image
Dim imCnfRecLen As Integer  'CNF record length
Dim tmCnf As CNF            'CNF record image
'Copy inventory record information
Dim hmCif As Integer        'Copy line file handle
Dim tmCifSrchKey As LONGKEY0  'CIF key record image
Dim imCifRecLen As Integer  'CIF record length
Dim tmCif As CIF            'CIF record image
'Contract record information
Dim hmCHF As Integer        'Contract header file handle
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
Dim hmClf As Integer        'Contract header file handle
Dim imClfRecLen As Integer  'CHF record length
Dim tmClf As CLF            'CHF record image
'Short Title Vehicle Table record information
Dim hmVsf As Integer        'Short Title Vehicle Table file handle
Dim tmVsfSrchKey As LONGKEY0  'VSF key record image
Dim imVsfRecLen As Integer  'VSF record length
Dim tmVsf As VSF            'VSF record image
Dim hmSif As Integer        'Short Title Vehicle Table file handle
Dim imSifRecLen As Integer  'VSF record length
Dim tmSif As SIF            'VSF record image
' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer     'VEF record length
Dim smVehName As String
'Vehicle linkage record information
Dim hmVlf As Integer        'Vehicle linkage file handle
Dim imVlfRecLen As Integer  'VLF record length
Dim tmVlf As VLF            'VLF record image
'Spot record
Dim tmSdf As SDF
Dim hmSdf As Integer
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey3 As LONGKEY0
Dim tmSdfSrchKey1 As SDFKEY1
'5614 vff
Dim hmVff As Integer
Dim tmVff As VFF
Dim imVffRecLen As Integer

'9887
Dim hmCvf As Integer        'Copy Vehicle
Dim imCvfRecLen As Integer  'CVF record length
Dim tmCvf As CVF            'CVF record image
Dim tmCvfSrchKey As LONGKEY0
Dim ilVefCodesForCrf() As Integer
    
Dim tmAirSellLink() As AIRSELLLINK
Dim tmExptMP2Info() As EXPTMP2INFO
Dim lmCifCode() As Long
Dim lmCrfCode() As Long

Dim hmMP2 As Integer

Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmInputStartDate As Long    'Input Start Date
Dim lmInputEndDate As Long  'Input End Date
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer

Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)

' MsgBox parameters
Const vbOkOnly = 0                 ' OK button only
Const vbCritical = 16          ' Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0

Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    If lbcVehicle.ListCount <= 0 Then
        Exit Sub
    End If
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicle.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

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

Private Sub cmcCalDnTo_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendarTo_Paint
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
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

Private Sub cmcCalUpTo_Click()
    plcCalendar.Visible = False
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendarTo_Paint
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
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
    plcCalendarTo.Visible = False
End Sub

Private Sub cmcEndDate_Click()
    plcCalendarTo.Visible = Not plcCalendarTo.Visible
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
    mSetCommands
End Sub

Private Sub cmcEndDate_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
    mSetCommands
End Sub

Private Sub cmcExport_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim slStr As String
    Dim ilVefSelected As Integer

    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    lacProcessing.Caption = ""
    lacMsg.Caption = ""
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        Beep
        edcStartDate.SetFocus
        Exit Sub
    End If

    slStr = edcEndDate.Text
    If Not gValidDate(slStr) Then
        Beep
        edcEndDate.SetFocus
        Exit Sub
    End If
    slStr = Trim$(edcStartDate.Text)
    lmInputStartDate = gDateValue(slStr)
    slStr = Trim$(edcEndDate.Text)
    lmInputEndDate = gDateValue(slStr)
    ilVefSelected = False
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            ilVefSelected = True
        End If
    Next ilLoop
    If Not ilVefSelected Then
        Beep
        lbcVehicle.SetFocus
        Exit Sub
    End If

    lbcMsg.Clear
    If Not mOpenMsgFile() Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imExporting = True
    'Print #hmMsg, "Start Date " & edcStartDate.Text & " End Date " & edcEndDate.Text
    'gAutomationAlertAndLogHandler "Start Date " & edcStartDate.Text & " End Date " & edcEndDate.Text
    gAutomationAlertAndLogHandler "* Start Date = " & edcStartDate.Text
    gAutomationAlertAndLogHandler "* End Date = " & edcEndDate.Text
    If ckcAll.Value = vbChecked Then
        gAutomationAlertAndLogHandler "* All Vehicles = True" & edcEndDate.Text
    Else
        gAutomationAlertAndLogHandler "* All Vehicles = False" & edcEndDate.Text
    End If
    gAutomationAlertAndLogHandler "Getting Selling to Airing Links.."
    lacProcessing.Caption = "Getting Selling to Airing Links"
    DoEvents
    mBuildLinkTables
    gAutomationAlertAndLogHandler "Getting Rotations.."
    lacProcessing.Caption = "Getting Rotations"
    DoEvents
    mRotPop
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            smVehName = Trim$(slName)
            lacProcessing.Caption = "Processing: " & smVehName
            'Print #hmMsg, "Processing: " & smVehName
            gAutomationAlertAndLogHandler "Processing: " & smVehName
            DoEvents
            '6-1-05 assign copy to air time spots
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            tmVefSrchKey.iCode = ilVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then
                ilRet = mExptMP2(ilVefCode)
            Else                    'error, vehicle not found
                'Print #hmMsg, " "
                gAutomationAlertAndLogHandler " "
                'Print #hmMsg, "Name: " & slName & " not found"
                gAutomationAlertAndLogHandler "Name: " & slName & " not found"
                lbcMsg.AddItem slName & " not found: vehicle aborted"
            End If                  'ilret <> BTRV_err_none
        End If                  'lbcVehicle.Selected(ilLoop)
    Next ilLoop                 'next vehicle
    lacProcessing.Caption = "Completed"
    gAutomationAlertAndLogHandler "Completed Export."
    'Print #hmMsg, "** Completed " & smScreenCaption & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    Close #hmMsg
    On Error GoTo 0

    'lacProcessing.Caption = "Output for " & smScreenCaption & " sent to " & slToFile
    lacMsg.Caption = "Messages sent to " & sgDBPath & "Messages\" & "ExptMP2.Txt"
    Screen.MousePointer = vbDefault
    imExporting = False
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    Exit Sub
cmcExportErr: 'VBC NR
    ilRet = err.Number
    Resume Next

ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub
Private Sub cmcExport_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub cmcStartDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
    mSetCommands
End Sub
Private Sub cmcStartDate_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetCommands
End Sub

Private Sub edcEndDate_Change()
    Dim slStr As String
    plcCalendar.Visible = False

    slStr = edcEndDate.Text
    If Not gValidDate(slStr) Then
        lacDateTo.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendarTo_Paint   'mBoxCalDate called within paint
    mSetCommands
End Sub
Private Sub edcEndDate_Click()
    plcCalendar.Visible = False
    mSetCommands
End Sub

Private Sub edcEndDate_GotFocus()
    plcCalendar.Visible = False

    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDate
    mSetCommands
End Sub

Private Sub edcEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    mSetCommands
End Sub

Private Sub edcEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcEndDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    mSetCommands
End Sub

Private Sub edcEndDate_KeyUp(KeyCode As Integer, Shift As Integer)
Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendarTo.Visible = Not plcCalendarTo.Visible
        Else
            slDate = edcEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
    mSetCommands
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
    mSetCommands
End Sub
Private Sub edcStartDate_Click()
    mSetCommands
End Sub
Private Sub edcStartDate_GotFocus()
    plcCalendarTo.Visible = False
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDate
    mSetCommands
End Sub
Private Sub edcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    mSetCommands
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
    mSetCommands
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
    'Me.Visible = False
    'Me.Visible = True
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcCalendarTo.Visible = False
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
        'cmcCancel_Click
        tmcCancel.Enabled = True
        Me.Left = 2 * Screen.Width  'move off the screen so screen won't flash
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmAirSellLink
    Erase tmExptMP2Info
    Erase lmCifCode
    Erase lmCrfCode

    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmCnf)
    btrDestroy hmCnf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmVff)
    btrDestroy hmVff
    '9887
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    Set ExptCart = Nothing   'Remove data segment
    Set ExptMP2 = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lacEndDate_Click()
    mSetCommands
End Sub

Private Sub lacStartDate_Click()
    mSetCommands
End Sub

Private Sub lbcVehicle_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  '9-12-02 False
        imSetAll = True
    End If
    mSetCommands
End Sub
Private Sub lbcVehicle_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
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
Private Sub mBoxCalDate(EditDate As Control, LabelDate As Control)
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = EditDate.Text   'edcStartDate.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    LabelDate.Caption = slDay
                    LabelDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    LabelDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            LabelDate.Visible = False
        Else
            LabelDate.Visible = False
        End If
    Else
        LabelDate.Visible = False
    End If
End Sub


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
    Dim slStr As String
    imTerminate = False
    imFirstActivate = True
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imAllClicked = False
    imSetAll = True
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False

    '7496
    lacScreen.Caption = "Export Audio " & UCase(Mid(sgAudioExtension, 2))
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)
    hmCnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imCnfRecLen = Len(tmCnf)
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imSifRecLen = Len(tmSif)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)
    hmSdf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    '5614
    imVffRecLen = Len(tmVff)
    hmVff = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVff, "", sgDBPath & "Vff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptMP2
    On Error GoTo 0
    '9887
    hmCvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptCart
    On Error GoTo 0
    imCvfRecLen = Len(tmCvf)
    'Populate arrays to determine if records exist
    mVehPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        'mTerminate
        Exit Sub
    End If

    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    slStr = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slStr)
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    pbcCalendarTo_Paint
    lacDate.Visible = False
    lacDateTo.Visible = False
    gCenterStdAlone ExptMP2
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
    ''slToFile = sgExportPath & "ExptMP2.Txt"
    slToFile = sgDBPath & "Messages\" & "ExptMP2.Txt"
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
    
    'Print #hmMsg, "** Export MP2: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** Export MP2: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
Private Sub mSetCommands()
Dim ilEnabled As Integer
Dim ilLoop As Integer
    ilEnabled = False
    'at least one vehicle must be selected
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            ilEnabled = True
            Exit For
        End If
    Next ilLoop

    If ilEnabled Then
        If (Trim$(edcStartDate.Text) = "") Or (Trim$(edcEndDate.Text) = "") Then
            ilEnabled = False
        End If
    End If
    cmcExport.Enabled = ilEnabled
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
    igManUnload = YES
    Unload ExptMP2
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
    
    ilRet = gPopUserVehicleBox(ExptMP2, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)

    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExptMP2
        On Error GoTo 0
    End If
    
    For ilLoop = LBound(tgUserVehicle) To UBound(tgUserVehicle) - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If ilVefCode = tgVff(ilVff).iVefCode Then
                If tgVff(ilVff).sExportMP2 = "Y" Then
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
    mBoxCalDate edcStartDate, lacDate
End Sub

Private Sub pbcCalendarTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                edcEndDate.Text = Format$(llDate, "m/d/yy")
                edcEndDate.SelStart = 0
                edcEndDate.SelLength = Len(edcEndDate.Text)
                imBypassFocus = True
                edcEndDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcEndDate.SetFocus
End Sub
Private Sub pbcCalendarTo_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalNameTo.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendarTo, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate edcEndDate, lacDateTo
End Sub





Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mRotPop                         *
'*                                                     *
'*             Created:8/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain rotation specifications *
'*                      Same code is in BulkFeed.Frm   *
'*                                                     *
'*******************************************************
Private Sub mRotPop()


'
'   iRet = mRotPop
'   Where:
'
    Dim ilRet As Integer    'Return status
    Dim ilLoop As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim ilVefSelected As Integer
    Dim llRotStartDate As Long
    Dim llRotEndDate As Long
    Dim ilRotOk As Integer
    Dim llTstStartDate As Long
    Dim llTstEndDate As Long
    Dim ilLink As Integer
    Dim ilStartIndex As Integer
    Dim ilSelVefCode As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    '9887
    Dim ilCurrent As Integer
    Dim ilCvf As Integer
    Dim ilVefForCrf As Integer
    
    ReDim tmExptMP2Info(0 To 0) As EXPTMP2INFO

    btrExtClear hmCrf   'Clear any previous extend operation
    ilExtLen = Len(tmCrf)  'Extract operation record size
    tmCrfSrchKey1.sRotType = "A"
    tmCrfSrchKey1.iEtfCode = 0
    tmCrfSrchKey1.iEnfCode = 0
    tmCrfSrchKey1.iAdfCode = 0
    tmCrfSrchKey1.lChfCode = 0
    tmCrfSrchKey1.lFsfCode = 0
    tmCrfSrchKey1.iVefCode = 0
    tmCrfSrchKey1.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
    Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
    gPackDateLong lmInputEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Crf", "CrfStartDate")
    ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    gPackDateLong lmInputStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Crf", "CrfEndDate")
    'ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

    ilOffSet = 0
    ilRet = btrExtAddField(hmCrf, ilOffSet, ilExtLen)  'Extract start/end time, and days
    On Error GoTo mRotPopErr
    gBtrvErrorMsg ilRet, "mRotPop (btrExtAddField):" & "Crf.Btr", ExptMP2
    On Error GoTo 0
    'ilRet = btrExtGetNextExt(hmClf)    'Extract record
    ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtGetNext):" & "Crf.Btr", ExptMP2
        On Error GoTo 0
        'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
        If ilRet = BTRV_ERR_REJECT_COUNT Then
            ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
        End If
        Do While ilRet = BTRV_ERR_NONE
            llTstStartDate = lmInputStartDate
            llTstEndDate = lmInputEndDate
            ilRotOk = True
            gUnpackDateLong tmCrf.iStartDate(0), tmCrf.iStartDate(1), llRotStartDate
            gUnpackDateLong tmCrf.iEndDate(0), tmCrf.iEndDate(1), llRotEndDate

            If ((llRotEndDate >= lmInputStartDate) And (llRotStartDate <= lmInputEndDate)) And (tmCrf.sState <> "D") And (tmCrf.sZone <> "R") Then
                '9887 the crf may not have one vehicle tied to it, but many
                ReDim ilVefCodesForCrf(0 To 0)
                ilCurrent = 0
                If tmCrf.iVefCode < 1 Then
                    imCvfRecLen = Len(tmCvf)
                    tmCvfSrchKey.lCode = tmCrf.lCvfCode
                    ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Do While ilRet = BTRV_ERR_NONE
                        For ilCvf = 0 To 99 Step 1
                            If tmCvf.iVefCode(ilCvf) > 0 Then
                                ilVefCodesForCrf(ilCurrent) = tmCvf.iVefCode(ilCvf)
                                ilCurrent = ilCurrent + 1
                                ReDim Preserve ilVefCodesForCrf(ilCurrent)
                            End If
                        Next ilCvf
                        If tmCvf.lLkCvfCode <= 0 Then
                            Exit Do
                        End If
                        tmCvfSrchKey.lCode = tmCvf.lLkCvfCode
                        ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                Else
                    ilVefCodesForCrf(0) = tmCrf.iVefCode
                    ReDim Preserve ilVefCodesForCrf(UBound(ilVefCodesForCrf) + 1)
                End If
                '9887 now fill packages with their hidden lines
                For ilVefForCrf = 0 To UBound(ilVefCodesForCrf) - 1
                    ilRet = gBinarySearchVef(ilVefCodesForCrf(ilVefForCrf))
                    If ilRet <> -1 Then
                        If tgMVef(ilRet).sType = "P" Then
                            mExpandPackage tmCrf.lChfCode, ilVefCodesForCrf, ilVefForCrf
                        End If
                    End If
                Next ilVefForCrf
                For ilVefForCrf = 0 To UBound(ilVefCodesForCrf) - 1
                    ilRotOk = mSpotExist(tmCrf.iAdfCode, tmCrf.lChfCode, ilVefCodesForCrf(ilVefForCrf), llTstStartDate, llTstEndDate)
                    If Not ilRotOk Then
'                        ilRet = gBinarySearchVef(tmCrf.iVefCode)
                        ilRet = gBinarySearchVef(ilVefCodesForCrf(ilVefForCrf))
                        If ilRet <> -1 Then
                            'copy can use airing vehicles, but they won't have spots assigned to them.
                            If tgMVef(ilRet).sType = "A" Then
                                ilRotOk = True
                            End If
                        End If
                    End If
                    If ilRotOk Then
                        Exit For
                    End If
                Next ilVefForCrf
'before 9887
'                ilRotOk = mSpotExist(tmCrf.iAdfCode, tmCrf.lChfCode, tmCrf.iVefCode, llTstStartDate, llTstEndDate)
'                If Not ilRotOk Then
'                    ilRet = gBinarySearchVef(tmCrf.iVefCode)
'                    If ilRet <> -1 Then
'                        If tgMVef(ilRet).sType = "A" Then
'                            ilRotOk = True
'                        End If
'                    End If
'                End If
            Else
                ilRotOk = False
            End If
            ilStartIndex = 0
            If ilRotOk Then
                Do
                    ilVefSelected = False
                    For ilLoop = ilStartIndex To lbcVehicle.ListCount - 1 Step 1
                        If lbcVehicle.Selected(ilLoop) Then
                            slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            ilSelVefCode = Val(slCode)
                            '9887
'                            If tmCrf.iVefCode = ilSelVefCode Then
'                                ilStartIndex = ilLoop
'                                ilVefSelected = True
'                                Exit For
'                            End If
                            For ilVefForCrf = 0 To UBound(ilVefCodesForCrf) - 1
                                 If ilVefCodesForCrf(ilVefForCrf) = ilSelVefCode Then
                                    ilStartIndex = ilLoop
                                    ilVefSelected = True
                                    Exit For
                                End If
                            Next ilVefForCrf
                            If ilVefSelected Then
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If Not ilVefSelected Then
                        For ilLoop = ilStartIndex To lbcVehicle.ListCount - 1 Step 1
                            If lbcVehicle.Selected(ilLoop) Then
                                slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                'Check if airing
                                ilSelVefCode = Val(slCode)
                                ilRet = gBinarySearchVef(Val(slCode))
                                If ilRet <> -1 Then
                                    If tgMVef(ilRet).sType = "A" Then
                                        '9887
                                        For ilLink = 0 To UBound(tmAirSellLink) - 1 Step 1
                                            If tmAirSellLink(ilLink).iAirVefCode = ilSelVefCode Then
                                                For ilVefForCrf = 0 To UBound(ilVefCodesForCrf) - 1
                                                    If tmAirSellLink(ilLink).iSellVefCode = ilVefCodesForCrf(ilVefForCrf) Then
                                                        ilStartIndex = ilLoop
                                                        ilVefSelected = True
                                                        Exit For
                                                    End If
                                                Next ilVefForCrf
                                                If ilVefSelected = True Then
                                                    Exit For
                                                End If
                                            End If
'                                         For ilLink = 0 To UBound(tmAirSellLink) - 1 Step 1
'                                            If tmAirSellLink(ilLink).iAirVefCode = ilSelVefCode Then
'                                                If tmAirSellLink(ilLink).iSellVefCode = tmCrf.iVefCode Then
'                                                    ilStartIndex = ilLoop
'                                                    ilVefSelected = True
'                                                    Exit For
'                                                End If
'                                            End If
                                        Next ilLink
                                        If ilVefSelected = True Then
                                            Exit For
                                        End If
                                    '9887 to do packages
                                    Else
                                    End If
                                End If
                            End If
                        Next ilLoop
                    End If
                    If ilVefSelected Then
                        tmExptMP2Info(UBound(tmExptMP2Info)).iVefCode = ilSelVefCode
                        tmExptMP2Info(UBound(tmExptMP2Info)).lCrfCode = tmCrf.lCode
                        ReDim Preserve tmExptMP2Info(0 To UBound(tmExptMP2Info) + 1) As EXPTMP2INFO
                    End If
                    ilStartIndex = ilStartIndex + 1
                Loop While ilVefSelected
            End If
            ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            End If
        Loop
    End If
    Erase ilVefCodesForCrf
    Exit Sub
mRotPopErr:
    On Error GoTo 0
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSpotExist                      *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if spot exist for    *
'*                      contract within Dates          *
'*                                                     *
'*******************************************************
Private Function mSpotExist(ilAdfCode As Integer, llChfCode As Long, ilVefCode As Integer, llInputStartDate As Long, llInputEndDate As Long) As Integer
    Dim ilRet As Integer
    'Dim tlSsfSrchKey As SSFKEY0 'SSF key record image
    Dim tlSsfSrchKey2 As SSFKEY2
    Dim ilSsfRecLen As Integer  'SSF record length
    Dim llDate As Long
    Dim ilSsfDate0 As Integer
    Dim ilSsfDate1 As Integer
    Dim ilType As Integer
    Dim ilEvt As Integer
    ilType = 0
    For llDate = llInputStartDate To llInputEndDate Step 1
        ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
        gPackDateLong llDate, ilSsfDate0, ilSsfDate1
        'tlSsfSrchKey.iType = ilType
        'tlSsfSrchKey.iVefCode = ilVefCode
        'tlSsfSrchKey.iDate(0) = ilSsfDate0
        'tlSsfSrchKey.iDate(1) = ilSsfDate1
        'tlSsfSrchKey.iStartTime(0) = 0
        'tlSsfSrchKey.iStartTime(1) = 0
        'ilRet = gSSFGetEqual(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        tlSsfSrchKey2.iVefCode = ilVefCode
        tlSsfSrchKey2.iDate(0) = ilSsfDate0
        tlSsfSrchKey2.iDate(1) = ilSsfDate1
        ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVefCode) Then
            ilType = tmSsf.iType
        End If
        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilSsfDate0) And (tmSsf.iDate(1) = ilSsfDate1)
            ilEvt = 1
            Do While ilEvt <= tmSsf.iCount
               LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                If ((tmSpot.iRecType And &HF) >= 10) And ((tmSpot.iRecType And &HF) <= 11) Then
                    If tmSpot.iAdfCode = ilAdfCode Then
                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                        If (ilRet = BTRV_ERR_NONE) And (tmSdf.lChfCode = llChfCode) Then
                            mSpotExist = True
                            Exit Function
                        End If
                    End If
                End If
                ilEvt = ilEvt + 1
            Loop
            ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilRet = gSSFGetNext(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            If (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVefCode) Then
                ilType = tmSsf.iType
            End If
        Loop
    Next llDate
    For llDate = llInputStartDate To llInputEndDate Step 1
        gPackDateLong llDate, ilSsfDate0, ilSsfDate1
        tmSdfSrchKey1.iVefCode = ilVefCode
        tmSdfSrchKey1.iDate(0) = ilSsfDate0
        tmSdfSrchKey1.iDate(1) = ilSsfDate1
        tmSdfSrchKey1.iTime(0) = 0
        tmSdfSrchKey1.iTime(1) = 0
        tmSdfSrchKey1.sSchStatus = "M"
        'ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.iDate(0) = ilSsfDate0) And (tmSdf.iDate(1) = ilSsfDate1)
            If (tmSdf.sSchStatus = "M") Or (tmSdf.sSchStatus = "C") Or (tmSdf.sSchStatus = "H") Or (tmSdf.sSchStatus = "U") Or (tmSdf.sSchStatus = "R") Then
                If tmSdf.lChfCode = llChfCode Then
                    mSpotExist = True
                    Exit Function
                End If
            End If
            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next llDate
    mSpotExist = False
    Exit Function
End Function



Private Sub mBuildLinkTables()
    Dim slStartDate As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilAirVefCode As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    ReDim ilVefCode(0 To 0) As Integer
    ReDim tmAirSellLink(0 To 0) As AIRSELLLINK

    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilAirVefCode = Val(slCode)
            tmVefSrchKey.iCode = ilAirVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If tmVef.sType = "A" Then
                slStartDate = Format$(lmInputStartDate, "m/d/yy")
                gBuildLinkArray hmVlf, tmVef, slStartDate, ilVefCode()
                For ilVef = 0 To UBound(ilVefCode) - 1 Step 1
                    tmAirSellLink(UBound(tmAirSellLink)).iAirVefCode = ilAirVefCode
                    tmAirSellLink(UBound(tmAirSellLink)).iSellVefCode = ilVefCode(ilVef)
                    ReDim Preserve tmAirSellLink(0 To UBound(tmAirSellLink) + 1) As AIRSELLLINK
                Next ilVef
            End If
        End If
    Next ilLoop
End Sub

Private Function mExptMP2(ilVefCode As Integer) As Integer
    Dim ilCrf As Integer
    Dim ilCif As Integer
    Dim ilFound As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim slShortTitle As String
    Dim slISCI As String
    Dim slDateTime As String
    Dim llSifCode As Long
    Dim ilVsf As Integer
    'ttp 5614
    Dim slPrefix As String
    
    ReDim lmCifCode(0 To 0) As Long
    ReDim lmCrfCode(0 To 0) As Long
    slStartDate = Format$(lmInputStartDate, "YYMMDD")
    slEndDate = Format$(lmInputEndDate, "YYMMDD")
    'Open Write file
    'On Error GoTo mExptMP2Err
    ilRet = 0
    slStr = sgExportPath & Trim$(smVehName) & "_" & slStartDate & "_" & slEndDate & ".bul"
    'slDateTime = FileDateTime(slStr)
    ilRet = gFileExist(slStr)
    If ilRet = 0 Then
        Kill slStr
    End If
    ilRet = 0
    'hmMP2 = FreeFile
    'Open slStr For Output As hmMP2
    ilRet = gFileOpen(slStr, "Output", hmMP2)
    If ilRet <> 0 Then
        Close #hmMP2
        'Print #hmMsg, "Open File Error: " & ilRet & " on " & slStr
        gAutomationAlertAndLogHandler "Open File Error: " & ilRet & " on " & slStr
        lbcMsg.AddItem "Open File Error: " & ilRet & " on " & slStr
        mExptMP2 = False
        Exit Function
    End If
    'Get Unique CIF's
    For ilCrf = 0 To UBound(tmExptMP2Info) - 1 Step 1
        If tmExptMP2Info(ilCrf).iVefCode = ilVefCode Then
            tmCnfSrchKey.lCrfCode = tmExptMP2Info(ilCrf).lCrfCode
            tmCnfSrchKey.iInstrNo = 0
            ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmExptMP2Info(ilCrf).lCrfCode)
                ilFound = False
                For ilCif = 0 To UBound(lmCifCode) - 1 Step 1
                    If tmCnf.lCifCode = lmCifCode(ilCif) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilCif
                If Not ilFound Then
                    lmCifCode(UBound(lmCifCode)) = tmCnf.lCifCode
                    ReDim Preserve lmCifCode(0 To UBound(lmCifCode) + 1) As Long
                    lmCrfCode(UBound(lmCrfCode)) = tmExptMP2Info(ilCrf).lCrfCode
                    ReDim Preserve lmCrfCode(0 To UBound(lmCrfCode) + 1) As Long
                End If
                ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    Next ilCrf
    'Create Image
    For ilCif = 0 To UBound(lmCifCode) - 1 Step 1
        tmCrfSrchKey.lCode = lmCrfCode(ilCif)
        ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tmChfSrchKey.lCode = tmCrf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If tmAdf.iCode <> tmChf.iAdfCode Then
                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                End If
                llSifCode = 0
                If tmChf.lVefCode < 0 Then
                    tmVsfSrchKey.lCode = -tmChf.lVefCode
                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Do While ilRet = BTRV_ERR_NONE
                        For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                            If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
                                If tmVsf.lFSComm(ilVsf) > 0 Then
                                    llSifCode = tmVsf.lFSComm(ilVsf)
                                End If
                                Exit For
                            End If
                        Next ilVsf
                        If llSifCode <> 0 Then
                            Exit Do
                        End If
                        If tmVsf.lLkVsfCode <= 0 Then
                            Exit Do
                        End If
                        tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                '7219
                If ((Asc(tgSpf.sUsingFeatures10) And ADDADVTTOISCI) = ADDADVTTOISCI) Then
                    slShortTitle = gXDSShortTitle(tmAdf, "", False, False)
                Else
                    slShortTitle = gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf, 6)
                    slShortTitle = UCase$(slShortTitle)
                End If
'                slShortTitle = gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf, 6)
'                If ((Asc(tgSpf.sUsingFeatures10) And ADDADVTTOISCI) = ADDADVTTOISCI) Then
'                    '2/7/13: Use Advertiser name only because XDS limited to 32 characters
'                    'slShortTitle = UCase$(Left$(slShortTitle, 15))
'                    If Trim$(tmAdf.sAbbr) <> "" Then
'                        slShortTitle = Left$(UCase(Trim(tmAdf.sAbbr)), 6)
'                    Else
'                        slShortTitle = UCase(Trim(Left(tmAdf.sName, 6)))
'                    End If
'                End If
'                slShortTitle = UCase$(slShortTitle)
                tmCifSrchKey.lCode = lmCifCode(ilCif)
                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        'ttp 5614 add vehicle prefix
                        slPrefix = mGetISCIPrefix(ilVefCode)
                        slISCI = Trim$(tmCpf.sISCI)
                        '7496
                        slStr = gFileNameFilter(slShortTitle) & "(" & slPrefix & gFileNameFilter(slISCI) & ")" & sgAudioExtension
                        'slStr = gFileNameFilter(slShortTitle) & "(" & slPrefix & gFileNameFilter(slISCI) & ").mp2"
                        Print #hmMP2, slStr
                    End If
                End If
            End If
        End If
    Next ilCif
    Close #hmMP2
    mExptMP2 = True
    Exit Function
'mExptMP2Err:
'    ilRet = Err.Number
'    Resume Next
End Function
Private Function mGetISCIPrefix(ilVefCode As Integer) As String
    Dim ilRet As Integer
    
    ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, ilVefCode, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        mGetISCIPrefix = gFileNameFilter(Trim$(tmVff.sXDISCIPrefix))
    End If
    'If Len(mGetISCIPrefix) > 0 Then
    '    mGetISCIPrefix = mGetISCIPrefix & "_"
    'End If
End Function
Private Sub mExpandPackage(llChfCode As Long, ilVefCode() As Integer, ilVefIndex As Integer)
    '9887
    Dim blFirstVehicle As Boolean
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim ilTest As Integer
    Dim ilPkVefCode As Integer
    Dim tlClf() As CLFLIST
    
    blFirstVehicle = True
    ilPkVefCode = ilVefCode(ilVefIndex)
    ilRet = gObtainChfClf(hmCHF, hmClf, llChfCode, 0, tmChf, tlClf)
    If ilRet = -1 Then
        ilVefCode(ilVefIndex) = 0
        For ilClf = 0 To UBound(tlClf) - 1 Step 1
            If tlClf(ilClf).ClfRec.iVefCode = ilPkVefCode Then
                For ilTest = 0 To UBound(tlClf) - 1 Step 1
                    If tlClf(ilTest).ClfRec.iPkLineNo = tlClf(ilClf).ClfRec.iLine Then
                        If mOkToAddStationVehicle(tlClf(ilTest).ClfRec.iVefCode, ilVefCode()) Then
                            If blFirstVehicle Then
                                ilVefCode(ilVefIndex) = tlClf(ilTest).ClfRec.iVefCode
                                blFirstVehicle = False
                            Else
                                ilVefCode(UBound(ilVefCode)) = tlClf(ilTest).ClfRec.iVefCode
                                ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                            End If
                        End If
                    End If
                Next ilTest
            End If
        Next ilClf
    Else
        ilVefCode(ilVefIndex) = 0
    End If

End Sub
Private Function mOkToAddStationVehicle(ilStationVefCode As Integer, ilVefCode() As Integer) As Boolean
    '9887
    Dim ilVef As Integer
    
    mOkToAddStationVehicle = True
    For ilVef = 0 To UBound(ilVefCode) - 1 Step 1
        If ilStationVefCode = ilVefCode(ilVef) Then
            mOkToAddStationVehicle = False
            Exit Function
        End If
    Next ilVef
End Function
