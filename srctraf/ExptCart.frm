VERSION 5.00
Begin VB.Form ExptCart 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5655
   ClientLeft      =   525
   ClientTop       =   3375
   ClientWidth     =   7275
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
   ScaleHeight     =   5655
   ScaleWidth      =   7275
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   2235
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   5
      Top             =   975
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
         TabIndex        =   8
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
         TabIndex        =   6
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
         Picture         =   "ExptCart.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   9
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
            TabIndex        =   10
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
         TabIndex        =   7
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcSelect 
      ForeColor       =   &H00000000&
      Height          =   3990
      Left            =   465
      ScaleHeight     =   3930
      ScaleWidth      =   6300
      TabIndex        =   1
      Top             =   420
      Width           =   6360
      Begin VB.CheckBox ckcAll 
         Caption         =   "All"
         Height          =   225
         Left            =   195
         TabIndex        =   28
         Top             =   3555
         Width           =   1785
      End
      Begin VB.ListBox lbcVehicle 
         Height          =   2010
         ItemData        =   "ExptCart.frx":2E1A
         Left            =   180
         List            =   "ExptCart.frx":2E1C
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   1395
         Width           =   4440
      End
      Begin VB.CheckBox ckcStatus 
         Caption         =   "History"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3975
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   990
         Width           =   915
      End
      Begin VB.CheckBox ckcStatus 
         Caption         =   "Purged"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2985
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   990
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox ckcStatus 
         Caption         =   "Active"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2055
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   990
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.TextBox edcAddedDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   3
         Top             =   285
         Width           =   930
      End
      Begin VB.CommandButton cmcAddedDate 
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
         Index           =   0
         Left            =   2685
         Picture         =   "ExptCart.frx":2E1E
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   285
         Width           =   195
      End
      Begin VB.CommandButton cmcAddedDate 
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
         Index           =   1
         Left            =   5670
         Picture         =   "ExptCart.frx":2F18
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   285
         Width           =   195
      End
      Begin VB.TextBox edcAddedDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   4740
         MaxLength       =   10
         TabIndex        =   12
         Top             =   285
         Width           =   930
      End
      Begin VB.CommandButton cmcPurgedDate 
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
         Index           =   0
         Left            =   2685
         Picture         =   "ExptCart.frx":3012
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   630
         Width           =   195
      End
      Begin VB.TextBox edcPurgedDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   15
         Top             =   630
         Width           =   930
      End
      Begin VB.CommandButton cmcPurgedDate 
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
         Index           =   1
         Left            =   5670
         Picture         =   "ExptCart.frx":310C
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   630
         Width           =   195
      End
      Begin VB.TextBox edcPurgedDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   4725
         MaxLength       =   10
         TabIndex        =   18
         Top             =   630
         Width           =   930
      End
      Begin VB.Label lacStatus 
         Appearance      =   0  'Flat
         Caption         =   "Cart Types to Include"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   975
         Width           =   5685
      End
      Begin VB.Label lacAddedDate 
         Appearance      =   0  'Flat
         Caption         =   "Added From Date"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lacAddedDate 
         Appearance      =   0  'Flat
         Caption         =   "Added To Date"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   3300
         TabIndex        =   11
         Top             =   270
         Width           =   1440
      End
      Begin VB.Label lacPurgedDate 
         Appearance      =   0  'Flat
         Caption         =   "Purged From Date"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label lacPurgedDate 
         Appearance      =   0  'Flat
         Caption         =   "Purged To Date"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   3300
         TabIndex        =   17
         Top             =   630
         Width           =   1440
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
      Left            =   30
      ScaleHeight     =   270
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   15
      Width           =   2175
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
      Left            =   2415
      TabIndex        =   24
      Top             =   5070
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
      Left            =   3750
      TabIndex        =   25
      Top             =   5070
      Width           =   1050
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   4980
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   105
      TabIndex        =   26
      Top             =   4530
      Width           =   7035
   End
End
Attribute VB_Name = "ExptCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ExptCart.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExptCart.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the agency conversion input screen code
Option Explicit
Option Compare Text
Dim imDateBox As Integer    '0=Added From Date; 1=Added To Date; 2=Purged From Date; 3=Purged To Date
Dim hmTo As Integer   'From file hanle
Dim hmMsg As Integer
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0 'ANF key record image
Dim imAdfRecLen As Integer  'ANF record length
'Media Code
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey As INTKEY0  'MCF key record image
Dim hmMcf As Integer        'MCF Handle
Dim imMcfRecLen As Integer      'MCF record length
'Vehicle
Dim hmVef As Integer
Dim tmVef As VEF
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
'Copy Rotation
Dim tmCrf As CRF            'CPF record image
Dim tmCrfSrchKey As LONGKEY0  'CPF key record image
Dim hmCrf As Integer        'CPF Handle
Dim imCrfRecLen As Integer      'CPF record length
'Inventory
Dim tmCif As CIF            'CIF record image
Dim hmCif As Integer        'CIF Handle
Dim imCifRecLen As Integer      'CIF record length
'Product
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0  'CPF key record image
Dim hmCpf As Integer        'CPF Handle
Dim imCpfRecLen As Integer      'CPF record length
'Copy Usage
Dim hmCuf As Integer        'Copy Usage file handle
Dim imCufRecLen As Integer  'CUF record length
Dim tmCuf As CUF            'CUF record image
Dim tmCufSrchKey1 As CUFKEY1
'Short Title record information
Dim hmSif As Integer        'Short Title file handle
Dim imSifRecLen As Integer  'SIF record length
Dim tmSif As SIF            'SIF record image
'9887
Dim hmCvf As Integer        'Copy Vehicle
Dim imCvfRecLen As Integer  'CVF record length
Dim tmCvf As CVF            'CVF record image
Dim tmCvfSrchKey As LONGKEY0

Dim smNowDate As String
Dim smGenDate As String
Dim smGenTime As String
Dim imSeqNo As Integer
Dim imFirstActivate As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim tmUserVehicle() As SORTCODE
Dim smUserVehicleTag As String
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim smHideMediaCode As String

Private Type CARTINFO
    sKey As String * 10
    sRecord As String * 120
End Type

Private tmCartInfo() As CARTINFO

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
'' MsgBox parameters
'Const vbOkOnly = 0                 ' OK button only
'Const vbCritical = 16          ' Critical message
'Const vbApplicationModal = 0
'Const INDEXKEY0 = 0
Private Sub cmcAddedDate_Click(Index As Integer)
    plcCalendar.Visible = Not plcCalendar.Visible
    edcAddedDate(Index).SelStart = 0
    edcAddedDate(Index).SelLength = Len(edcAddedDate(Index).Text)
    edcAddedDate(Index).SetFocus
End Sub
Private Sub cmcAddedDate_GotFocus(Index As Integer)
    Dim slStr As String
    If imDateBox <> Index Then
        plcCalendar.Visible = False
        slStr = edcAddedDate(Index).Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    imDateBox = Index
    gCtrlGotFocus ActiveControl
    plcCalendar.Move plcSelect.Left + edcAddedDate(Index).Left, plcSelect.Top + edcAddedDate(Index).Top + edcAddedDate(Index).Height
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    If imDateBox <= 1 Then
        edcAddedDate(imDateBox).SelStart = 0
        edcAddedDate(imDateBox).SelLength = Len(edcAddedDate(imDateBox).Text)
        edcAddedDate(imDateBox).SetFocus
    ElseIf imDateBox <= 3 Then
        edcPurgedDate(imDateBox - 2).SelStart = 0
        edcPurgedDate(imDateBox - 2).SelLength = Len(edcPurgedDate(imDateBox - 2).Text)
        edcPurgedDate(imDateBox - 2).SetFocus
    End If
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    If imDateBox <= 1 Then
        edcAddedDate(imDateBox).SelStart = 0
        edcAddedDate(imDateBox).SelLength = Len(edcAddedDate(imDateBox).Text)
        edcAddedDate(imDateBox).SetFocus
    ElseIf imDateBox <= 3 Then
        edcPurgedDate(imDateBox - 2).SelStart = 0
        edcPurgedDate(imDateBox - 2).SelLength = Len(edcPurgedDate(imDateBox - 2).Text)
        edcPurgedDate(imDateBox - 2).SetFocus
    End If
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
    Dim slStr As String
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    slStr = edcAddedDate(0).Text
    If slStr <> "" Then
        If Not gValidDate(slStr) Then
            Beep
            edcAddedDate(0).SetFocus
            Exit Sub
        End If
    End If
    slStr = edcAddedDate(1).Text
    If slStr <> "" Then
        If Not gValidDate(slStr) Then
            Beep
            edcAddedDate(1).SetFocus
            Exit Sub
        End If
    End If
    slStr = edcPurgedDate(0).Text
    If slStr <> "" Then
        If Not gValidDate(slStr) Then
            Beep
            edcPurgedDate(0).SetFocus
            Exit Sub
        End If
    End If
    slStr = edcPurgedDate(1).Text
    If slStr <> "" Then
        If Not gValidDate(slStr) Then
            Beep
            edcPurgedDate(1).SetFocus
            Exit Sub
        End If
    End If
    smGenTime = Format$(gNow(), "hh:mm")
    smGenDate = Format$(gNow(), "mm/dd/yyyy")
    gObtainYearMonthDayStr smGenDate, True, slFYear, slFMonth, slFDay
    'slToFile = sgExportPath & Right$(slFYear, 2) & slFMonth & slFDay & Left$(smGenTime, 2) & ".xrm"
    slToFile = sgExportPath & slFMonth & slFDay & right$(slFYear, 2) & Left$(smGenTime, 2) & ".xrm"
    ilRet = 0
    'On Error GoTo cmcExportErr:
    'hmTo = FreeFile
    'Open slToFile For Output As hmTo
    ilRet = gFileOpen(slToFile, "Output", hmTo)
    If ilRet <> 0 Then
        ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    If Not mOpenMsgFile() Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imExporting = True
    'Print #hmMsg, "   Exported to: " & slToFile
    gAutomationAlertAndLogHandler "* Exported to: " & slToFile
    gAutomationAlertAndLogHandler "* AddedFromDate = " & edcAddedDate(0).Text
    gAutomationAlertAndLogHandler "* AddedToDate = " & edcAddedDate(1).Text
    gAutomationAlertAndLogHandler "* PurgedFromDate = " & edcPurgedDate(0).Text
    gAutomationAlertAndLogHandler "* PurgedToDate = " & edcPurgedDate(1).Text
    If ckcStatus(0).Value = vbChecked Then
        gAutomationAlertAndLogHandler "* CartType Active = True"
    Else
        gAutomationAlertAndLogHandler "* CartType Active = False"
    End If
    If ckcStatus(1).Value = vbChecked Then
        gAutomationAlertAndLogHandler "* CartType Purged = True"
    Else
        gAutomationAlertAndLogHandler "* CartType Purged = False"
    End If
    If ckcStatus(2).Value = vbChecked Then
        gAutomationAlertAndLogHandler "* CartType History = True"
    Else
        gAutomationAlertAndLogHandler "* CartType History = False"
    End If
    
    mExptCart
    Close hmTo
    'Print #hmMsg, "** Export Copy Completed: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** Export Copy Completed: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    Close #hmMsg
    'lacProcessing.Caption = "See: " & sgExportPath & "ExptCart.Txt" & " for Messages"
    lacProcessing.Caption = "See: " & sgDBPath & "Messages\" & "ExptCart.Txt" & " for Messages or use the Message Viewer."
    cmcCancel.Caption = "Done"
    cmcExport.Enabled = False
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
Private Sub cmcPurgedDate_Click(Index As Integer)
    plcCalendar.Visible = Not plcCalendar.Visible
    edcPurgedDate(Index).SelStart = 0
    edcPurgedDate(Index).SelLength = Len(edcPurgedDate(Index).Text)
    edcPurgedDate(Index).SetFocus
End Sub
Private Sub cmcPurgedDate_GotFocus(Index As Integer)
    Dim slStr As String
    If imDateBox <> Index + 2 Then
        plcCalendar.Visible = False
        slStr = edcPurgedDate(Index).Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    imDateBox = Index + 2
    plcCalendar.Move plcSelect.Left + edcPurgedDate(Index).Left, plcSelect.Top + edcPurgedDate(Index).Top + edcPurgedDate(Index).Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcAddedDate_Change(Index As Integer)
    Dim slStr As String
    slStr = edcAddedDate(Index).Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    mSetCommands
End Sub
Private Sub edcAddedDate_GotFocus(Index As Integer)
    If imDateBox <> Index Then
        plcCalendar.Visible = False
    End If
    plcCalendar.Move plcSelect.Left + edcAddedDate(Index).Left, plcSelect.Top + edcAddedDate(Index).Top + edcAddedDate(Index).Height
    imDateBox = Index
    gCtrlGotFocus edcAddedDate(Index)
    'slStr = edcAddedDate(Index).Text
    'If gValidDate(slStr) Then
    '    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    '    pbcCalendar_Paint   'mBoxCalDate called within paint
    'Else
    '    lacDate.Visible = False
    'End If
End Sub
Private Sub edcAddedDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcAddedDate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcAddedDate(Index).SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcAddedDate_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcAddedDate(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcAddedDate(Index).Text = slDate
            End If
        End If
        edcAddedDate(Index).SelStart = 0
        edcAddedDate(Index).SelLength = Len(edcAddedDate(Index).Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcAddedDate(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcAddedDate(Index).Text = slDate
            End If
        End If
        edcAddedDate(Index).SelStart = 0
        edcAddedDate(Index).SelLength = Len(edcAddedDate(Index).Text)
    End If
End Sub
Private Sub edcPurgedDate_Change(Index As Integer)
    Dim slStr As String
    slStr = edcPurgedDate(Index).Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    mSetCommands
End Sub
Private Sub edcPurgedDate_GotFocus(Index As Integer)
    If imDateBox <> Index + 2 Then
        plcCalendar.Visible = False
    End If
    plcCalendar.Move plcSelect.Left + edcPurgedDate(Index).Left, plcSelect.Top + edcPurgedDate(Index).Top + edcPurgedDate(Index).Height
    imDateBox = Index + 2
    gCtrlGotFocus edcPurgedDate(Index)
    'slStr = edcAddedDate(Index).Text
    'If gValidDate(slStr) Then
    '    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    '    pbcCalendar_Paint   'mBoxCalDate called within paint
    'Else
    '    lacDate.Visible = False
    'End If
End Sub
Private Sub edcPurgedDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcPurgedDate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcPurgedDate(Index).SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcPurgedDate_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcPurgedDate(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcPurgedDate(Index).Text = slDate
            End If
        End If
        edcPurgedDate(Index).SelStart = 0
        edcPurgedDate(Index).SelLength = Len(edcPurgedDate(Index).Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcPurgedDate(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcPurgedDate(Index).Text = slDate
            End If
        End If
        edcPurgedDate(Index).SelStart = 0
        edcPurgedDate(Index).SelLength = Len(edcPurgedDate(Index).Text)
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

Private Sub Form_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
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
    Erase tmUserVehicle
    Erase tmCartInfo
    
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmCuf)
    btrDestroy hmCuf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    '9887
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    Set ExptCart = Nothing   'Remove data segment
    
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
    If imDateBox <= 1 Then
        slStr = edcAddedDate(imDateBox).Text
    ElseIf imDateBox <= 3 Then
        slStr = edcPurgedDate(imDateBox - 2).Text
    End If
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
    Dim ilLoop As Integer
    
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imExporting = False
    imBypassFocus = False
    imAllClicked = False
    imSetAll = True
    smNowDate = Format$(gNow(), "mm/dd/yy")
    imSeqNo = 0
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptCart
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptCart
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptCart
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptCart
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptCart
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptCart
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmCuf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCuf, "", sgDBPath & "Cuf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptCart
    On Error GoTo 0
    imCufRecLen = Len(tmCuf)
    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptCart
    On Error GoTo 0
    imSifRecLen = Len(tmSif)
    'Populate arrays to determine if records exist
    'plcGauge.Move ExptCart.Width / 2 - plcGauge.Width / 2
    'cmcFileConv.Move ExptCart.Width / 2 - cmcFileConv.Width / 2
    'cmcCancel.Move ExptCart.Width / 2 - cmcCancel.Width / 2 - cmcCancel.Width
    'cmcReport.Move ExptCart.Width / 2 - cmcReport.Width / 2 + cmcReport.Width
    '9887
    hmCvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptCart
    On Error GoTo 0
    imCvfRecLen = Len(tmCvf)
        
    ilRet = gObtainSAF()
    smHideMediaCode = "N"
    For ilLoop = 0 To UBound(tgSaf) - 1 Step 1
        If tgSaf(ilLoop).iVefCode <= 0 Then
            If (Asc(tgSaf(0).sFeatures1) And ENGRHIDEMEDIACODE) = ENGRHIDEMEDIACODE Then
                smHideMediaCode = "Y"
            End If
            Exit For
        End If
    Next ilLoop

    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    mVehPop
    ckcAll.Value = vbChecked
    'edcAddedDate.Text = smNowDate
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    lacDate.Visible = False
    gCenterStdAlone ExptCart
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
Private Function mOpenMsgFile() As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    'On Error GoTo mOpenMsgFileErr:
    'slToFile = sgExportPath & "ExptCart.Txt"
    slToFile = sgDBPath & "Messages\" & "ExptCart.Txt"
    sgMessageFile = slToFile
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = gDateValue(smNowDate) Then  'Append
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
    
    'Print #hmMsg, "** Copy Export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** Copy Export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gExptCart                       *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the Cif records to be    *
'*                     exported and export             *
'*                                                     *
'*******************************************************
Private Sub mExptCart()
    Dim llDate As Long
    Dim ilCifOk As Integer
    Dim slAddedStartDate As String
    Dim slAddedEndDate As String
    Dim slPurgedStartDate As String
    Dim slPurgedEndDate As String
    Dim llAddedStartDate As Long
    Dim llAddedEndDate As Long
    Dim llPurgedStartDate As Long
    Dim llPurgedEndDate As Long
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slCart As String
    Dim slISCI As String
    Dim slShortTitle As String
    Dim slAdvtName As String
    Dim slAction As String
    Dim slRecord As String
    Dim ilPass As Integer
    Dim ilOffSet As Integer
    Dim ilCrf As Integer
    Dim ilVef As Integer
    Dim llLoop As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    '9887
    Dim ilVefCodesForCrf() As Integer
    Dim ilVefForCrf As Integer
    Dim ilCvf As Integer
    Dim ilCurrent As Integer
    
    On Error GoTo ErrHand
    imSeqNo = 0
    slAddedStartDate = edcAddedDate(0).Text
    If slAddedStartDate = "" Then
        llAddedStartDate = 0
        slAddedStartDate = "1/1/1970"
    Else
        llAddedStartDate = gDateValue(slAddedStartDate)
    End If
    slAddedEndDate = edcAddedDate(1).Text
    If slAddedEndDate = "" Then
        llAddedEndDate = 999999999
        slAddedEndDate = "12/31/2069"
    Else
        llAddedEndDate = gDateValue(slAddedEndDate)
    End If
    slPurgedStartDate = edcPurgedDate(0).Text
    If slPurgedStartDate = "" Then
        llPurgedStartDate = 0
        slPurgedStartDate = "1/1/1970"
    Else
        llPurgedStartDate = gDateValue(slPurgedStartDate)
    End If
    slPurgedEndDate = edcPurgedDate(1).Text
    If slPurgedEndDate = "" Then
        llPurgedEndDate = 999999999
        slPurgedEndDate = "12/31/2069"
    Else
        llPurgedEndDate = gDateValue(slPurgedEndDate)
    End If
    For ilPass = 1 To 2 Step 1
        ReDim tmCartInfo(0 To 0) As CARTINFO
        btrExtClear hmCif   'Clear any previous extend operation
        ilExtLen = Len(tmCif)
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
        btrExtClear hmCif   'Clear any previous extend operation
        If ilPass = 1 Then
            gPackDate slAddedStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilRet = btrGetGreaterOrEqual(hmCif, tmCif, imCifRecLen, tlDateTypeBuff, INDEXKEY5, BTRV_LOCK_NONE)   'Get first record as starting point
        Else
            gPackDate slPurgedStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilRet = btrGetGreaterOrEqual(hmCif, tmCif, imCifRecLen, tlDateTypeBuff, INDEXKEY6, BTRV_LOCK_NONE)   'Get first record as starting point
        End If
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            If ilRet <> BTRV_ERR_NONE Then
                Exit Sub
            End If
            Call btrExtSetBounds(hmCif, llNoRec, -1, "UC", "CIF", "") 'Set extract limits (all records)
            If ilPass = 1 Then
                gPackDate slAddedStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilOffSet = gFieldOffset("cif", "cifEntryDate")
            Else
                gPackDate slPurgedStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilOffSet = gFieldOffset("cif", "cifPurgeDate")
            End If
            ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            If ilPass = 1 Then
                gPackDate slAddedEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilOffSet = gFieldOffset("cif", "cifEntryDate")
            Else
                gPackDate slPurgedEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilOffSet = gFieldOffset("cif", "cifPurgeDate")
            End If
            ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            ilRet = btrExtAddField(hmCif, 0, ilExtLen)  'Extract Name
            If ilRet <> BTRV_ERR_NONE Then
                Exit Sub
            End If
            ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                    Exit Sub
                End If
                'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    ilCifOk = False
                    'gUnpackDateLong tmCif.iDateEntrd(0), tmCif.iDateEntrd(1), llDate
                    'If (llDate >= llAddedStartDate) And (llDate <= llAddedEndDate) Then
                    '    ilCifOk = True
                    '    slAction = "A"
                    'End If
                    'gUnpackDateLong tmCif.iPurgeDate(0), tmCif.iPurgeDate(1), llDate
                    'If (tmCif.sPurged = "P") And (llDate >= llPurgedStartDate) And (llDate <= llPurgedEndDate) Then
                    '    ilCifOk = True
                    '    slAction = "D"
                    'End If
                    If ckcStatus(0).Value = vbChecked Then
                        gUnpackDateLong tmCif.iDateEntrd(0), tmCif.iDateEntrd(1), llDate
                        If (tmCif.sPurged = "A") And (llDate >= llAddedStartDate) And (llDate <= llAddedEndDate) Then
                            ilCifOk = True
                            slAction = "A"
                        End If
                    End If
                    If ckcStatus(1).Value = vbChecked Then
                        gUnpackDateLong tmCif.iPurgeDate(0), tmCif.iPurgeDate(1), llDate
                        If (tmCif.sPurged = "P") And (llDate >= llPurgedStartDate) And (llDate <= llPurgedEndDate) Then
                            ilCifOk = True
                            slAction = "D"
                        End If
                    End If
                    If ckcStatus(2).Value = vbChecked Then
                        gUnpackDateLong tmCif.iDateEntrd(0), tmCif.iDateEntrd(1), llDate
                        If (tmCif.sPurged = "H") And (llDate >= llAddedStartDate) And (llDate <= llAddedEndDate) Then
                            ilCifOk = True
                            slAction = "A"
                        End If
                    End If
    
                    If ilCifOk Then
                        If ckcAll.Value = vbUnchecked Then
                            'Check vehicle selection
                            ilCifOk = False
                            tmCufSrchKey1.lCifCode = tmCif.lCode
                            ilRet = btrGetEqual(hmCuf, tmCuf, imCufRecLen, tmCufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            Do While (ilRet = BTRV_ERR_NONE) And (tmCuf.lCifCode = tmCif.lCode)
                                For ilCrf = LBound(tmCuf.lCrfCode) To UBound(tmCuf.lCrfCode) Step 1
                                    If tmCuf.lCrfCode(ilCrf) > 0 Then
                                        tmCrfSrchKey.lCode = tmCuf.lCrfCode(ilCrf)
                                        ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        If ilRet = BTRV_ERR_NONE Then
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
                                            For ilVefForCrf = 0 To UBound(ilVefCodesForCrf) - 1
                                                For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
                                                    If lbcVehicle.Selected(ilVef) Then
                                                        If ilVefCodesForCrf(ilVefForCrf) = Val(lbcVehicle.ItemData(ilVef)) Then
                                                            ilCifOk = True
                                                            Exit For
                                                        End If
                                                    End If
                                                Next ilVef
                                                If ilCifOk = True Then
                                                    Exit For
                                                End If
                                            Next ilVefForCrf
'                                            For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
'                                                If lbcVehicle.Selected(ilVef) Then
'                                                    If tmCrf.iVefCode = Val(lbcVehicle.ItemData(ilVef)) Then
'                                                        ilCifOk = True
'                                                        Exit For
'                                                    End If
'                                                End If
'                                            Next ilVef
                                        End If
                                    End If
                                    If ilCifOk Then
                                        Exit For
                                    End If
                                Next ilCrf
                                If ilCifOk Then
                                    Exit Do
                                End If
                                ilRet = btrGetNext(hmCuf, tmCuf, imCufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        End If
                    End If
                    
                    If ilCifOk Then
                        tmCpfSrchKey.lCode = tmCif.lcpfCode
                        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            '12/27/13: Use advertiser name instead of short title for WWO
                            'slShortTitle = gFileNameFilter(Trim$(tmCpf.sName))
                            slISCI = gFileNameFilter(Trim$(tmCpf.sISCI))
                            If tmMcf.iCode <> tmCif.iMcfCode Then
                                tmMcfSrchKey.iCode = tmCif.iMcfCode
                                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                    tmMcf.sName = ""
                                End If
                            End If
                            If tmAdf.iCode <> tmCif.iAdfCode Then
                                tmAdfSrchKey.iCode = tmCif.iAdfCode
                                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                    tmAdf.sName = ""
                                End If
                            End If
                            '12/27/13: Using advertiser instead of shorttitle for WWO
                            slShortTitle = gFileNameFilter(Trim$(tmAdf.sName))
                            slAdvtName = Trim$(tmAdf.sName)
                            '3/4/14: Suppress Media code
                            'slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                            'If (Len(Trim$(tmCif.sCut)) <> 0) Then
                            '    slCart = slCart & "-" & tmCif.sCut
                            'End If
                            '4/16/14: If Live, retain Media Code Name
                            'If (smHideMediaCode = "Y") Then
                            If (smHideMediaCode = "Y") And (Trim$(tmMcf.sName) <> "L") Then
                                slCart = Trim$(tmCif.sName)
                                If (Len(Trim$(tmCif.sCut)) <> 0) Then
                                    slCart = slCart & "-" & tmCif.sCut
                                End If
                            Else
                                slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                                If (Len(Trim$(tmCif.sCut)) <> 0) Then
                                    slCart = slCart & "-" & tmCif.sCut
                                End If
                            End If
                            'slRecord = """" & slCart & """"
                            'slRecord = slRecord & "," & slAction
                            'slRecord = slRecord & "," & """" & slShortTitle & """"
                            'slRecord = slRecord & "," & """" & slAdvtName & """"
                            'slRecord = slRecord & "," & Trim$(Str$(tmCif.iLen))
                            'slRecord = slRecord & "," & """" & slISCI & """"
                            slRecord = slCart
                            slRecord = slRecord & "," & slAction
                            slRecord = slRecord & "," & UCase(slShortTitle)
                            slRecord = slRecord & "," & UCase(gAdvtNameFilter(slAdvtName))
                            slRecord = slRecord & "," & Trim$(str$(tmCif.iLen))
                            slRecord = slRecord & "," & UCase(slISCI)
                            'Print #hmTo, slRecord
                            tmCartInfo(UBound(tmCartInfo)).sKey = slCart
                            tmCartInfo(UBound(tmCartInfo)).sRecord = slRecord
                            ReDim Preserve tmCartInfo(0 To UBound(tmCartInfo) + 1) As CARTINFO
                        Else
                            'Print #hmMsg, "   Unable to Read Product/ISCI: CIFCode" & str$(tmCif.lCode) & " CPFCode " & str$(tmCif.lcpfCode)
                            gAutomationAlertAndLogHandler "   Unable to Read Product/ISCI: CIFCode" & str$(tmCif.lCode) & " CPFCode " & str$(tmCif.lcpfCode)
                        End If
                    End If
                    ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        If UBound(tmCartInfo) > 1 Then
            ArraySortTyp fnAV(tmCartInfo(), 0), UBound(tmCartInfo), 0, LenB(tmCartInfo(0)), 0, LenB(tmCartInfo(0).sKey), 0
        End If
        For llLoop = 0 To UBound(tmCartInfo) - 1 Step 1
            Print #hmTo, Trim$(tmCartInfo(llLoop).sRecord)
        Next llLoop
    Next ilPass
    Erase ilVefCodesForCrf
    Exit Sub
ErrHand:
    ilRet = err.Number
    Resume Next
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
    Unload ExptCart
    igManUnload = NO
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
                If imDateBox <= 1 Then
                    edcAddedDate(imDateBox).Text = Format$(llDate, "m/d/yy")
                    edcAddedDate(imDateBox).SelStart = 0
                    edcAddedDate(imDateBox).SelLength = Len(edcAddedDate(imDateBox).Text)
                    imBypassFocus = True
                    edcAddedDate(imDateBox).SetFocus
                    Exit Sub
                ElseIf imDateBox <= 3 Then
                    edcPurgedDate(imDateBox - 2).Text = Format$(llDate, "m/d/yy")
                    edcPurgedDate(imDateBox - 2).SelStart = 0
                    edcPurgedDate(imDateBox - 2).SelLength = Len(edcPurgedDate(imDateBox - 2).Text)
                    imBypassFocus = True
                    edcPurgedDate(imDateBox - 2).SetFocus
                    Exit Sub
                End If
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If imDateBox <= 1 Then
        edcAddedDate(imDateBox).SetFocus
    ElseIf imDateBox <= 3 Then
        edcPurgedDate(imDateBox - 2).SetFocus
    End If
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
    plcScreen.Print "Export Carts"
End Sub
Private Sub mVehPop()
    Dim ilRet As Integer
    
    On Error GoTo mVehPopErr
    ilRet = btrGetFirst(hmVef, tmVef, imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If (tmVef.sType <> "L") And (tmVef.sType <> "R") And (tmVef.sType <> "N") Then
            lbcVehicle.AddItem Trim$(tmVef.sName)
            lbcVehicle.ItemData(lbcVehicle.NewIndex) = tmVef.iCode
        End If
        ilRet = btrGetNext(hmVef, tmVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
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
Private Sub mSetCommands()
    Dim ilEnabled As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    
    ilEnabled = False
    'at least one vehicle must be selected
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            ilEnabled = True
            Exit For
        End If
    Next ilLoop

    If ilEnabled Then
        slStr = edcAddedDate(0).Text
        If slStr <> "" Then
            If Not gValidDate(slStr) Then
                ilEnabled = False
            End If
        End If
    End If
    If ilEnabled Then
        slStr = edcAddedDate(1).Text
        If slStr <> "" Then
            If Not gValidDate(slStr) Then
                ilEnabled = False
            End If
        End If
    End If

    If ilEnabled Then
        slStr = edcPurgedDate(0).Text
        If slStr <> "" Then
            If Not gValidDate(slStr) Then
                ilEnabled = False
            End If
        End If
    End If
    If ilEnabled Then
        slStr = edcPurgedDate(1).Text
        If slStr <> "" Then
            If Not gValidDate(slStr) Then
                ilEnabled = False
            End If
        End If
    End If
    If (ckcStatus(0).Value = vbUnchecked) And (ckcStatus(1).Value = vbUnchecked) And (ckcStatus(2).Value = vbUnchecked) Then
        ilEnabled = False
    End If
    cmcExport.Enabled = ilEnabled
End Sub
