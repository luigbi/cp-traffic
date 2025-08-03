VERSION 5.00
Begin VB.Form RCReallo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3675
   ClientLeft      =   1515
   ClientTop       =   1590
   ClientWidth     =   6450
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   6450
   Begin VB.PictureBox plcCalendar 
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
      Height          =   1770
      Left            =   4095
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   18
      Top             =   1755
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Rcreallo.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   22
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
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
         Top             =   45
         Width           =   285
      End
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
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
         Height          =   210
         Left            =   330
         TabIndex        =   23
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.ListBox lbcAud 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2790
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2265
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3450
      TabIndex        =   25
      Top             =   3255
      Width           =   1050
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   165
      Left            =   90
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3945
      Width           =   120
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   30
      Width           =   2715
   End
   Begin VB.CommandButton cmcRealloc 
      Appearance      =   0  'Flat
      Caption         =   "&Reallocate"
      Height          =   285
      Left            =   2130
      TabIndex        =   17
      Top             =   3255
      Width           =   1050
   End
   Begin VB.PictureBox plcRealloc 
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
      Height          =   2505
      Left            =   105
      ScaleHeight     =   2445
      ScaleWidth      =   6120
      TabIndex        =   1
      Top             =   390
      Width           =   6180
      Begin VB.Frame frcAud 
         Caption         =   "Audience Source"
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
         Height          =   1125
         Left            =   120
         TabIndex        =   9
         Top             =   1245
         Width           =   4905
         Begin VB.TextBox edcAud 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   675
            Width           =   780
         End
         Begin VB.CommandButton cmcAud 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2220
            Picture         =   "Rcreallo.frx":2E1A
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   675
            Width           =   195
         End
         Begin VB.TextBox edcAud 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   3480
            MaxLength       =   15
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   330
            Width           =   780
         End
         Begin VB.CommandButton cmcAud 
            Appearance      =   0  'Flat
            Caption         =   "t"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   4260
            Picture         =   "Rcreallo.frx":2F14
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   330
            Width           =   195
         End
         Begin VB.OptionButton rbcAud 
            Caption         =   "Always Use"
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
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   675
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton rbcAud 
            Caption         =   "Client Target      If not defined, use"
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
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   330
            Width           =   3345
         End
      End
      Begin VB.CheckBox ckcChg 
         Caption         =   "Orders"
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
         Height          =   270
         Index           =   1
         Left            =   2010
         TabIndex        =   8
         Top             =   915
         Value           =   1  'Checked
         Width           =   1170
      End
      Begin VB.CheckBox ckcChg 
         Caption         =   "Proposals"
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
         Height          =   270
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   915
         Value           =   1  'Checked
         Width           =   1170
      End
      Begin VB.TextBox edcDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3135
         MaxLength       =   10
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   615
         Width           =   930
      End
      Begin VB.CommandButton cmcDate 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4065
         Picture         =   "Rcreallo.frx":300E
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   615
         Width           =   195
      End
      Begin VB.Label lacChg 
         Appearance      =   0  'Flat
         Caption         =   "Change"
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
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   945
         Width           =   660
      End
      Begin VB.Label lacMsg 
         Appearance      =   0  'Flat
         Caption         =   "This will reallocate prices for audience-based packages using latest research books."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   1050
         TabIndex        =   2
         Top             =   75
         Width           =   3960
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Earliest Date to Reallocate Dollars"
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
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   630
         Width           =   3135
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Schedule all Contracts prior to running Reallocate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   26
      Top             =   2955
      Width           =   6135
   End
End
Attribute VB_Name = "RCReallo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rcreallo.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RCReallo.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract model screen code
Option Explicit
Option Compare Text
'****Comment out all the code to eliminate an Out of Memory when compiling Traffic as this module sis not used
''Contract
'Dim hmChf As Integer        'Contract file handle
'Dim tmChf As CHF            'CHF record image
'Dim tmChfSrchKey As LONGKEY0 'CHF key record image
'Dim imChfRecLen As Integer  'CHF record length
''Line
'Dim hmClf As Integer        'Line file handle
'Dim tmClf As CLF            'CLF record image
'Dim imClfRecLen As Integer  'CLF record length
''Flight
'Dim hmCff As Integer        'Flight file handle
'Dim tmCff As CFF            'CFF record image
'Dim imCffRecLen As Integer  'CFF record length
'Dim imLastCffUsed As Integer
'Dim tmCffAud() As CFFAUD
'Dim lmCffRecPos() As Long
'Dim tmChfAdvtExt() As CHFADVTEXT
''Multi-Name
'Dim hmMnf As Integer        'Multi-Name file handle
'Dim tmMnf As MNF            'MNF record image
'Dim imMnfRecLen As Integer  'MNF record length
''Contract
'Dim hmDrf As Integer        'Research Data file handle
'Dim tmDrf As DRF
'Dim imDrfRecLen As Integer  'DRF record length
''Site Option
'Dim hmSpf As Integer        'Site Option file handle
'Dim tmSpf As SPF            'SPF record image
'Dim imSpfRecLen As Integer  'SPF record length
''Demo Plus
'Dim hmDpf As Integer
''Research Estimate
'Dim hmDef As Integer
'Dim hmRaf As Integer
'
'Dim smSyncDate As String
'Dim smSyncTime As String
'Dim hmMsg As Integer
'Dim lmNowDate As Long   'Todays date
'Dim smStartDate As String
'Dim imAudIndex As Integer
'Dim lmTotalOld As Long
'Dim lmTotalNew As Long
'Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
'Dim imBSMode As Integer     'Backspace flag
'Dim imFirstActivate As Integer
'Dim imTerminate As Integer  'True = terminating task, False= OK
'Dim imBypassFocus As Integer
'Dim imLbcArrowSetting As Integer
'Dim imComboBoxIndex As Integer
'Dim imLbcMouseDown As Integer
'Dim imUpdateAllowed As Integer
'
''Calendar
'Dim tmCDCtrls(0 To 7) As FIELDAREA
'Dim imLBCDCtrls As Integer
'Dim imCalYear As Integer    'Month of displayed calendar
'Dim imCalMonth As Integer   'Year of displayed calendar
'Dim lmCalStartDate As Long  'Start date of displayed calendar
'Dim lmCalEndDate As Long    'End date of displayed calendar
'Dim imCalType As Integer
''Dim tmRec As LPOPREC
'Private Sub ckcChg_GotFocus(Index As Integer)
'    plcCalendar.Visible = False
'    lbcAud.Visible = False
'End Sub
'Private Sub cmcAud_Click(Index As Integer)
'    lbcAud.Visible = Not lbcAud.Visible
'    edcAud(Index).SelStart = 0
'    edcAud(Index).SelLength = Len(edcAud(Index).Text)
'    edcAud(Index).SetFocus
'End Sub
'Private Sub cmcAud_GotFocus(Index As Integer)
'    Dim slStr As String
'    If imAudIndex <> Index Then
'        imAudIndex = Index
'        lbcAud.Visible = False
'        slStr = edcAud(Index).Text
'        gFindMatch slStr, 0, lbcAud
'        If gLastFound(lbcAud) >= 0 Then
'            lbcAud.ListIndex = gLastFound(lbcAud)
'        End If
'    End If
'    plcCalendar.Visible = False
'    lbcAud.Move plcRealloc.Left + frcAud.Left + edcAud(Index).Left, plcRealloc.Top + frcAud.Top + edcAud(Index).Top - lbcAud.Height '+ edcAud(Index).Height
'    gCtrlGotFocus ActiveControl
'End Sub
'Private Sub cmcCalDn_Click()
'    imCalMonth = imCalMonth - 1
'    If imCalMonth <= 0 Then
'        imCalMonth = 12
'        imCalYear = imCalYear - 1
'    End If
'    pbcCalendar_Paint
'    edcDate.SelStart = 0
'    edcDate.SelLength = Len(edcDate.Text)
'    edcDate.SetFocus
'End Sub
'Private Sub cmcCalUp_Click()
'    imCalMonth = imCalMonth + 1
'    If imCalMonth > 12 Then
'        imCalMonth = 1
'        imCalYear = imCalYear + 1
'    End If
'    pbcCalendar_Paint
'    edcDate.SelStart = 0
'    edcDate.SelLength = Len(edcDate.Text)
'    edcDate.SetFocus
'End Sub
'Private Sub cmcCancel_Click()
'    igTerminateReturn = 0
'    mTerminate
'End Sub
'Private Sub cmcCancel_GotFocus()
'    plcCalendar.Visible = False
'    lbcAud.Visible = False
'    gCtrlGotFocus cmcCancel
'End Sub
'Private Sub cmcDate_Click()
'    plcCalendar.Visible = Not plcCalendar.Visible
'    edcDate.SelStart = 0
'    edcDate.SelLength = Len(edcDate.Text)
'    edcDate.SetFocus
'End Sub
'Private Sub cmcDate_GotFocus()
'    gCtrlGotFocus ActiveControl
'    lbcAud.Visible = False
'End Sub
'Private Sub cmcRealloc_Click()
'    Dim slDate As String
'    Dim slStr As String
'    Dim ilRet As Integer
'    Dim llStdDate As Long
'    Dim llCalDate As Long
'    Dim slPrtCntr As String * 8
'    Dim slPrtName As String * 55
'    Dim slPrtLn As String * 4
'    Dim slPrtVeh As String * 20
'    Dim slPrtDate As String * 8
'    Dim slPrtOld As String * 12
'    Dim slPrtNew As String * 12
'    If Not imUpdateAllowed Then
'        Exit Sub
'    End If
'    slDate = edcDate.Text
'    If Not gValidDate(slDate) Then
'        Beep
'        edcDate.SetFocus
'        Exit Sub
'    End If
'    If gWeekDayStr(slDate) <> 0 Then
'        Beep
'        ilRet = MsgBox("Start Date must be a Monday", vbOkOnly + vbExclamation, "Error")
'        edcDate.SetFocus
'        Exit Sub
'    End If
'    gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llStdDate
'    gUnpackDateLong tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), llCalDate
'    If gDateValue(slDate) <= llStdDate Then
'        Beep
'        ilRet = MsgBox("Start Date must be after Last Billed Date " & Format$(llStdDate, "m/d/yy"), vbOkOnly + vbExclamation, "Error")
'        edcDate.SetFocus
'        Exit Sub
'    End If
'    If gDateValue(slDate) <= llCalDate Then
'        Beep
'        ilRet = MsgBox("Start Date must be after Last Billed Date " & Format$(llCalDate, "m/d/yy"), vbOkOnly + vbExclamation, "Error")
'        edcDate.SetFocus
'        Exit Sub
'    End If
'    Screen.MousePointer = vbHourglass
'    If Not mOpenMsgFile() Then
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    mRealloc slDate
'    mUpdateSpf slDate
'    Print #hmMsg, ""
'    slStr = gLongToStrDec(lmTotalOld, 2)
'    gFormatStr slStr, FMTCOMMA, 2, slStr
'    Do While Len(slStr) < Len(slPrtOld)
'        slStr = " " & slStr
'    Loop
'    slPrtOld = slStr
'    slStr = gLongToStrDec(lmTotalNew, 2)
'    gFormatStr slStr, FMTCOMMA, 2, slStr
'    Do While Len(slStr) < Len(slPrtNew)
'        slStr = " " & slStr
'    Loop
'    slPrtNew = slStr
'    slPrtCntr = ""
'    slPrtName = "Grand Total"
'    slPrtLn = ""
'    slPrtVeh = ""
'    slPrtDate = ""
'    Print #hmMsg, slPrtCntr; Spc(1); slPrtName; Spc(1); slPrtLn; Spc(1); slPrtVeh; Spc(1); slPrtDate; Spc(1); slPrtOld; Spc(1); slPrtNew
'    Print #hmMsg, ""
'    Print #hmMsg, "** Completed Reallocation of Dollars: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
'    Close #hmMsg
'    Screen.MousePointer = vbDefault
'    ilRet = MsgBox("Results Stored into : " & sgDBPath & "Messages\" & "Realloc.Txt", vbOkOnly + vbExclamation, "Completed")
'    mTerminate
'End Sub
'Private Sub cmcRealloc_GotFocus()
'    plcCalendar.Visible = False
'    lbcAud.Visible = False
'    gCtrlGotFocus cmcRealloc
'End Sub
'Private Sub edcAud_Change(Index As Integer)
'    imLbcArrowSetting = True
'    gMatchLookAhead edcAud(Index), lbcAud, imBSMode, imComboBoxIndex
'    imLbcArrowSetting = False
'End Sub
'Private Sub edcAud_GotFocus(Index As Integer)
'    Dim slStr As String
'    If imAudIndex <> Index Then
'        imAudIndex = Index
'        lbcAud.Visible = False
'        slStr = edcAud(Index).Text
'        gFindMatch slStr, 0, lbcAud
'        If gLastFound(lbcAud) >= 0 Then
'            lbcAud.ListIndex = gLastFound(lbcAud)
'        End If
'    End If
'    plcCalendar.Visible = False
'    lbcAud.Move plcRealloc.Left + frcAud.Left + edcAud(Index).Left, plcRealloc.Top + frcAud.Top + edcAud(Index).Top - lbcAud.Height '+edcAud(Index).Height
'    If Not imBypassFocus Then
'        gCtrlGotFocus ActiveControl
'        imComboBoxIndex = lbcAud.ListIndex
'    End If
'    imBypassFocus = False
'End Sub
'Private Sub edcAud_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    'Delete key causes the charact to the right of the cursor to be deleted
'    imBSMode = False
'End Sub
'Private Sub edcAud_KeyPress(Index As Integer, KeyAscii As Integer)
'    Dim ilKey As Integer
'    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
'        If edcAud(Index).SelLength <> 0 Then    'avoid deleting two characters
'            imBSMode = True 'Force deletion of character prior to selected text
'        End If
'    End If
'    ilKey = KeyAscii
'    If Not gCheckKeyAscii(ilKey) Then
'        KeyAscii = 0
'        Exit Sub
'    End If
'End Sub
'Private Sub edcAud_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
'        gProcessArrowKey Shift, KeyCode, lbcAud, imLbcArrowSetting
'        edcAud(Index).SelStart = 0
'        edcAud(Index).SelLength = Len(edcAud(Index).Text)
'    End If
'End Sub
'Private Sub edcDate_Change()
'    Dim slStr As String
'    slStr = edcDate.Text
'    If Not gValidDate(slStr) Then
'        Exit Sub
'    End If
'    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
'    pbcCalendar_Paint   'mBoxCalDate called within paint
'End Sub
'Private Sub edcDate_GotFocus()
'    lbcAud.Visible = False
'    If Not imBypassFocus Then
'        gCtrlGotFocus ActiveControl
'    End If
'    imBypassFocus = False
'End Sub
'Private Sub edcDate_KeyDown(KeyCode As Integer, Shift As Integer)
'    'Delete key causes the charact to the right of the cursor to be deleted
'    imBSMode = False
'End Sub
'Private Sub edcDate_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
'        If edcDate.SelLength <> 0 Then    'avoid deleting two characters
'            imBSMode = True 'Force deletion of character prior to selected text
'        End If
'    End If
'    'Filter characters (allow only BackSpace, numbers 0 thru 9
'    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
'        Beep
'        KeyAscii = 0
'        Exit Sub
'    End If
'End Sub
'Private Sub edcDate_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim slDate As String
'    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
'        If (Shift And vbAltMask) > 0 Then
'            plcCalendar.Visible = Not plcCalendar.Visible
'        Else
'            slDate = edcDate.Text
'            If gValidDate(slDate) Then
'                If KeyCode = KEYUP Then 'Up arrow
'                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
'                Else
'                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
'                End If
'                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
'                edcDate.Text = slDate
'            End If
'        End If
'        edcDate.SelStart = 0
'        edcDate.SelLength = Len(edcDate.Text)
'    End If
'    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
'        If (Shift And vbAltMask) > 0 Then
'        Else
'            slDate = edcDate.Text
'            If gValidDate(slDate) Then
'                If KeyCode = KEYLEFT Then 'Up arrow
'                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
'                Else
'                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
'                End If
'                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
'                edcDate.Text = slDate
'            End If
'        End If
'        edcDate.SelStart = 0
'        edcDate.SelLength = Len(edcDate.Text)
'    End If
'End Sub
'
'Private Sub Form_Activate()
'    If Not imFirstActivate Then
'        DoEvents    'Process events so pending keys are not sent to this
'                    'form when keypreview turn on
'        Me.KeyPreview = True
'        Exit Sub
'    End If
'    imFirstActivate = False
'    If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
'        imUpdateAllowed = False
'    Else
'        imUpdateAllowed = True
'    End If
'    Me.KeyPreview = True
'    Me.Refresh
'End Sub
'
'Private Sub Form_Click()
'    pbcClickFocus.SetFocus
'End Sub
'
'Private Sub Form_Deactivate()
'    Me.KeyPreview = False
'End Sub
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
'        plcCalendar.Visible = False
'        gFunctionKeyBranch KeyCode
'    End If
'End Sub
'
Private Sub Form_Load()
'    mInit
End Sub
'Private Sub lbcAud_Click()
'    If imLbcMouseDown Then
'        imLbcArrowSetting = False
'    Else
'        imLbcArrowSetting = True
'    End If
'    gProcessLbcClick lbcAud, edcAud(imAudIndex), imChgMode, imLbcArrowSetting
'    imLbcMouseDown = False
'End Sub
'Private Sub lbcAud_GotFocus()
'    gCtrlGotFocus ActiveControl
'End Sub
'Private Sub lbcAud_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    imLbcMouseDown = True
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mAddWeek                        *
''*                                                     *
''*             Created:8/25/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Add week into flights          *
''*                                                     *
''*            Note: This code taken from Contract.Bas  *
''*                                                     *
''*******************************************************
'Private Sub mAddWeek(ilLastCffUsed As Integer, ilLnRowNo As Integer, llDate As Long, ilNoSpots As Integer, ilAllowedDays() As Integer)
''
''   ilLnRowNo(I)- Line row number
''   llDate(I)- Date within week to add
''   ilNoSpots(I)- Number of spots
''   ilAllowedDays(I)- Allowed days (True= Air Day; False=Not Air Day)
''
'
'    Dim ilCff As Integer
'    Dim ilPrevCff As Integer
'    Dim ilLoop As Integer
'    Dim llFlStartDate As Long
'    Dim llFlEndDate As Long
'    Dim llFlMoStartDate As Long
'    Dim llFlSuEndDate As Long
'    Dim llFlFirstMoDate As Long
'    Dim llFlLastMoDate As Long
'    Dim slStartDate As String
'    Dim slEndDate As String
'    Dim ilDay As Integer
'    Dim ilCffIndex As Integer
'    Dim ilAdd As Integer
'    Dim llMoDate As Long
'    Dim llSuDate As Long
'    Dim slDate As String
'    Dim ilAddSplitCff As Integer
'    Dim ilLIndex As Integer
'    Dim ilClf As Integer
'    Dim ilTestCff As Integer
'    Dim ilFound As Integer
'    Dim tlCff As CFFLIST
'    ilAdd = 0
'    ilPrevCff = -1
'    ilAddSplitCff = False
'    llMoDate = llDate
'    llSuDate = llDate
'    ilLIndex = LBound(ilAllowedDays)
'    Do While gWeekDayLong(llMoDate) <> 0
'        llMoDate = llMoDate - 1
'    Loop
'    Do While gWeekDayLong(llSuDate) <> 6
'        llSuDate = llSuDate + 1
'    Loop
'    ilCff = tgClfRC(ilLnRowNo - 1).iFirstCff
'    Do While ilCff <> -1
'        If tgClfRC(ilLnRowNo - 1).iCancel Then
'            tgClfRC(ilLnRowNo - 1).iCancel = False
'            tgClfRC(ilLnRowNo - 1).ClfRec.sHideCBS = "N"
'            slStartDate = gObtainPrevMonday(Format$(llDate, "m/d/yy"))
'            slEndDate = gObtainNextSunday(Format$(llDate, "m/d/yy"))
'            gPackDate slStartDate, tgCffRC(ilCff).CffRec.iStartDate(0), tgCffRC(ilCff).CffRec.iStartDate(1)
'            tgCffRC(ilCff).lStartDate = gDateValue(slStartDate)
'            gPackDate slEndDate, tgCffRC(ilCff).CffRec.iEndDate(0), tgCffRC(ilCff).CffRec.iEndDate(1)
'            tgCffRC(ilCff).lEndDate = gDateValue(slEndDate)
'            For ilDay = 0 To 6 Step 1
'                If ilAllowedDays(ilDay + ilLIndex) Then
'                    tgCffRC(ilCff).CffRec.iDay(ilDay) = 1
'                Else
'                    tgCffRC(ilCff).CffRec.iDay(ilDay) = 0
'                End If
'            Next ilDay
'            tgCffRC(ilCff).CffRec.sDyWk = "W"
'            tgCffRC(ilCff).CffRec.iSpotsWk = ilNoSpots
'            Exit Sub
'        Else
'            If (tgCffRC(ilCff).iStatus = 0) Or (tgCffRC(ilCff).iStatus = 1) Then
'            'gUnpackDateLong tgCffRC(ilCff).CffRec.iStartDate(0), tgCffRC(ilCff).CffRec.iStartDate(1), llFlStartDate    'Week Start date
'            'gUnpackDateLong tgCffRC(ilCff).CffRec.iEndDate(0), tgCffRC(ilCff).CffRec.iEndDate(1), llFlEndDate    'Week Start date
'                llFlStartDate = tgCffRC(ilCff).lStartDate
'                llFlEndDate = tgCffRC(ilCff).lEndDate
'                llFlMoStartDate = llFlStartDate
'                Do While gWeekDayLong(llFlMoStartDate) <> 0
'                    llFlMoStartDate = llFlMoStartDate - 1
'                Loop
'                llFlSuEndDate = llFlEndDate
'                Do While gWeekDayLong(llFlSuEndDate) <> 6
'                    llFlSuEndDate = llFlSuEndDate + 1
'                Loop
'                If (llDate >= llFlMoStartDate) And (llDate <= llFlSuEndDate) Then
'                    If (llMoDate <= llFlStartDate) And (llSuDate >= llFlEndDate) Then
'                        'slStartDate = gObtainPrevMonday(Format$(llDate, "m/d/yy"))
'                        'slEndDate = gObtainNextSunday(Format$(llDate, "m/d/yy"))
'                        'gPackDate slStartDate, tgCffRC(ilCff).CffRec.iStartDate(0), tgCffRC(ilCff).CffRec.iStartDate(1)
'                        'gPackDate slEndDate, tgCffRC(ilCff).CffRec.iEndDate(0), tgCffRC(ilCff).CffRec.iEndDate(1)
'                        For ilDay = 0 To 6 Step 1
'                            If ilAllowedDays(ilDay + ilLIndex) Then
'                                tgCffRC(ilCff).CffRec.iDay(ilDay) = 1
'                            Else
'                                tgCffRC(ilCff).CffRec.iDay(ilDay) = 0
'                            End If
'                        Next ilDay
'                        tgCffRC(ilCff).CffRec.sDyWk = "W"
'                        tgCffRC(ilCff).CffRec.iSpotsWk = ilNoSpots
'                        Exit Sub
'                    End If
'                    'Split flight
'                    ilAdd = 1
'                    Exit Do
'                End If
'                If llDate < llFlStartDate Then
'                    'Add prior to current flight
'                    ilAdd = 2
'                    Exit Do
'                End If
'            End If
'            ilPrevCff = ilCff
'            ilCff = tgCffRC(ilCff).iNextCff
'        End If
'    Loop
'    'Add to end of the flights
'    GoSub lObtainNextCff
'    tgCffRC(ilCffIndex).iNextCff = -1
'    tgCffRC(ilCffIndex).iStatus = 0
'    tgCffRC(ilCffIndex).lRecPos = 0
'    If ilAdd = 0 Then
'        If tgClfRC(ilLnRowNo - 1).iFirstCff = -1 Then
'            tgClfRC(ilLnRowNo - 1).iFirstCff = ilCffIndex
'        Else
'            tgCffRC(ilPrevCff).iNextCff = ilCffIndex
'        End If
'    ElseIf ilAdd = 1 Then
'        llFlFirstMoDate = llFlStartDate
'        Do While gWeekDayLong(llFlFirstMoDate) <> 0
'            llFlFirstMoDate = llFlFirstMoDate - 1
'        Loop
'        llFlLastMoDate = llFlEndDate
'        Do While gWeekDayLong(llFlLastMoDate) <> 0
'            llFlLastMoDate = llFlLastMoDate - 1
'        Loop
'        tlCff = tgCffRC(ilCff)
'        If llMoDate = llFlFirstMoDate Then
'            'Replace first week of flight
'            If ilPrevCff = -1 Then
'                tgCffRC(ilCffIndex).iNextCff = tgClfRC(ilLnRowNo - 1).iFirstCff
'                tgClfRC(ilLnRowNo - 1).iFirstCff = ilCffIndex
'            Else
'                tgCffRC(ilCffIndex).iNextCff = tgCffRC(ilPrevCff).iNextCff
'                tgCffRC(ilPrevCff).iNextCff = ilCffIndex
'            End If
'            'gUnpackDateLong tgCffRC(ilCff).CffRec.iStartDate(0), tgCffRC(ilCff).CffRec.iStartDate(1), llFlStartDate    'Week Start date
'            llFlStartDate = tgCffRC(ilCff).lStartDate
'            Do
'                llFlStartDate = llFlStartDate + 1
'            Loop Until gWeekDayLong(llFlStartDate) = 0
'            slDate = Format$(llFlStartDate, "m/d/yy")
'            gPackDate slDate, tgCffRC(ilCff).CffRec.iStartDate(0), tgCffRC(ilCff).CffRec.iStartDate(1)    'Week Start date
'            tgCffRC(ilCff).lStartDate = llFlStartDate
'        ElseIf llMoDate = llFlLastMoDate Then
'            'Replace last week of flight
'            tgCffRC(ilCffIndex).iNextCff = tgCffRC(ilCff).iNextCff
'            tgCffRC(ilCff).iNextCff = ilCffIndex
'            'gUnpackDateLong tgCffRC(ilCff).CffRec.iEndDate(0), tgCffRC(ilCff).CffRec.iEndDate(1), llFlEndDate    'Week Start date
'            llFlEndDate = tgCffRC(ilCff).lEndDate
'            Do
'                llFlEndDate = llFlEndDate - 1
'            Loop Until gWeekDayLong(llFlEndDate) = 6
'            slDate = Format$(llFlEndDate, "m/d/yy")
'            gPackDate slDate, tgCffRC(ilCff).CffRec.iEndDate(0), tgCffRC(ilCff).CffRec.iEndDate(1)    'Week Start date
'            tgCffRC(ilCff).lEndDate = llFlEndDate
'        Else
'            'Split flight
'            'tlCff = tgCffRC(ilCff)
'            tlCff.iStatus = 0
'            tlCff.lRecPos = 0
'            tgCffRC(ilCff).iNextCff = ilCffIndex
'            'tgCffRC(ilCffIndex).iNextCff = UBound(tgCffRC) + 1
'            slDate = Format$(llMoDate - 1, "m/d/yy")
'            gPackDate slDate, tgCffRC(ilCff).CffRec.iEndDate(0), tgCffRC(ilCff).CffRec.iEndDate(1)    'Week Start date
'            tgCffRC(ilCff).lEndDate = llMoDate - 1
'            slDate = Format$(llSuDate + 1, "m/d/yy")
'            gPackDate slDate, tlCff.CffRec.iStartDate(0), tlCff.CffRec.iStartDate(1)    'Week Start date
'            tlCff.lStartDate = llSuDate + 1
'            ilAddSplitCff = True
'        End If
'    Else
'        If ilPrevCff = -1 Then
'            tgCffRC(ilCffIndex).iNextCff = tgClfRC(ilLnRowNo - 1).iFirstCff
'            tgClfRC(ilLnRowNo - 1).iFirstCff = ilCffIndex
'        Else
'            tgCffRC(ilCffIndex).iNextCff = tgCffRC(ilPrevCff).iNextCff
'            tgCffRC(ilPrevCff).iNextCff = ilCffIndex
'        End If
'    End If
'    'ReDim Preserve tgCffRC(0 To ilCffIndex + 1) As CFFLIST
'    'tgCffRC(ilCffIndex + 1).iStatus = -1 'Not Used
'    'tgCffRC(ilCffIndex + 1).lRecPos = 0
'    'tgCffRC(ilCffIndex + 1).iNextCff = -1
'    tgCffRC(ilCffIndex).iStatus = 0   'New to not used
'    tgCffRC(ilCffIndex).CffRec.lChfCode = tgChfRC.lCode
'    tgCffRC(ilCffIndex).CffRec.iClfLine = tgClfRC(ilLnRowNo - 1).ClfRec.iLine
'    tgCffRC(ilCffIndex).CffRec.iCntRevNo = tgClfRC(ilLnRowNo - 1).ClfRec.iCntRevNo
'    tgCffRC(ilCffIndex).CffRec.iPropVer = tgClfRC(ilLnRowNo - 1).ClfRec.iPropVer
'    slStartDate = gObtainPrevMonday(Format$(llDate, "m/d/yy"))
'    slEndDate = gObtainNextSunday(Format$(llDate, "m/d/yy"))
'    gPackDate slStartDate, tgCffRC(ilCffIndex).CffRec.iStartDate(0), tgCffRC(ilCffIndex).CffRec.iStartDate(1)
'    tgCffRC(ilCffIndex).lStartDate = gDateValue(slStartDate)
'    gPackDate slEndDate, tgCffRC(ilCffIndex).CffRec.iEndDate(0), tgCffRC(ilCffIndex).CffRec.iEndDate(1)
'    tgCffRC(ilCffIndex).lEndDate = gDateValue(slEndDate)
'    tgCffRC(ilCffIndex).CffRec.sDyWk = "W"
'    tgCffRC(ilCffIndex).CffRec.iSpotsWk = ilNoSpots
'    For ilDay = 0 To 6 Step 1
'        If ilAllowedDays(ilDay + ilLIndex) Then
'            tgCffRC(ilCffIndex).CffRec.iDay(ilDay) = 1
'        Else
'            tgCffRC(ilCffIndex).CffRec.iDay(ilDay) = 0
'        End If
'        tgCffRC(ilCffIndex).CffRec.sXDay(ilDay) = "0"
'    Next ilDay
'    tgCffRC(ilCffIndex).CffRec.sDelete = "N"
'    tgCffRC(ilCffIndex).CffRec.iXSpotsWk = 0
'    If ilAdd <> 1 Then
'        tgCffRC(ilCffIndex).CffRec.sPriceType = "*"   '* used to indicate price needs to be set
'        tgCffRC(ilCffIndex).CffRec.lActPrice = 0  'Later- might want to store average package price
'        tgCffRC(ilCffIndex).CffRec.lPropPrice = 0
'    Else
'        tgCffRC(ilCffIndex).CffRec.sPriceType = tlCff.CffRec.sPriceType
'        tgCffRC(ilCffIndex).CffRec.lActPrice = tlCff.CffRec.lActPrice
'        tgCffRC(ilCffIndex).CffRec.lPropPrice = tlCff.CffRec.lPropPrice
'    End If
'    If ilAddSplitCff Then
'        'ReDim Preserve tgCffRC(0 To UBound(tgCffRC) + 1) As CFFLIST
'        'tgCffRC(UBound(tgCffRC)).iStatus = -1 'Not Used
'        'tgCffRC(UBound(tgCffRC)).lRecPos = 0
'        'tgCffRC(UBound(tgCffRC)).iNextCff = -1
'        'tgCffRC(UBound(tgCffRC) - 1) = tlCff
'        ilCff = ilCffIndex
'        GoSub lObtainNextCff
'        tgCffRC(ilCff).iNextCff = ilCffIndex
'        tgCffRC(ilCffIndex) = tlCff
'    End If
'    Exit Sub
'lObtainNextCff:
'    ilCffIndex = -1
'    For ilLoop = ilLastCffUsed To UBound(tgCffRC) - 1 Step 1
'        If tgCffRC(ilLoop).iStatus = -1 Then
'            'Test Chain
'            ilFound = False
'            For ilClf = LBound(tgClfRC) To UBound(tgClfRC) Step 1
'            ilTestCff = tgClfRC(ilClf).iFirstCff
'            If ilTestCff = ilLoop Then
'                ilFound = True
'                Exit For
'            End If
'            Do While ilTestCff <> -1
'                ilTestCff = tgCffRC(ilTestCff).iNextCff
'                If ilTestCff = ilLoop Then
'                    ilFound = True
'                    Exit For
'                End If
'            Loop
'            Next ilClf
'            If Not ilFound Then
'                ilCffIndex = ilLoop
'                ilLastCffUsed = ilLoop
'                Exit For
'            End If
'        End If
'    Next ilLoop
'    If ilCffIndex = -1 Then
'        ilCffIndex = UBound(tgCffRC)
'        ilLastCffUsed = ilCffIndex
'        ReDim Preserve tgCffRC(0 To ilCffIndex + 100) As CFFLIST
'        For ilLoop = ilCffIndex + 1 To UBound(tgCffRC) Step 1
'            tgCffRC(ilLoop).iStatus = -1 'Not Used
'            tgCffRC(ilLoop).lRecPos = 0
'            tgCffRC(ilLoop).iNextCff = -1
'        Next ilLoop
'    End If
'    Return
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mAdjustCntr                     *
''*                                                     *
''*             Created:8/25/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Redistribute dollars to        *
''*                      specified contract             *
''*                                                     *
''*******************************************************
'Private Function mAdjustCntr(slSDate As String) As Integer
''   tmChf (I)- Contract image
'    Dim ilRet As Integer
'    Dim ilCRet As Integer
'    Dim ilClf As Integer
'    Dim ilCff As Integer
'    Dim ilVef As Integer
'    Dim ilAdf As Integer
'    Dim ilHClf As Integer
'    Dim ilHCff As Integer
'    Dim ilHVef As Integer
'    Dim llStartDate As Long
'    Dim llEndDate As Long
'    Dim llSDate As Long
'    Dim llEDate As Long
'    Dim llDate As Long
'    Dim llTSPrice As Long
'    Dim llTPrice As Long
'    Dim ilDay As Integer
'    Dim ilMnfSocEco As Integer
'    Dim ilMnfDemo As Integer
'    Dim slNameCode As String
'    Dim slCode As String
'    Dim ilLoop As Integer
'    Dim ilSpots As Integer
'    Dim ilWkSpots As Integer
'    Dim ilMSpots As Integer
'    Dim ilMCff As Integer
'    Dim llAvgAud As Long
'    Dim slStr As String
'    Dim llTAud As Long
'    Dim llOvStartTime As Long
'    Dim llOvEndTime As Long
'    Dim llTstDate As Long
'    Dim ilHdShown As Integer
'    Dim ilUpdateCntr As Integer
'    Dim slPrtCntr As String * 8
'    Dim slPrtName As String * 55
'    Dim slPrtLn As String * 4
'    Dim slPrtVeh As String * 20
'    Dim slPrtDate As String * 8
'    Dim slPrtOld As String * 12
'    Dim slPrtNew As String * 12
'    Dim llCntrOld As Long
'    Dim llCntrNew As Long
'    Dim ilPassDnfCode As Integer
'    ReDim ilSDate(0 To 1) As Integer
'    ReDim ilEDate(0 To 1) As Integer
'    Dim tlCff As CFF
'    ReDim ilAudDays(0 To 6) As Integer
'    Dim llPopEst As Long
'    Dim ilAudFromSource As Integer
'    Dim llAudFromCode As Long
'
'    ilMnfSocEco = 0
'    ilUpdateCntr = False
'    llCntrOld = 0
'    llCntrNew = 0
'    If rbcAud(0).Value Then
'        ilMnfDemo = tmChf.iMnfDemo(0)
'        If ilMnfDemo <= 0 Then
'            'If lbcAud(0).ListIndex >= 1 Then
'            '    slNameCode = tgDemoCode(lbcDemo(0).ListIndex - 1).sKey  'Traffic!lbcDemoCode.List(lbcDemo(ilLoop).ListIndex - 1)
'            '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
'            '    ilMnfDemo = CInt(slCode)
'            'Else
'            '    ilMnfDemo = 0
'            'End If
'            slStr = edcAud(0).Text
'            gFindMatch slStr, 0, lbcAud
'            If gLastFound(lbcAud) >= 0 Then
'                slNameCode = tgDemoCode(gLastFound(lbcAud)).sKey  'Traffic!lbcDemoCode.List(lbcDemo(ilLoop).ListIndex - 1)
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                ilMnfDemo = CInt(slCode)
'            Else
'                ilMnfDemo = 0
'            End If
'        End If
'    Else
'        slStr = edcAud(1).Text
'        gFindMatch slStr, 0, lbcAud
'        If gLastFound(lbcAud) >= 0 Then
'            slNameCode = tgDemoCode(gLastFound(lbcAud)).sKey  'Traffic!lbcDemoCode.List(lbcDemo(ilLoop).ListIndex - 1)
'            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'            ilMnfDemo = CInt(slCode)
'        Else
'            ilMnfDemo = 0
'        End If
'    End If
'    If ilMnfDemo <= 0 Then
'        mAdjustCntr = True
'        Exit Function
'    End If
'    ilRet = gObtainCntr(hmChf, hmClf, hmCff, tmChf.lCode, False, tgChfRC, tgClfRC(), tgCffRC())
'    Print #hmMsg, ""
'    Print #hmMsg, "Contract"; Spc(1); "Advertiser/Product"; Spc(38); "Line"; Spc(1); "Pkg/Vehicle"; Spc(10); "Wk Date"; Spc(5); "Old Price"; Spc(4); "New Price"
'    ilHdShown = False
'    imLastCffUsed = UBound(tgCffRC) - 1
'    ReDim lmCffRecPos(0 To 0) As Long
'    For ilLoop = LBound(tgCffRC) To UBound(tgCffRC) - 1 Step 1
'        If tgCffRC(ilLoop).iStatus > 0 Then
'            lmCffRecPos(UBound(lmCffRecPos)) = tgCffRC(ilLoop).lRecPos
'            ReDim Preserve lmCffRecPos(0 To UBound(lmCffRecPos) + 1) As Long
'        End If
'    Next ilLoop
'    'Find package line
'    For ilClf = LBound(tgClfRC) To UBound(tgClfRC) - 1 Step 1
'        tmClf = tgClfRC(ilClf).ClfRec
'        If (tmClf.sType = "O") Or (tmClf.sType = "A") Or (tmClf.sType = "E") Then
'            'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
'            '    If tmClf.iVefCode = tgMVef(ilVef).iCode Then
'                ilVef = gBinarySearchVef(tmClf.iVefCode)
'                If ilVef <> -1 Then
'                    'If (tgMVef(ilVef).sStdPrice = "A") Then
'                    If ((tgMVef(ilVef).lPvfCode > 0) And (tgMVef(ilVef).sStdPrice = "A")) Or ((tgMVef(ilVef).lPvfCode = 0) And (tgSpf.sCAudPkg = "Y")) Then
'                        'Redistribute dollars to hidden lines
'                        gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llStartDate
'                        gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llEndDate
'                        Do While gWeekDayLong(llStartDate) <> 0
'                            llStartDate = llStartDate - 1
'                        Loop
'                        Do While gWeekDayLong(llEndDate) <> 6
'                            llEndDate = llEndDate + 1
'                        Loop
'                        For llDate = llStartDate To llEndDate Step 7
'                            If llDate >= gDateValue(slSDate) Then
'                                'Find week price
'                                ilCff = tgClfRC(ilClf).iFirstCff
'                                Do While ilCff >= 0
'                                    gUnpackDateLong tgCffRC(ilCff).CffRec.iStartDate(0), tgCffRC(ilCff).CffRec.iStartDate(1), llSDate
'                                    gUnpackDateLong tgCffRC(ilCff).CffRec.iEndDate(0), tgCffRC(ilCff).CffRec.iEndDate(1), llEDate
'                                    If llEDate < llSDate Then
'                                        Exit Do
'                                    End If
'                                    Do While gWeekDayLong(llSDate) <> 0
'                                        llSDate = llSDate - 1
'                                    Loop
'                                    If (llDate >= llSDate) And (llDate <= llEDate) Then
'                                        llTSPrice = 0
'                                        ReDim tmCffAud(0 To 0) As CFFAUD
'                                        If tmClf.sType <> "E" Then
'                                            If tgCffRC(ilCff).CffRec.sPriceType = "T" Then
'                                                If tgCffRC(ilCff).CffRec.sDyWk = "D" Then
'                                                    For ilDay = 0 To 6 Step 1
'                                                        If llDate + ilDay <= llEDate Then
'                                                            llTSPrice = llTSPrice + tgCffRC(ilCff).CffRec.iDay(ilDay) * tgCffRC(ilCff).CffRec.lActPrice
'                                                        End If
'                                                    Next ilDay
'                                                Else
'                                                    llTSPrice = llTSPrice + (tgCffRC(ilCff).CffRec.iSpotsWk + tgCffRC(ilCff).CffRec.iXSpotsWk) * tgCffRC(ilCff).CffRec.lActPrice
'                                                End If
'                                            End If
'                                            ilWkSpots = 0
'                                            For ilHClf = LBound(tgClfRC) To UBound(tgClfRC) - 1 Step 1
'                                                If (tgClfRC(ilHClf).ClfRec.sType = "H") And (tmClf.iLine = tgClfRC(ilHClf).ClfRec.iPkLineNo) Then
'                                                    ilHCff = tgClfRC(ilHClf).iFirstCff
'                                                    Do While ilHCff >= 0
'                                                        gUnpackDateLong tgCffRC(ilHCff).CffRec.iStartDate(0), tgCffRC(ilHCff).CffRec.iStartDate(1), llSDate
'                                                        gUnpackDateLong tgCffRC(ilHCff).CffRec.iEndDate(0), tgCffRC(ilHCff).CffRec.iEndDate(1), llEDate
'                                                        If llEDate < llSDate Then
'                                                            Exit Do
'                                                        End If
'                                                        Do While gWeekDayLong(llSDate) <> 0
'                                                            llSDate = llSDate - 1
'                                                        Loop
'                                                        If (llDate >= llSDate) And (llDate <= llEDate) Then
'                                                            If tgCffRC(ilHCff).CffRec.sDyWk = "D" Then
'                                                                For ilDay = 0 To 6 Step 1
'                                                                    If llDate + ilDay <= llEndDate Then
'                                                                        ilSpots = ilSpots + tgCffRC(ilHCff).CffRec.iDay(ilDay)
'                                                                    End If
'                                                                Next ilDay
'                                                            Else
'                                                                ilSpots = tgCffRC(ilHCff).CffRec.iSpotsWk + tgCffRC(ilHCff).CffRec.iXSpotsWk
'                                                            End If
'                                            '                llTSPrice = llTSPrice + (ilSpots) * tgCffRC(ilHCff).CffRec.lActPrice
'                                                            If tgCffRC(ilHCff).CffRec.sPriceType = "T" Then
'                                                                ilWkSpots = ilWkSpots + ilSpots
'                                                            End If
'                                                            Exit Do
'                                                        End If
'                                                        ilHCff = tgCffRC(ilHCff).iNextCff
'                                                    Loop
'                                                End If
'                                            Next ilHClf
'                                        Else
'                                            llTSPrice = tgCffRC(ilCff).CffRec.lActPrice
'                                        End If
'                                        'Get audience value for each hidden line of the package
'                                        llTAud = 0
'                                        For ilHClf = LBound(tgClfRC) To UBound(tgClfRC) - 1 Step 1
'                                            If (tgClfRC(ilHClf).ClfRec.sType = "H") And (tmClf.iLine = tgClfRC(ilHClf).ClfRec.iPkLineNo) Then
'                                                'For ilHVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
'                                                '    If tgClfRC(ilHClf).ClfRec.iVefCode = tgMVef(ilHVef).iCode Then
'                                                    ilHVef = gBinarySearchVef(tgClfRC(ilHClf).ClfRec.iVefCode)
'                                                    If ilHVef <> -1 Then
'                                                        ilHCff = tgClfRC(ilHClf).iFirstCff
'                                                        Do While ilHCff >= 0
'                                                            gUnpackDateLong tgCffRC(ilHCff).CffRec.iStartDate(0), tgCffRC(ilHCff).CffRec.iStartDate(1), llSDate
'                                                            gUnpackDateLong tgCffRC(ilHCff).CffRec.iEndDate(0), tgCffRC(ilHCff).CffRec.iEndDate(1), llEDate
'                                                            If llEDate < llSDate Then
'                                                                Exit Do
'                                                            End If
'                                                            Do While gWeekDayLong(llSDate) <> 0
'                                                                llSDate = llSDate - 1
'                                                            Loop
'                                                            If (llDate >= llSDate) And (llDate <= llEDate) Then
'                                                                For ilDay = 0 To 6 Step 1
'                                                                    ilAudDays(ilDay) = False
'                                                                Next ilDay
'                                                                For ilDay = 0 To 6 Step 1
'                                                                    If (llDate + ilDay <= llEDate) Then
'                                                                        If tgCffRC(ilHCff).CffRec.iDay(ilDay) > 0 Then
'                                                                            ilAudDays(ilDay) = True
'                                                                        End If
'                                                                    End If
'                                                                Next ilDay
'                                                                If ((tgClfRC(ilHClf).ClfRec.iStartTime(0) <> 1) Or (tgClfRC(ilHClf).ClfRec.iStartTime(1) <> 0)) And ((tgClfRC(ilHClf).ClfRec.iEndTime(0) <> 1) Or (tgClfRC(ilHClf).ClfRec.iEndTime(1) <> 0)) Then
'                                                                    gUnpackTimeLong tgClfRC(ilHClf).ClfRec.iStartTime(0), tgClfRC(ilHClf).ClfRec.iStartTime(1), False, llOvStartTime
'                                                                    gUnpackTimeLong tgClfRC(ilHClf).ClfRec.iEndTime(0), tgClfRC(ilHClf).ClfRec.iEndTime(1), True, llOvEndTime
'                                                                Else
'                                                                    llOvStartTime = 0
'                                                                    llOvEndTime = 0
'                                                                End If
'                                                                If tgCffRC(ilHCff).CffRec.sDyWk = "D" Then
'                                                                    For ilDay = 0 To 6 Step 1
'                                                                        If llDate + ilDay <= llEndDate Then
'                                                                            ilSpots = ilSpots + tgCffRC(ilHCff).CffRec.iDay(ilDay)
'                                                                        End If
'                                                                    Next ilDay
'                                                                Else
'                                                                    ilSpots = tgCffRC(ilHCff).CffRec.iSpotsWk + tgCffRC(ilHCff).CffRec.iXSpotsWk
'                                                                End If
'                                                                If tgCffRC(ilHCff).CffRec.sPriceType = "T" Then
'                                                                    ilPassDnfCode = tgMVef(ilHVef).iReallDnfCode
'                                                                    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilPassDnfCode, tgClfRC(ilHClf).ClfRec.iVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, tgClfRC(ilHClf).ClfRec.iRdfcode, llOvStartTime, llOvEndTime, ilAudDays(), tgClfRC(ilHClf).ClfRec.sType, tgClfRC(ilHClf).ClfRec.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
'                                                                Else
'                                                                    llAvgAud = 0
'                                                                End If
'                                                                tlCff = tgCffRC(ilHCff).CffRec
'                                                                mAddWeek imLastCffUsed, ilHClf + 1, llDate, ilSpots, ilAudDays()
'                                                                ilHCff = tgClfRC(ilHClf).iFirstCff
'                                                                Do While ilHCff >= 0
'                                                                    gUnpackDateLong tgCffRC(ilHCff).CffRec.iStartDate(0), tgCffRC(ilHCff).CffRec.iStartDate(1), llTstDate
'                                                                    If llTstDate = llDate Then
'                                                                        ilSDate(0) = tgCffRC(ilHCff).CffRec.iStartDate(0)
'                                                                        ilSDate(1) = tgCffRC(ilHCff).CffRec.iStartDate(1)
'                                                                        ilEDate(0) = tgCffRC(ilHCff).CffRec.iEndDate(0)
'                                                                        ilEDate(1) = tgCffRC(ilHCff).CffRec.iEndDate(1)
'                                                                        tgCffRC(ilHCff).CffRec = tlCff
'                                                                        tgCffRC(ilHCff).CffRec.iStartDate(0) = ilSDate(0)
'                                                                        tgCffRC(ilHCff).CffRec.iStartDate(1) = ilSDate(1)
'                                                                        tgCffRC(ilHCff).CffRec.iEndDate(0) = ilEDate(0)
'                                                                        tgCffRC(ilHCff).CffRec.iEndDate(1) = ilEDate(1)
'                                                                        tmCffAud(UBound(tmCffAud)).iCffIndex = ilHCff
'                                                                        tmCffAud(UBound(tmCffAud)).iVefIndex = ilHVef
'                                                                        tmCffAud(UBound(tmCffAud)).lAud = llAvgAud
'                                                                        ReDim Preserve tmCffAud(0 To UBound(tmCffAud) + 1) As CFFAUD
'                                                                        llTAud = llTAud + ilSpots * llAvgAud
'                                                                        Exit Do
'                                                                    End If
'                                                                    ilHCff = tgCffRC(ilHCff).iNextCff
'                                                                Loop
'                                                                Exit Do
'                                                            End If
'                                                            ilHCff = tgCffRC(ilHCff).iNextCff
'                                                        Loop
'                                                '        Exit For
'                                                    End If
'                                                'Next ilHVef
'                                            End If
'                                        Next ilHClf
'                                        'Redistribute dollars
'                                        llTPrice = llTSPrice
'                                        ilMCff = -1
'
'                                        For ilLoop = 0 To UBound(tmCffAud) - 1 Step 1
'                                            ilUpdateCntr = True
'                                            ilHCff = tmCffAud(ilLoop).iCffIndex
'                                            ilSpots = 0
'                                            If tgCffRC(ilHCff).CffRec.sDyWk = "D" Then
'                                                For ilDay = 0 To 6 Step 1
'                                                    If llDate + ilDay <= llEndDate Then
'                                                        ilSpots = ilSpots + tgCffRC(ilHCff).CffRec.iDay(ilDay)
'                                                    End If
'                                                Next ilDay
'                                            Else
'                                                ilSpots = tgCffRC(ilHCff).CffRec.iSpotsWk + tgCffRC(ilHCff).CffRec.iXSpotsWk
'                                            End If
'                                            If tgCffRC(ilHCff).CffRec.sPriceType = "T" Then
'                                                If ilSpots > 0 Then
'                                                    If ilMCff = -1 Then
'                                                        ilMSpots = ilSpots
'                                                        ilMCff = ilHCff
'                                                    Else
'                                                        If ilSpots < ilMSpots Then
'                                                            ilMSpots = ilSpots
'                                                            ilMCff = ilHCff
'                                                        End If
'                                                    End If
'                                                End If
'                                            End If
'                                            If ilHdShown Then
'                                                slPrtCntr = ""
'                                                slPrtName = ""
'                                            Else
'                                                slPrtCntr = Trim$(str$(tgChfRC.lCntrNo))
'                                                'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
'                                                '    If tgCommAdf(ilAdf).iCode = tgChfRC.iAdfCode Then
'                                                    ilAdf = gBinarySearchAdf(tgChfRC.iadfCode)
'                                                    If ilAdf <> -1 Then
'                                                        If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
'                                                            slPrtName = Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID) & "/" & Trim$(tgChfRC.sProduct)
'                                                        Else
'                                                            slPrtName = Trim$(tgCommAdf(ilAdf).sName) & "/" & Trim$(tgChfRC.sProduct)
'                                                        End If
'                                                '        Exit For
'                                                    End If
'                                                'Next ilAdf
'                                                ilHdShown = True
'                                            End If
'                                            slPrtLn = Trim$(str$(tgCffRC(ilHCff).CffRec.iClfLine))
'                                            slPrtVeh = tgMVef(tmCffAud(ilLoop).iVefIndex).sName
'                                            gUnpackDate tgCffRC(ilHCff).CffRec.iStartDate(0), tgCffRC(ilHCff).CffRec.iStartDate(1), slStr
'                                            slPrtDate = slStr
'                                            llCntrOld = llCntrOld + ilSpots * tgCffRC(ilHCff).CffRec.lActPrice
'                                            lmTotalOld = lmTotalOld + ilSpots * tgCffRC(ilHCff).CffRec.lActPrice
'                                            slStr = gLongToStrDec(ilSpots * tgCffRC(ilHCff).CffRec.lActPrice, 2)
'                                            gFormatStr slStr, FMTCOMMA, 2, slStr
'                                            Do While Len(slStr) < Len(slPrtOld)
'                                                slStr = " " & slStr
'                                            Loop
'                                            slPrtOld = slStr
'                                            If tgCffRC(ilHCff).CffRec.sPriceType = "T" Then
'                                                If (llTAud > 0) And (ilSpots > 0) Then
'                                                    tgCffRC(ilHCff).CffRec.lActPrice = (CSng(tmCffAud(ilLoop).lAud) * llTSPrice) / (llTAud)
'                                                ElseIf (llTAud = 0) And (ilWkSpots > 0) Then
'                                                    tgCffRC(ilHCff).CffRec.lActPrice = (llTSPrice) / (ilWkSpots)
'                                                Else
'                                                    tgCffRC(ilHCff).CffRec.lActPrice = 0
'                                                End If
'                                            Else
'                                                tgCffRC(ilHCff).CffRec.lActPrice = 0
'                                            End If
'                                            llCntrNew = llCntrNew + ilSpots * tgCffRC(ilHCff).CffRec.lActPrice
'                                            lmTotalNew = lmTotalNew + ilSpots * tgCffRC(ilHCff).CffRec.lActPrice
'                                            slStr = gLongToStrDec(ilSpots * tgCffRC(ilHCff).CffRec.lActPrice, 2)
'                                            gFormatStr slStr, FMTCOMMA, 2, slStr
'                                            Do While Len(slStr) < Len(slPrtNew)
'                                                slStr = " " & slStr
'                                            Loop
'                                            slPrtNew = slStr
'                                            Print #hmMsg, slPrtCntr; Spc(1); slPrtName; Spc(1); slPrtLn; Spc(1); slPrtVeh; Spc(1); slPrtDate; Spc(1); slPrtOld; Spc(1); slPrtNew
'                                            llTPrice = llTPrice - ilSpots * tgCffRC(ilHCff).CffRec.lActPrice
'                                        Next ilLoop
'                                        'Balance dollars
'                                        If (llTPrice <> 0) And (ilMCff <> -1) Then
'                                            tgCffRC(ilMCff).CffRec.lActPrice = tgCffRC(ilMCff).CffRec.lActPrice + (llTPrice / ilMSpots)
'                                        End If
'                                    End If
'                                    ilCff = tgCffRC(ilCff).iNextCff
'                                Loop
'                            End If
'                        Next llDate
'                    End If
'            '        Exit For
'                End If
'            'Next ilVef
'        End If
'    Next ilClf
'    If ilUpdateCntr Then
'        slStr = gLongToStrDec(llCntrOld, 2)
'        gFormatStr slStr, FMTCOMMA, 2, slStr
'        Do While Len(slStr) < Len(slPrtOld)
'            slStr = " " & slStr
'        Loop
'        slPrtOld = slStr
'        slStr = gLongToStrDec(llCntrNew, 2)
'        gFormatStr slStr, FMTCOMMA, 2, slStr
'        Do While Len(slStr) < Len(slPrtNew)
'            slStr = " " & slStr
'        Loop
'        slPrtNew = slStr
'        slPrtName = "Contract Total"
'        slPrtLn = ""
'        slPrtVeh = ""
'        slPrtDate = ""
'        Print #hmMsg, slPrtCntr; Spc(1); slPrtName; Spc(1); slPrtLn; Spc(1); slPrtVeh; Spc(1); slPrtDate; Spc(1); slPrtOld; Spc(1); slPrtNew
'        Print #hmMsg, "Updating Contract"; Spc(1); Trim$(str$(tgChfRC.lCntrNo))
'        gGetSyncDateTime smSyncDate, smSyncTime
'        'If (tgChfRC.sStatus = "O") Or (tgChfRC.sStatus = "H") Then
'        '    'Update CntRevNo
'        'Else
'            'Replace Proposal or unscheduled (O or H) Flights
'            ilRet = btrBeginTrans(hmChf, 1000)
'            For ilLoop = 0 To UBound(lmCffRecPos) - 1 Step 1
'                Do
'                    ilRet = btrGetDirect(hmCff, tlCff, imCffRecLen, lmCffRecPos(ilLoop), INDEXKEY0, BTRV_LOCK_NONE)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        ilCRet = btrAbortTrans(hmChf)
'                        Print #hmMsg, "Updating Contract"; Spc(1); Trim$(str$(tgChfRC.lCntrNo)); Spc(1); "Failed (GetDirect/Cff): error #" & str$(ilRet)
'                        mAdjustCntr = False
'                        Exit Function
'                    End If
'                    'tmRec = tlCff
'                    'ilRet = gGetByKeyForUpdate("Cff", hmCff, tmRec)
'                    'If ilRet <> BTRV_ERR_NONE Then
'                    '    ilCRet = btrAbortTrans(hmChf)
'                    '    Print #hmMsg, "Updating Contract"; Spc(1); Trim$(Str$(tgChfRC.lCntrNo)); Spc(1); "Failed (GetByKey/Cff): error #" & Str$(ilRet)
'                    '    mAdjustCntr = False
'                    '    Exit Function
'                    'End If
'                    'tlCff = tmRec
'                    ilRet = btrDelete(hmCff)
'                Loop While ilRet = BTRV_ERR_CONFLICT
'                If ilRet <> BTRV_ERR_NONE Then
'                    ilCRet = btrAbortTrans(hmChf)
'                    Print #hmMsg, "Updating Contract"; Spc(1); Trim$(str$(tgChfRC.lCntrNo)); Spc(1); "Failed (Delete/Cff): error #" & str$(ilRet)
'                    mAdjustCntr = False
'                    Exit Function
'                End If
'            Next ilLoop
'            For ilLoop = LBound(tgCffRC) To UBound(tgCffRC) - 1 Step 1
'                If tgCffRC(ilLoop).iStatus >= 0 Then
'                    tgCffRC(ilLoop).CffRec.lCode = 0
'                    ilRet = btrInsert(hmCff, tgCffRC(ilLoop).CffRec, imCffRecLen, INDEXKEY1)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        ilCRet = btrAbortTrans(hmChf)
'                        Print #hmMsg, "Updating Contract"; Spc(1); Trim$(str$(tgChfRC.lCntrNo)); Spc(1); "Failed (Insert/Cff): error #" & str$(ilRet)
'                        mAdjustCntr = False
'                        Exit Function
'                    End If
'                End If
'            Next ilLoop
'            Do
'                tmChfSrchKey.lCode = tgChfRC.lCode
'                ilRet = btrGetEqual(hmChf, tgChfRC, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'                If ilRet <> BTRV_ERR_NONE Then
'                    ilCRet = btrAbortTrans(hmChf)
'                    Print #hmMsg, "Updating Contract"; Spc(1); Trim$(str$(tgChfRC.lCntrNo)); Spc(1); "Failed (GetEqual/Chf): error #" & str$(ilRet)
'                    mAdjustCntr = False
'                    Exit Function
'                End If
'                tgChfRC.iUrfCode = tgUrf(0).iCode
'                'tgChfRC.iSourceID = tgUrf(0).iRemoteUserID
'                'gPackDate smSyncDate, tgChfRC.iSyncDate(0), tgChfRC.iSyncDate(1)
'                'gPackTime smSyncTime, tgChfRC.iSyncTime(0), tgChfRC.iSyncTime(1)
'                ilRet = btrUpdate(hmChf, tgChfRC, imChfRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'            If ilRet <> BTRV_ERR_NONE Then
'                Print #hmMsg, "Updating Contract"; Spc(1); Trim$(str$(tgChfRC.lCntrNo)); Spc(1); "Failed (Update/Chf): error #" & str$(ilRet)
'                ilCRet = btrAbortTrans(hmChf)
'                mAdjustCntr = False
'                Exit Function
'            End If
'            ilRet = btrEndTrans(hmChf)
'            Print #hmMsg, "Updated Contract"; Spc(1); Trim$(str$(tgChfRC.lCntrNo)); Spc(1); "Successfully"
'        'End If
'    End If
'    mAdjustCntr = True
'    Exit Function
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:mBoxCalDate                     *
''*                                                     *
''*             Created:8/25/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Place box around calendar date *
''*                                                     *
''*******************************************************
'Private Sub mBoxCalDate()
'    Dim slStr As String
'    Dim ilRowNo As Integer
'    Dim llInputDate As Long
'    Dim ilWkDay As Integer
'    Dim slDay As String
'    Dim llDate As Long
'    slStr = edcDate.Text
'    If gValidDate(slStr) Then
'        llInputDate = gDateValue(slStr)
'        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
'            ilRowNo = 0
'            llDate = lmCalStartDate
'            Do
'                ilWkDay = gWeekDayLong(llDate)
'                slDay = Trim$(str$(Day(llDate)))
'                If llDate = llInputDate Then
'                    lacDate.Caption = slDay
'                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
'                    lacDate.Visible = True
'                    Exit Sub
'                End If
'                If ilWkDay = 6 Then
'                    ilRowNo = ilRowNo + 1
'                End If
'                llDate = llDate + 1
'            Loop Until llDate > lmCalEndDate
'            lacDate.Visible = False
'        Else
'            lacDate.Visible = False
'        End If
'    Else
'        lacDate.Visible = False
'    End If
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mDemoPop                        *
''*                                                     *
''*             Created:7/19/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Populate Demo list             *
''*                      box                            *
''*                                                     *
''*******************************************************
'Private Sub mDemoPop()
'    Dim ilRet As Integer
'    ilRet = gPopMnfPlusFieldsBox(RCReallo, lbcAud, tgDemoCode(), sgDemoCodeTag, "D")
'    lbcAud.Height = gListBoxHeight(lbcAud.ListCount, 5)
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mInit                           *
''*                                                     *
''*             Created:9/02/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Initialize modular             *
''*                                                     *
''*******************************************************
'Private Sub mInit()
''
''   mInit
''   Where:
''
'    Dim ilRet As Integer
'    Dim slStr As String
'
'    Screen.MousePointer = vbHourglass
'    imFirstActivate = True
'    imAudIndex = -1
'    imBypassFocus = False
'    imChgMode = False
'    imBSMode = False
'    imLbcArrowSetting = False
'    imLbcMouseDown = False
'    imCalType = 0   'Standard
'    mInitBox
'    hmChf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", RCReallo
'    On Error GoTo 0
'    imChfRecLen = Len(tmChf)
'    hmClf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", RCReallo
'    On Error GoTo 0
'    imClfRecLen = Len(tmClf)
'    hmCff = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", RCReallo
'    On Error GoTo 0
'    imCffRecLen = Len(tmCff)
'    hmMnf = CBtrvTable(ONEHANDLE)
'    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", RCReallo
'    On Error GoTo 0
'    imMnfRecLen = Len(tmMnf)
'    hmDrf = CBtrvTable(ONEHANDLE)
'    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Drf.Btr)", RCReallo
'    On Error GoTo 0
'    imDrfRecLen = Len(tmDrf)
'    hmDpf = CBtrvTable(ONEHANDLE)
'    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dpf.Btr)", RCReallo
'    On Error GoTo 0
'    hmDef = CBtrvTable(ONEHANDLE)
'    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Def.Btr)", RCReallo
'    On Error GoTo 0
'    hmRaf = CBtrvTable(ONEHANDLE)
'    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", RCReallo
'    On Error GoTo 0
'    ilRet = gObtainAdvt()
'    ilRet = gObtainVef()
'    RCReallo.Height = cmcRealloc.Top + 5 * cmcRealloc.Height / 3
'    gCenterModalForm RCReallo
'    slStr = Format$(gNow(), "m/d/yy")
'    lmNowDate = gDateValue(slStr)
'    smStartDate = gObtainNextMonday(Format$(gNow(), "m/d/yy"))
'    edcDate.Text = smStartDate
'    mDemoPop
'    imAudIndex = 0
'    edcAud(0).Text = "P12+"
'    imAudIndex = 1
'    edcAud(1).Text = "P12+"
'    imAudIndex = -1
'    Screen.MousePointer = vbDefault
'    Exit Sub
'mInitErr:
'    On Error GoTo 0
'    imTerminate = True
'    Exit Sub
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mInitBox                        *
''*                                                     *
''*             Created:6/30/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Set mouse and control locations*
''*                                                     *
''*******************************************************
'Private Sub mInitBox()
''
''   mInitBox
''   Where:
''
'    'Dim flTextHeight As Single  'Standard text height
'    Dim ilLoop As Integer
'    'flTextHeight = pbcDate.TextHeight("1") - 35
'    'Position panel and picture areas with panel
'    plcRealloc.Move 105, 375
'    plcCalendar.Move plcRealloc.Left + edcDate.Left, plcRealloc.Top + edcDate.Top + edcDate.Height
'    'Calendar
'    For ilLoop = 1 To 7 Step 1
'        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
'    Next ilLoop
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mOpenMsgFile                    *
''*                                                     *
''*             Created:5/18/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Open error message file         *
''*                                                     *
''*******************************************************
'Private Function mOpenMsgFile() As Integer
'    Dim slToFile As String
'    Dim slDateTime As String
'    Dim slFileDate As String
'    Dim ilRet As Integer
'    On Error GoTo mOpenMsgFileErr:
'    'slToFile = sgExportPath & "Realloc.Txt"
'    slToFile = sgDBPath & "Messages\" & "Realloc.Txt"
'    slDateTime = FileDateTime(slToFile)
'    If ilRet = 0 Then
'        slFileDate = Format$(slDateTime, "m/d/yy")
'        If gDateValue(slFileDate) = lmNowDate Then  'Append
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Append As hmMsg
'            If ilRet <> 0 Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        Else
'            Kill slToFile
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Output As hmMsg
'            If ilRet <> 0 Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        End If
'    Else
'        On Error GoTo 0
'        ilRet = 0
'        On Error GoTo mOpenMsgFileErr:
'        hmMsg = FreeFile
'        Open slToFile For Output As hmMsg
'        If ilRet <> 0 Then
'            Screen.MousePointer = vbDefault
'            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'            mOpenMsgFile = False
'            Exit Function
'        End If
'    End If
'    On Error GoTo 0
'    Print #hmMsg, "** Reallocation of Dollars: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
'    Print #hmMsg, ""
'    mOpenMsgFile = True
'    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:mRealloc                        *
''*                                                     *
''*             Created:9/02/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Reallocate dollars to hidden   *
''*                      lines                          *
''*                                                     *
''*******************************************************
'Private Sub mRealloc(slSDate As String)
'    Dim slCntrType As String
'    Dim slCntrStatus As String
'    Dim ilHOType As Integer
'    Dim slEDate As String
'    Dim ilPass As Integer
'    Dim ilLoop As Integer
'    Dim ilVef As Integer
'    Dim ilRet As Integer
'    Dim ilClf As Integer
'    Dim ilFound As Integer
'    For ilPass = 0 To 1 Step 1
'        slCntrType = "" 'All types
'        If (ckcChg(0).Value = vbChecked) And (ilPass = 0) Then
'            slCntrStatus = "WCI"
'            ilHOType = 1
'        End If
'        If (ckcChg(1).Value = vbChecked) And (ilPass = 1) Then
'            slCntrStatus = "HO"
'            ilHOType = 2
'        End If
'        If ckcChg(ilPass).Value = vbChecked Then
'            slEDate = "12/31/2069"
'            sgCntrForDateStamp = ""
'            ilRet = gObtainCntrForDate(RCReallo, slSDate, slEDate, slCntrStatus, slCntrType, ilHOType, tmChfAdvtExt())
'            If (ilRet = CP_MSG_NOPOPREQ) Or (ilRet = CP_MSG_NONE) Then
'                For ilLoop = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
'                    ilRet = gObtainChfClf(hmChf, hmClf, tmChfAdvtExt(ilLoop).lCode, False, tmChf, tgClfRC())
'                    If ilRet Then
'                        ilFound = False
'                        For ilClf = LBound(tgClfRC) To UBound(tgClfRC) - 1 Step 1
'                            tmClf = tgClfRC(ilClf).ClfRec
'                            If (tmClf.sType = "O") Or (tmClf.sType = "A") Or (tmClf.sType = "E") Then
'                                'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
'                                '    If tmClf.iVefCode = tgMVef(ilVef).iCode Then
'                                    ilVef = gBinarySearchVef(tmClf.iVefCode)
'                                    If ilVef <> -1 Then
'                                        'If (tgMVef(ilVef).sStdPrice = "A") Then
'                                        If ((tgMVef(ilVef).lPvfCode > 0) And (tgMVef(ilVef).sStdPrice = "A")) Or ((tgMVef(ilVef).lPvfCode = 0) And (tgSpf.sCAudPkg = "Y")) Then
'                                            'Redistribute dollars to hidden lines
'                                            ilRet = mAdjustCntr(slSDate)
'                                            ilFound = True
'                                        End If
'                                '        Exit For
'                                    End If
'                                'Next ilVef
'                            End If
'                            If ilFound Then
'                                Exit For
'                            End If
'                        Next ilClf
'                    End If
'                Next ilLoop
'            End If
'        End If
'    Next ilPass
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mTerminate                      *
''*                                                     *
''*             Created:5/18/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: terminate form                 *
''*                                                     *
''*******************************************************
'Private Sub mTerminate()
''
''   mTerminate
''   Where:
''
'    Dim ilRet As Integer
'    Erase tgClfRC
'    Erase tmChfAdvtExt
'    Erase tgCffRC
'    Erase tmCffAud
'    Erase lmCffRecPos
'    btrExtClear hmDrf   'Clear any previous extend operation
'    ilRet = btrClose(hmRaf)
'    btrDestroy hmRaf
'    ilRet = btrClose(hmDef)
'    btrDestroy hmDef
'    ilRet = btrClose(hmDpf)
'    btrDestroy hmDpf
'    ilRet = btrClose(hmDrf)
'    btrDestroy hmDrf
'    btrExtClear hmMnf   'Clear any previous extend operation
'    ilRet = btrClose(hmMnf)
'    btrDestroy hmMnf
'    btrExtClear hmCff   'Clear any previous extend operation
'    ilRet = btrClose(hmCff)
'    btrDestroy hmCff
'    btrExtClear hmClf   'Clear any previous extend operation
'    ilRet = btrClose(hmClf)
'    btrDestroy hmClf
'    btrExtClear hmChf   'Clear any previous extend operation
'    ilRet = btrClose(hmChf)
'    btrDestroy hmChf
'    Screen.MousePointer = vbDefault
'    Unload RCReallo
'    Set RCReallo = Nothing   'Remove data segment
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mUpdateSpf                      *
''*                                                     *
''*             Created:8/25/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Update Reallocate Date in Spf  *
''*                                                     *
''*                                                     *
''*******************************************************
'Private Sub mUpdateSpf(slReallDate As String)
'    Dim ilRet As Integer
'    Dim llSpfRecPos As Long
'    hmSpf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    imSpfRecLen = Len(tmSpf)
'    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'    If ilRet = BTRV_ERR_NONE Then
'        ilRet = btrGetPosition(hmSpf, llSpfRecPos)
'        Do
'            ilRet = btrGetDirect(hmSpf, tmSpf, imSpfRecLen, llSpfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'            'tmSRec = tmSpf
'            'ilRet = gGetByKeyForUpdate("Spf", hmSpf, tmSRec)
'            'tmSpf = tmSRec
'            gPackDate slReallDate, tmSpf.iReallDate(0), tmSpf.iReallDate(1)
'            ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
'        Loop While ilRet = BTRV_ERR_CONFLICT
'    End If
'    btrExtClear hmSpf   'Clear any previous extend operation
'    ilRet = btrClose(hmSpf)
'End Sub
'Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim llDate As Long
'    Dim ilWkDay As Integer
'    Dim ilRowNo As Integer
'    Dim slDay As String
'    ilRowNo = 0
'    llDate = lmCalStartDate
'    Do
'        ilWkDay = gWeekDayLong(llDate)
'        slDay = Trim$(str$(Day(llDate)))
'        If (x >= tmCDCtrls(ilWkDay + 1).fBoxX) And (x <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
'            If (y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
'                edcDate.Text = Format$(llDate, "m/d/yy")
'                edcDate.SelStart = 0
'                edcDate.SelLength = Len(edcDate.Text)
'                imBypassFocus = True
'                edcDate.SetFocus
'                Exit Sub
'            End If
'        End If
'        If ilWkDay = 6 Then
'            ilRowNo = ilRowNo + 1
'        End If
'        llDate = llDate + 1
'    Loop Until llDate > lmCalEndDate
'    edcDate.SetFocus
'End Sub
'Private Sub pbcCalendar_Paint()
'    Dim slStr As String
'    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
'    lacCalName.Caption = gMonthYearFormat(slStr)
'    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
'    mBoxCalDate
'End Sub
'Private Sub pbcClickFocus_GotFocus()
'    plcCalendar.Visible = False
'    lbcAud.Visible = False
'End Sub
'Private Sub plcScreen_Click()
'    pbcClickFocus.SetFocus
'End Sub
'Private Sub rbcAud_Click(Index As Integer)
'    'Code added because Value removed as parameter
'    Dim Value As Integer
'    Value = rbcAud(Index).Value
'    'End of coded added
'    If Value Then
'        If Index = 0 Then
'            edcAud(1).Enabled = False
'            cmcAud(1).Enabled = False
'            edcAud(0).Enabled = True
'            cmcAud(0).Enabled = True
'        Else
'            edcAud(0).Enabled = False
'            cmcAud(0).Enabled = False
'            edcAud(1).Enabled = True
'            cmcAud(1).Enabled = True
'        End If
'    End If
'End Sub
'Private Sub rbcAud_GotFocus(Index As Integer)
'    plcCalendar.Visible = False
'    lbcAud.Visible = False
'End Sub
'Private Sub plcScreen_Paint()
'    plcScreen.CurrentX = 0
'    plcScreen.CurrentY = 0
'    plcScreen.Print "Reallocation of Package Dollars"
'End Sub
