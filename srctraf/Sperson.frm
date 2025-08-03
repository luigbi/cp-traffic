VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form SPerson 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4545
   ClientLeft      =   285
   ClientTop       =   1530
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
   ScaleHeight     =   4545
   ScaleWidth      =   7050
   Begin VB.TextBox edcIncClientComm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3480
      MaxLength       =   7
      TabIndex        =   22
      Top             =   2865
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox edcNewClientComm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   675
      MaxLength       =   7
      TabIndex        =   21
      Top             =   2835
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2445
      TabIndex        =   1
      Top             =   45
      Width           =   3945
   End
   Begin VB.TextBox edcRemOverComm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3480
      MaxLength       =   7
      TabIndex        =   20
      Top             =   2475
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox edcRemUnderComm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   645
      MaxLength       =   7
      TabIndex        =   19
      Top             =   2490
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   6435
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   465
      Visible         =   0   'False
      Width           =   1995
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
         Picture         =   "Sperson.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   42
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
            TabIndex        =   43
            Top             =   405
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
         TabIndex        =   41
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   345
         TabIndex        =   44
         Top             =   75
         Width           =   1305
      End
   End
   Begin VB.TextBox edcCommDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   4485
      MaxLength       =   10
      TabIndex        =   25
      Top             =   3195
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ListBox lbcTitle 
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
      Height          =   240
      Left            =   630
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1755
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4425
      ScaleHeight     =   210
      ScaleWidth      =   1395
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1785
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox edcOverComm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   4440
      MaxLength       =   7
      TabIndex        =   18
      Top             =   2130
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox edcUnderComm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2535
      MaxLength       =   7
      TabIndex        =   17
      Top             =   2130
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6300
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3780
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6570
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3750
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcStationCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3465
      MaxLength       =   5
      TabIndex        =   10
      Top             =   1110
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6690
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3645
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   645
      Top             =   3960
   End
   Begin VB.ListBox lbcTeam 
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
      Height          =   240
      Left            =   4920
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2070
      Visible         =   0   'False
      Width           =   2685
   End
   Begin MSMask.MaskEdBox mkcPhone 
      Height          =   210
      Left            =   690
      TabIndex        =   11
      Tag             =   "The number and extension of the buyer."
      Top             =   1455
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   16776960
      ForeColor       =   0
      MaxLength       =   24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(AAA) AAA-AAAA Ext(AAAA)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mkcFax 
      Height          =   210
      Left            =   3495
      TabIndex        =   12
      Tag             =   "The number and extension of the buyer."
      Top             =   1410
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   16776960
      ForeColor       =   0
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(AAA) AAA-AAAA"
      PromptChar      =   "_"
   End
   Begin VB.ListBox lbcSOffice 
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
      Height          =   240
      Left            =   4920
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   690
      MaxLength       =   20
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcDropDown 
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
      Left            =   3015
      Picture         =   "Sperson.frx":2E1A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmcCombo 
      Appearance      =   0  'Flat
      Caption         =   "Com&bo"
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
      Left            =   6225
      TabIndex        =   32
      Top             =   4110
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox edcLastName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3510
      MaxLength       =   20
      TabIndex        =   6
      Top             =   765
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.TextBox edcSales 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   2550
      MaxLength       =   10
      TabIndex        =   24
      Top             =   3210
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.TextBox edcCommPaid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   645
      MaxLength       =   8
      TabIndex        =   23
      Top             =   3180
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.TextBox edcGoal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   660
      MaxLength       =   9
      TabIndex        =   16
      Top             =   2130
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox edcFirstName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   645
      MaxLength       =   20
      TabIndex        =   5
      Top             =   765
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
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
      Left            =   3990
      TabIndex        =   34
      Top             =   4140
      Width           =   1050
   End
   Begin VB.CommandButton cmcMerge 
      Appearance      =   0  'Flat
      Caption         =   "&Merge"
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
      Left            =   2865
      TabIndex        =   33
      Top             =   4140
      Width           =   1050
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
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
      Left            =   1740
      TabIndex        =   31
      Top             =   4140
      Width           =   1050
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
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
      Left            =   4545
      TabIndex        =   30
      Top             =   3735
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
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
      Left            =   3420
      TabIndex        =   29
      Top             =   3735
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
      Left            =   2295
      TabIndex        =   28
      Top             =   3735
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
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
      Left            =   1170
      TabIndex        =   27
      Top             =   3735
      Width           =   1050
   End
   Begin VB.PictureBox pbcSlfID 
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
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   615
      Picture         =   "Sperson.frx":2F14
      ScaleHeight     =   2790
      ScaleWidth      =   5670
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   630
      Width           =   5670
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   90
      ScaleHeight     =   90
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   495
      Width           =   45
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   105
      TabIndex        =   26
      Top             =   2775
      Width           =   105
   End
   Begin VB.PictureBox plcSlfID 
      ForeColor       =   &H00000000&
      Height          =   2910
      Left            =   555
      ScaleHeight     =   2850
      ScaleWidth      =   5730
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   570
      Width           =   5790
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4800
      Width           =   75
   End
   Begin VB.Label plcScreen 
      Caption         =   "Salespeople"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   2835
   End
   Begin VB.Label lacCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   5520
      TabIndex        =   45
      Top             =   390
      Width           =   840
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   165
      Top             =   3990
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "SPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Sperson.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SPerson.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Salesperson input screen code
Option Explicit
Option Compare Text
'Salesperson Field Areas
Dim tmCtrls(0 To 19)  As FIELDAREA
Dim imLBCtrls As Integer
Dim tmSOfficeCode() As SORTCODE
Dim smSOfficeCodeTag As String
Dim tmTeamCode() As SORTCODE
Dim smTeamCodeTag As String
Dim imBoxNo As Integer   'Current Salesperson Box
Dim tmSlf As SLF        'SLF record image
Dim tmSlfSrchKey As INTKEY0    'SLF key record image
Dim imSlfRecLen As Integer        'SLF record length
'Dim hmDsf As Integer 'Delete Stamp file handle
'Dim tmDsf As DSF        'DSF record image
'Dim tmDsfSrchKey As LONGKEY0    'DSF key record image
'Dim imDsfRecLen As Integer        'DSF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim hmSlf As Integer 'Salesperson file handle
Dim imUpdateAllowed As Integer    'User can update records
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imState As Integer  '0=Active; 1=Dormant
Dim smPhoneImage As String  'Blank phone image- obtained from mkcPhone.text before input
Dim smFaxImage As String    'Blank fax image
Dim smSOffice As String     'Sales office name, saved to determine if changed
Dim smTeam As String        'Team name, saved to determine if changed
Dim imFirstFocus As Integer 'True=cbcSelect has not had focus yet, used to branch to another control
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imComboBoxIndex As Integer
Dim imBypassFocus As Integer
Dim imChgSaveFlag As Integer    'Indicates if any changed saved
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Const FIRSTNAMEINDEX = 1     'Name control/field
Const LASTNAMEINDEX = 2 'Last name control/field
Const SOFFICEINDEX = 3  'Sales Office control/field
Const SCODEINDEX = 4  'Station salesperson code control/field
Const PHONEINDEX = 5    'Phone/extension control/field
Const FAXINDEX = 6      'Fax control/field
Const TITLEINDEX = 7
Const TEAMINDEX = 8     'Team control/field
Const STATEINDEX = 9
Const GOALINDEX = 10      'Commission control/field
Const UNDERINDEX = 11
Const OVERINDEX = 12
Const REMUNDERINDEX = 13
Const REMOVERINDEX = 14
Const NEWCLIENTINDEX = 15
Const INCCLIENTINDEX = 16
Const COMMPAIDINDEX = 17     'Start Commission Paid
Const COMMSALESINDEX = 18       'Start Sales
Const COMMDATEINDEX = 19      'Start Commission Date
Private Sub cbcSelect_Change()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim ilPos As Integer
    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    If ilRet = 0 Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cbcSelectErr
        End If
        If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            lacCode.Caption = str$(tmSlf.iCode)
        Else
            lacCode.Caption = ""
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
        lacCode.Caption = ""
    End If
    pbcSlfID.Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            If igSlfFirstNameFirst Then
                ilPos = InStr(slStr, " ")
                If ilPos > 0 Then
                    edcFirstName.Text = Left$(slStr, ilPos - 1)
                    edcLastName.Text = right$(slStr, Len(slStr) - ilPos)
                Else
                    edcFirstName.Text = slStr
                End If
            Else
                ilPos = InStr(slStr, ",")
                If ilPos > 0 Then
                    edcLastName.Text = Left$(slStr, ilPos - 1)
                    edcFirstName.Text = Trim$(right$(slStr, Len(slStr) - ilPos))
                Else
                    edcLastName.Text = slStr
                End If
            End If
            'Set change as imBoxNo is not set
        End If
    End If
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcSlfID_Paint
    Screen.MousePointer = vbDefault
    imChgMode = False
'    mSetCommands
    imBypassSetting = False
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_DropDown()
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        If igSlfCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgSlfName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgSlfName    'New name
            End If
            cbcSelect_Change
            If sgSlfName <> "" Then
                mSetCommands
                gFindMatch sgSlfName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
            End If
            If pbcSTab.Enabled Then
                pbcSTab.SetFocus
            Else
                cmcCancel.SetFocus
            End If
            Exit Sub
        End If
    End If
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields 'Make sure all fields cleared
        If pbcSTab.Enabled Then
            pbcSTab.SetFocus
        Else
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    gCtrlGotFocus cbcSelect
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcSelect.ListIndex = 0
        cbcSelect_Change    'Call change so picture area repainted
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                cbcSelect_Change    'Call change so picture area repainted
                imPopReqd = False
            End If
        Else
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            cbcSelect_Change    'Call change so picture area repainted
        End If
    End If
End Sub
Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcCommDate.SelStart = 0
    edcCommDate.SelLength = Len(edcCommDate.Text)
    edcCommDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcCommDate.SelStart = 0
    edcCommDate.SelLength = Len(edcCommDate.Text)
    edcCommDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    If igSlfCallSource <> CALLNONE Then
        igSlfCallSource = CALLCANCELLED
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcCombo_Click()
    'Screen.MousePointer = vbHourGlass  'Wait
    'sgVsfCallType = "S"
'    Combo.Show vbModal
    'Screen.MousePointer = vbDefault
End Sub
Private Sub cmcCombo_GotFocus()
    gCtrlGotFocus cmcCombo
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    imChgSaveFlag = False
    If igSlfCallSource <> CALLNONE Then
        If igSlfFirstNameFirst Then
            sgSlfName = Trim$(edcFirstName.Text) & " " & Trim$(edcLastName.Text)
        Else
            sgSlfName = Trim$(edcLastName.Text) & ", " & Trim$(edcFirstName.Text)
        End If
        If mSaveRecChg(False) = False Then
            sgSlfName = "[New]"
            If Not imTerminate Then
                mEnableBox imBoxNo
                Exit Sub
            Else
                cmcCancel_Click
                Exit Sub
            End If
        End If
    Else
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    If imChgSaveFlag Then
        sgSalespersonTag = ""
        sgMSlfStamp = ""
        mPopulate
    End If
    If igSlfCallSource <> CALLNONE Then
        If sgSlfName = "[New]" Then
            igSlfCallSource = CALLCANCELLED
        Else
            igSlfCallSource = CALLDONE
        End If
        mTerminate
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilLoop As Integer
    If imBoxNo = -1 Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If Not cmcUpdate.Enabled Then
        'Cycle to first unanswered mandatory
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            If mTestFields(ilLoop, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                imBoxNo = ilLoop
                mEnableBox imBoxNo
                Exit Sub
            End If
        Next ilLoop
    End If
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case SOFFICEINDEX
            lbcSOffice.Visible = Not lbcSOffice.Visible
        Case TITLEINDEX
            lbcTitle.Visible = Not lbcTitle.Visible
        Case TEAMINDEX
            lbcTeam.Visible = Not lbcTeam.Visible
        Case COMMDATEINDEX 'Startup Commission Date
            plcCalendar.Visible = Not plcCalendar.Visible
            edcCommDate.SelStart = 0
            edcCommDate.SelLength = Len(edcCommDate.Text)
            edcCommDate.SetFocus
            Exit Sub
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim slStamp As String   'Date/Time stamp for file
    Dim slMsg As String
    Dim ilCode As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilLoop As Integer
    
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        If tgSpf.sRemoteUsers = "Y" Then
            slMsg = "Cannot erase - Remote User System in Use"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        ilCode = tmSlf.iCode
        'Check that record is not referenced-Code missing
        ilRet = gIICodeRefExist(SPerson, ilCode, "Adf.Btr", "AdfSlfCode") 'adfslfCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Advertiser references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(SPerson, ilCode, "Agf.Btr", "AgfSlfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Agency references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(SPerson, ilCode, "Bsf.Btr", "BsfSlfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Budget by Salesperson references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        For ilLoop = 1 To 10 Step 1
            ilRet = gIICodeRefExist(SPerson, ilCode, "Chf.Btr", "ChfSlfCode" & Trim$(str$(ilLoop)))
            If ilRet Then
                Screen.MousePointer = vbDefault
                slMsg = "Cannot erase - an Agency references name"
                ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
        Next ilLoop
        ilRet = gIICodeRefExist(SPerson, ilCode, "Pjf.Btr", "PjfSlfCode") 'chfslfCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Contract Projection references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gSlfCodeExistInChf(SPerson, ilCode)     ', "Chf.Btr", "ChfSlfCode1") 'chfslfCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Contract references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(SPerson, ilCode, "Rvf.Btr", "RvfSlfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Receivables references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(SPerson, ilCode, "Phf.Btr", "PhfSlfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Payment History references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(SPerson, ilCode, "Scf.Btr", "ScfSlfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Sales Commission references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
'        ilRet = gIICodeRefExist(SPerson, ilCode, "Urf.Btr", "UrfSlfCode")
        ilRet = gCodeInUser(SPerson, "S", ilCode)
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a User option references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIFSCodeRefExist(SPerson, "S", ilCode)
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Salesperson Combo references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmSlf.sLastName & ", " & tmSlf.sFirstName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        gGetSyncDateTime slSyncDate, slSyncTime
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        slStamp = gFileDateTime(sgDBPath & "Slf.btr")
        ilRet = btrDelete(hmSlf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", SPerson
        On Error GoTo 0
'        If tgSpf.sRemoteUsers = "Y" Then
'            tmDsf.lCode = 0
'            tmDsf.sFileName = "SLF"
'            gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'            gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'            tmDsf.iRemoteID = tmSlf.iRemoteID
'            tmDsf.lAutoCode = tmSlf.iAutoCode
'            tmDsf.lCntrNo = 0
'            ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'            On Error GoTo cmcEraseErr
'            gBtrvErrorMsg ilRet, "cmcErase_Click (btrInsert)", SPerson
'            On Error GoTo 0
'        End If
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If Traffic!lbcSalesperson.Tag <> "" Then
        '    If slStamp = Traffic!lbcSalesperson.Tag Then
        '        Traffic!lbcSalesperson.Tag = FileDateTime(sgDBPath & "Slf.btr")
        '    End If
        'End If
        If sgSalespersonTag <> "" Then
            If slStamp = sgSalespersonTag Then
                sgSalespersonTag = gFileDateTime(sgDBPath & "Slf.btr")
            End If
        End If
        'Traffic!lbcSalesperson.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tgSalesperson()
        cbcSelect.RemoveItem imSelectedIndex
    End If
    'Remove focus from control and make invisible
    mSetShow imBoxNo
    imBoxNo = -1
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcSlfID.Cls
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcEraseErr:
    On Error GoTo 0
    imTerminate = True
    Resume Next
End Sub
Private Sub cmcErase_GotFocus()
    gCtrlGotFocus cmcErase
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcMerge_Click()
    Dim slMsg As String
    Dim ilRet As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    'If tgSpf.sRemoteUsers = "Y" Then
    If tgUrf(0).iRemoteID > 0 Then
        slMsg = "Remote User Cannot Run Merge"
        ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Merge")
        Exit Sub
    End If
    ilRet = MsgBox("Backup of database must be done before merge, has it been done", vbYesNo + vbQuestion, "Merge Salespeople")
    If ilRet = vbNo Then
        Exit Sub
    End If
    ilRet = MsgBox("Are all other users off the traffic system", vbYesNo + vbQuestion, "Merge Salespeople")
    If ilRet = vbNo Then
        Exit Sub
    End If
    igMergeCallSource = SALESPEOPLELIST
    Merge.Show vbModal
    Screen.MousePointer = vbHourglass
    pbcSlfID.Cls
    cbcSelect.Clear
    mPopulate
    cbcSelect.ListIndex = 0
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcMerge_GotFocus()
    gCtrlGotFocus cmcMerge
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = SALESPEOPLELIST
    igRptType = 0
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "SPerson^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        Else
            slStr = "SPerson^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "SPerson^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    Else
    '        slStr = "SPerson^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'SPerson.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'SPerson.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    ''Screen.MousePointer = vbDefault    'Default
    sgCommandStr = slStr
    RptList.Show vbModal
End Sub
Private Sub cmcReport_GotFocus()
    gCtrlGotFocus cmcReport
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcUndo_Click()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    ilIndex = imSelectedIndex
    If ilIndex > 0 Then
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cmcUndoErr
        End If
        pbcSlfID.Cls
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            mSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcSlfID_Paint
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcSlfID.Cls
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcUndoErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub cmcUndo_GotFocus()
    gCtrlGotFocus cmcUndo
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcUpdate_Click()
    Dim imSvSelectedIndex As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
'    If igSlfFirstNameFirst Then
'        slName = Trim$(edcFirstName.Text) & " " & Trim$(edcLastName.Text)
'    Else
'        slName = Trim$(edcLastName.Text) & ", " & Trim$(edcFirstName.Text)
'    End If
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    imBoxNo = -1
'    'Must reset display so altered flag is cleared and setcommand will turn select on
'    If imSvSelectedIndex <> 0 Then
'        cbcSelect.Text = slName
'    Else
'        cbcSelect.ListIndex = 0
'    End If
'    cbcSelect_Change    'Call change so picture area repainted
    ilCode = tmSlf.iCode
    cbcSelect.Clear
    sgSalespersonTag = ""
    sgMSlfStamp = ""
    mPopulate
    If imSvSelectedIndex <> 0 Then
        For ilLoop = 0 To UBound(tgSalesperson) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
            slNameCode = tgSalesperson(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = ilCode Then
                If cbcSelect.ListIndex = ilLoop + 1 Then
                    cbcSelect_Change
                Else
                    cbcSelect.ListIndex = ilLoop + 1
                End If
                Exit For
            End If
        Next ilLoop
    Else
        cbcSelect.ListIndex = 0
    End If
    mSetCommands
    cbcSelect.SetFocus
End Sub
Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus cmcUpdate
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub edcCommDate_Change()
    Dim slStr As String
    slStr = edcCommDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcCommDate_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcCommDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcCommDate_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcCommDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcCommDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcCommDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcCommDate.Text = slDate
            End If
        End If
        edcCommDate.SelStart = 0
        edcCommDate.SelLength = Len(edcCommDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcCommDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcCommDate.Text = slDate
            End If
        End If
        edcCommDate.SelStart = 0
        edcCommDate.SelLength = Len(edcCommDate.Text)
    End If
End Sub
Private Sub edcCommPaid_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcCommPaid_GotFocus()
    gCtrlGotFocus edcCommPaid
End Sub
Private Sub edcCommPaid_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcCommPaid.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcCommPaid.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcCommPaid.Text
    slStr = Left$(slStr, edcCommPaid.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCommPaid.SelStart - edcCommPaid.SelLength)
    If gCompNumberStr(slStr, "99999999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo
        Case SOFFICEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcSOffice, imBSMode, slStr)
            If ilRet = 1 Then
                lbcSOffice.ListIndex = 0
            End If
        Case TITLEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcTitle, imBSMode, imComboBoxIndex
        Case TEAMINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcTeam, imBSMode, slStr)
            If ilRet = 1 Then
                lbcTeam.ListIndex = 1
            End If
    End Select
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case SOFFICEINDEX
            If lbcSOffice.ListCount = 1 Then
                lbcSOffice.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case TITLEINDEX
        Case TEAMINDEX
            If lbcTeam.ListCount = 1 Then
                lbcTeam.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case SOFFICEINDEX
                gProcessArrowKey Shift, KeyCode, lbcSOffice, imLbcArrowSetting
            Case TITLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcTitle, imLbcArrowSetting
            Case TEAMINDEX
                gProcessArrowKey Shift, KeyCode, lbcTeam, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case SOFFICEINDEX, TEAMINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
    End If
End Sub
Private Sub edcFirstName_Change()
    mSetChg FIRSTNAMEINDEX   'can't use imBoxNo as not set when edcProd set via cbcSelect- altered flag set so field is saved
End Sub
Private Sub edcFirstName_GotFocus()
    gCtrlGotFocus edcFirstName
End Sub
Private Sub edcFirstName_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    '2/3/16: Disallow forward slash
    'If Not gCheckKeyAscii(ilKey) Then
    If Not gCheckKeyAsciiIncludeSlash(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcFirstName_LostFocus()
    '9760
    edcFirstName.Text = gRemoveIllegalPastedChar(edcFirstName.Text)
    'Dim ilRet As Integer
    ''Test if name changed and if new name is valid
    'ilRet = mOKName()
End Sub
Private Sub edcGoal_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcGoal_GotFocus()
    gCtrlGotFocus edcGoal
End Sub
Private Sub edcGoal_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcGoal.Text
    slStr = Left$(slStr, edcGoal.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcGoal.SelStart - edcGoal.SelLength)
    If gCompNumberStr(slStr, "99999999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcIncClientComm_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcIncClientComm_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcIncClientComm_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcIncClientComm.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcIncClientComm.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcIncClientComm.Text
    slStr = Left$(slStr, edcIncClientComm.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcIncClientComm.SelStart - edcIncClientComm.SelLength)
    If gCompNumberStr(slStr, "99.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcLastName_Change()
    mSetChg LASTNAMEINDEX   'can't use imBoxNo as not set when edcProd set via cbcSelect- altered flag set so field is saved
End Sub
Private Sub edcLastName_GotFocus()
    gCtrlGotFocus edcLastName
End Sub
Private Sub edcLastName_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    '2/3/16: Disallow forward slash
    'If Not gCheckKeyAscii(ilKey) Then
    If Not gCheckKeyAsciiIncludeSlash(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLastName_LostFocus()
    '9760
    edcLastName.Text = gRemoveIllegalPastedChar(edcLastName.Text)
    'Dim ilRet As Integer
    ''Test if name changed and if new name is valid
    'ilRet = mOKName()
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcNewClientComm_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcNewClientComm_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcNewClientComm_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcNewClientComm.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcNewClientComm.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcNewClientComm.Text
    slStr = Left$(slStr, edcNewClientComm.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcNewClientComm.SelStart - edcNewClientComm.SelLength)
    If gCompNumberStr(slStr, "99.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcOverComm_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcOverComm_GotFocus()
    gCtrlGotFocus edcOverComm
End Sub
Private Sub edcOverComm_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcOverComm.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcOverComm.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcOverComm.Text
    slStr = Left$(slStr, edcOverComm.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcOverComm.SelStart - edcOverComm.SelLength)
    If gCompNumberStr(slStr, "99.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRemOverComm_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcRemOverComm_GotFocus()
    gCtrlGotFocus edcRemOverComm
End Sub
Private Sub edcRemOverComm_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcRemOverComm.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcRemOverComm.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcRemOverComm.Text
    slStr = Left$(slStr, edcRemOverComm.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcRemOverComm.SelStart - edcRemOverComm.SelLength)
    If gCompNumberStr(slStr, "99.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRemUnderComm_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcRemUnderComm_GotFocus()
    gCtrlGotFocus edcRemUnderComm
End Sub
Private Sub edcRemUnderComm_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcRemUnderComm.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcRemUnderComm.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcRemUnderComm.Text
    slStr = Left$(slStr, edcRemUnderComm.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcRemUnderComm.SelStart - edcRemUnderComm.SelLength)
    If gCompNumberStr(slStr, "99.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcSales_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcSales_GotFocus()
    gCtrlGotFocus edcSales
End Sub
Private Sub edcSales_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcSales.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcSales.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcSales.Text
    slStr = Left$(slStr, edcSales.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSales.SelStart - edcSales.SelLength)
    If gCompNumberStr(slStr, "99999999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcStationCode_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcStationCode_GotFocus()
    gCtrlGotFocus edcStationCode
End Sub
Private Sub edcStationCode_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcUnderComm_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcUnderComm_GotFocus()
    gCtrlGotFocus edcUnderComm
End Sub
Private Sub edcUnderComm_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcUnderComm.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcUnderComm.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcUnderComm.Text
    slStr = Left$(slStr, edcUnderComm.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcUnderComm.SelStart - edcUnderComm.SelLength)
    If gCompNumberStr(slStr, "99.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(SALESPEOPLELIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcSlfID.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcSlfID.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
'    gShowBranner
    gShowBranner imUpdateAllowed
    mSetCommands
    Me.KeyPreview = True
    SPerson.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        If (cbcSelect.Enabled) And (imBoxNo > 0) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            cbcSelect.Enabled = True
        End If
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
    
    If Not igManUnload Then
        mSetShow imBoxNo
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            mEnableBox imBoxNo
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    Erase tmSOfficeCode
    Erase tmTeamCode
    btrExtClear hmSlf   'Clear any previous extend operation
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
    
    Set SPerson = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcSOffice_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcSOffice, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcSOffice_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcSOffice_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcSOffice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcSOffice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcSOffice, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcTeam_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcTeam, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcTeam_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcTeam_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcTeam_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcTeam_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcTeam, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcTitle_Click()
    gProcessLbcClick lbcTitle, edcDropDown, imChgMode, imLbcArrowSetting
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
    slStr = edcCommDate.Text
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
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear each control on the      *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim ilLoop As Integer
    edcFirstName.Text = ""
    edcLastName.Text = ""
    lbcSOffice.ListIndex = -1
    smSOffice = ""
    edcStationCode.Text = ""
    mkcPhone.Text = smPhoneImage
    mkcFax.Text = smFaxImage
    lbcTeam.ListIndex = -1
    lbcTitle.ListIndex = -1
    smTeam = ""
    imState = -1
    edcGoal.Text = ""
    edcUnderComm.Text = ""
    edcOverComm.Text = ""
    edcRemUnderComm.Text = ""
    edcRemOverComm.Text = ""
    edcNewClientComm.Text = ""
    edcIncClientComm.Text = ""
    edcCommPaid.Text = ""
    edcSales.Text = ""
    edcCommDate.Text = ""
    mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case FIRSTNAMEINDEX 'Name
            edcFirstName.Width = tmCtrls(ilBoxNo).fBoxW
            edcFirstName.MaxLength = 20
            gMoveFormCtrl pbcSlfID, edcFirstName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcFirstName.Visible = True  'Set visibility
            edcFirstName.SetFocus
        Case LASTNAMEINDEX 'Name
            edcLastName.Width = tmCtrls(ilBoxNo).fBoxW
            edcLastName.MaxLength = 20
            gMoveFormCtrl pbcSlfID, edcLastName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcLastName.Visible = True  'Set visibility
            edcLastName.SetFocus
        Case SOFFICEINDEX   'Sales Office
            mSOfficePop
            If imTerminate Then
                Exit Sub
            End If
            lbcSOffice.Height = gListBoxHeight(lbcSOffice.ListCount, 12)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 41
            gMoveFormCtrl pbcSlfID, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcSOffice.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcSOffice.ListIndex < 0 Then
                If lbcSOffice.ListCount <= 1 Then
                    lbcSOffice.ListIndex = 0   '[New]
                Else
                    lbcSOffice.ListIndex = 1
                End If
            End If
            If lbcSOffice.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcSOffice.List(lbcSOffice.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case SCODEINDEX 'Name
            edcStationCode.Width = tmCtrls(ilBoxNo).fBoxW
            edcStationCode.MaxLength = 5
            gMoveFormCtrl pbcSlfID, edcStationCode, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcStationCode.Visible = True  'Set visibility
            edcStationCode.SetFocus
        Case PHONEINDEX 'Phone and extension
            mkcPhone.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSlfID, mkcPhone, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            mkcPhone.Visible = True  'Set visibility
            mkcPhone.SetFocus
        Case FAXINDEX 'Fax
            mkcFax.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSlfID, mkcFax, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            mkcFax.Visible = True  'Set visibility
            mkcFax.SetFocus
        Case TITLEINDEX
            lbcTitle.Height = gListBoxHeight(lbcTitle.ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 11
            gMoveFormCtrl pbcSlfID, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcTitle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcTitle.ListIndex < 0 Then
                lbcTitle.ListIndex = 0   'Salesperson
            End If
            imComboBoxIndex = lbcTitle.ListIndex
            If lbcTitle.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcTitle.List(lbcTitle.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TEAMINDEX   'Sales Team
            mTeamPop
            If imTerminate Then
                Exit Sub
            End If
            lbcTeam.Height = gListBoxHeight(lbcTeam.ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcSlfID, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcTeam.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcTeam.ListIndex < 0 Then
                lbcTeam.ListIndex = 1   '[None]
            End If
            If lbcTeam.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcTeam.List(lbcTeam.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case STATEINDEX   'Active/Dormant
            If imState < 0 Then
                imState = 0    'Active
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcState.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSlfID, pbcState, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcState_Paint
            pbcState.Visible = True
            pbcState.SetFocus
        Case GOALINDEX 'Sales goal
            edcGoal.Width = tmCtrls(ilBoxNo).fBoxW
            edcGoal.MaxLength = 9
            gMoveFormCtrl pbcSlfID, edcGoal, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcGoal.Visible = True  'Set visibility
            edcGoal.SetFocus
        Case UNDERINDEX 'Sales commission
            edcUnderComm.Width = tmCtrls(ilBoxNo).fBoxW
            edcUnderComm.MaxLength = 5
            gMoveFormCtrl pbcSlfID, edcUnderComm, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcUnderComm.Visible = True  'Set visibility
            edcUnderComm.SetFocus
        Case OVERINDEX 'Sales commission
            edcOverComm.Width = tmCtrls(ilBoxNo).fBoxW
            edcOverComm.MaxLength = 5
            gMoveFormCtrl pbcSlfID, edcOverComm, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcOverComm.Visible = True  'Set visibility
            edcOverComm.SetFocus
        Case REMUNDERINDEX 'Sales commission
            edcRemUnderComm.Width = tmCtrls(ilBoxNo).fBoxW
            edcRemUnderComm.MaxLength = 5
            gMoveFormCtrl pbcSlfID, edcRemUnderComm, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcRemUnderComm.Visible = True  'Set visibility
            edcRemUnderComm.SetFocus
        Case REMOVERINDEX 'Sales commission
            edcRemOverComm.Width = tmCtrls(ilBoxNo).fBoxW
            edcRemOverComm.MaxLength = 5
            gMoveFormCtrl pbcSlfID, edcRemOverComm, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcRemOverComm.Visible = True  'Set visibility
            edcRemOverComm.SetFocus
        Case NEWCLIENTINDEX 'Sales commission
            edcNewClientComm.Width = tmCtrls(ilBoxNo).fBoxW
            edcNewClientComm.MaxLength = 5
            gMoveFormCtrl pbcSlfID, edcNewClientComm, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcNewClientComm.Visible = True  'Set visibility
            edcNewClientComm.SetFocus
        Case INCCLIENTINDEX 'Sales commission
            edcIncClientComm.Width = tmCtrls(ilBoxNo).fBoxW
            edcIncClientComm.MaxLength = 5
            gMoveFormCtrl pbcSlfID, edcIncClientComm, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcIncClientComm.Visible = True  'Set visibility
            edcIncClientComm.SetFocus
        Case COMMPAIDINDEX 'Startup Commission Paid
            edcCommPaid.Width = tmCtrls(ilBoxNo).fBoxW
            edcCommPaid.MaxLength = 11
            gMoveFormCtrl pbcSlfID, edcCommPaid, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCommPaid.Visible = True  'Set visibility
            edcCommPaid.SetFocus
        Case COMMSALESINDEX 'Startup Sales
            edcSales.Width = tmCtrls(ilBoxNo).fBoxW
            edcSales.MaxLength = 11
            gMoveFormCtrl pbcSlfID, edcSales, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcSales.Visible = True  'Set visibility
            edcSales.SetFocus
        Case COMMDATEINDEX 'Startup Commission Date
            edcCommDate.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcCommDate.MaxLength = 10
            gMoveFormCtrl pbcSlfID, edcCommDate, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcCommDate.Left + edcCommDate.Width, edcCommDate.Top
            plcCalendar.Move edcCommDate.Left, edcCommDate.Top - plcCalendar.Height
            edcCommDate.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcCommDate.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInitParameters
'   Where:
'
    Dim ilRet As Integer    'Return Status
    Dim slStr As String
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imLBCDCtrls = 1
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    SPerson.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    gCenterStdAlone SPerson
    'SPerson.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    imSlfRecLen = Len(tmSlf)  'Get and save Slf record length
    imBoxNo = -1 'Initialize current Box to N/A
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imCalType = 0   'Standard
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    imBypassFocus = False
    'imShowHelpMsg = True
    imComboBoxIndex = -1
    smPhoneImage = mkcPhone.Text
    smFaxImage = mkcFax.Text
    hmSlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Slf.btr)", SPerson
    On Error GoTo 0
'    hmDsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dsf.btr)", SPerson
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf)  'Get and save Dsf record length
    lbcTitle.AddItem "Salesperson"
    lbcTitle.AddItem "Negotiator"
    lbcTitle.AddItem "Planner"
    lbcTitle.AddItem "Manager"
    lbcSOffice.Clear 'Force list box to be populated
    mSOfficePop
    If imTerminate Then
        Exit Sub
    End If
    lbcTeam.Clear 'Force list box to be populated
    mTeamPop
    If imTerminate Then
        Exit Sub
    End If
'    gCenterModalForm SPerson
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0 'This will generate a select_change event
        mSetCommands
    End If
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    lacDate.Visible = False
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
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
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
    flTextHeight = pbcSlfID.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcSlfID.Move 630, 615, pbcSlfID.Width + fgPanelAdj, pbcSlfID.Height + fgPanelAdj
    pbcSlfID.Move plcSlfID.Left + fgBevelX, plcSlfID.Top + fgBevelY
    'Position panel and picture areas with panel
    'First Name
    gSetCtrl tmCtrls(FIRSTNAMEINDEX), 30, 30, 2805, fgBoxStH
    'Last Name
    gSetCtrl tmCtrls(LASTNAMEINDEX), 2850, tmCtrls(FIRSTNAMEINDEX).fBoxY, 2805, fgBoxStH
    'Sales Office
    gSetCtrl tmCtrls(SOFFICEINDEX), 30, tmCtrls(FIRSTNAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    'Station codes
    gSetCtrl tmCtrls(SCODEINDEX), 2850, tmCtrls(SOFFICEINDEX).fBoxY, 2805, fgBoxStH
    tmCtrls(SCODEINDEX).iReq = False
    'Phone
    gSetCtrl tmCtrls(PHONEINDEX), 30, tmCtrls(SOFFICEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    tmCtrls(PHONEINDEX).iReq = False
    'Fax
    gSetCtrl tmCtrls(FAXINDEX), 2850, tmCtrls(PHONEINDEX).fBoxY, 1470, fgBoxStH
    tmCtrls(FAXINDEX).iReq = False
    'Title
    gSetCtrl tmCtrls(TITLEINDEX), 30, tmCtrls(PHONEINDEX).fBoxY + fgStDeltaY, 1860, fgBoxStH
    tmCtrls(TITLEINDEX).iReq = False
    'Team
    gSetCtrl tmCtrls(TEAMINDEX), 1905, tmCtrls(TITLEINDEX).fBoxY, 1875, fgBoxStH
    tmCtrls(TEAMINDEX).iReq = False
    'State
    gSetCtrl tmCtrls(STATEINDEX), 3795, tmCtrls(TITLEINDEX).fBoxY, 1860, fgBoxStH
    'Goal
    gSetCtrl tmCtrls(GOALINDEX), 30, tmCtrls(TITLEINDEX).fBoxY + fgStDeltaY, 1860, fgBoxStH
    tmCtrls(GOALINDEX).iReq = False
    'Under
    gSetCtrl tmCtrls(UNDERINDEX), 1905, tmCtrls(GOALINDEX).fBoxY, 1875, fgBoxStH
    tmCtrls(UNDERINDEX).iReq = False
    'Over
    gSetCtrl tmCtrls(OVERINDEX), 3795, tmCtrls(GOALINDEX).fBoxY, 1860, fgBoxStH
    tmCtrls(OVERINDEX).iReq = False
    'Remnant Under Commission
    gSetCtrl tmCtrls(REMUNDERINDEX), 30, tmCtrls(GOALINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    tmCtrls(REMUNDERINDEX).iReq = False
    'Remnant Over Commission
    gSetCtrl tmCtrls(REMOVERINDEX), 2850, tmCtrls(REMUNDERINDEX).fBoxY, 2805, fgBoxStH
    tmCtrls(REMOVERINDEX).iReq = False
    'New Client Commission
    gSetCtrl tmCtrls(NEWCLIENTINDEX), 30, tmCtrls(REMUNDERINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    tmCtrls(NEWCLIENTINDEX).iReq = False
    'Increased Client Commission
    gSetCtrl tmCtrls(INCCLIENTINDEX), 2850, tmCtrls(NEWCLIENTINDEX).fBoxY, 2805, fgBoxStH
    tmCtrls(INCCLIENTINDEX).iReq = False
    'Startup Commission Paid
    gSetCtrl tmCtrls(COMMPAIDINDEX), 30, tmCtrls(NEWCLIENTINDEX).fBoxY + fgStDeltaY, 1860, fgBoxStH
    tmCtrls(COMMPAIDINDEX).iReq = False
    'Start Sales
    gSetCtrl tmCtrls(COMMSALESINDEX), 1905, tmCtrls(COMMPAIDINDEX).fBoxY, 1875, fgBoxStH
    tmCtrls(COMMSALESINDEX).iReq = False
    'Startup Commission date
    gSetCtrl tmCtrls(COMMDATEINDEX), 3795, tmCtrls(COMMPAIDINDEX).fBoxY, 1860, fgBoxStH
    tmCtrls(COMMDATEINDEX).iReq = False
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
Private Sub mkcFax_Change()
    mSetChg imBoxNo
End Sub
Private Sub mkcFax_GotFocus()
    gCtrlGotFocus mkcFax
End Sub
Private Sub mkcPhone_Change()
    mSetChg imBoxNo
End Sub
Private Sub mkcPhone_GotFocus()
    gCtrlGotFocus mkcPhone
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    Dim slNameCode As String  'Name and code
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'Code number
    Dim slStr As String
    If Not ilTestChg Or tmCtrls(FIRSTNAMEINDEX).iChg Then
        tmSlf.sFirstName = edcFirstName.Text
    End If
    If Not ilTestChg Or tmCtrls(LASTNAMEINDEX).iChg Then
        tmSlf.sLastName = edcLastName.Text
    End If
    If Not ilTestChg Or tmCtrls(SOFFICEINDEX).iChg Then
        If lbcSOffice.ListIndex >= 1 Then
            slNameCode = tmSOfficeCode(lbcSOffice.ListIndex - 1).sKey  'lbcSOfficeCode.List(lbcSOffice.ListIndex - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", SPerson
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmSlf.iSofCode = CInt(slCode)
        Else
            tmSlf.iSofCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(SCODEINDEX).iChg Then
        tmSlf.sCodeStn = edcStationCode.Text
    End If
    If Not ilTestChg Or tmCtrls(PHONEINDEX).iChg Then
        gGetPhoneNo mkcPhone, tmSlf.sPhone
    End If
    If Not ilTestChg Or tmCtrls(FAXINDEX).iChg Then
        gGetPhoneNo mkcFax, tmSlf.sFax
    End If
    If Not ilTestChg Or tmCtrls(TITLEINDEX).iChg Then
        Select Case lbcTitle.ListIndex
            Case 0  'Salesperson
                tmSlf.sJobTitle = "S"
            Case 1  'Negotiator
                tmSlf.sJobTitle = "N"
            Case 2  'Planner
                tmSlf.sJobTitle = "P"
            Case 3  'Manager
                tmSlf.sJobTitle = "M"
            Case Else
                tmSlf.sJobTitle = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(TEAMINDEX).iChg Then
        If lbcTeam.ListIndex >= 2 Then
            slNameCode = tmTeamCode(lbcTeam.ListIndex - 2).sKey    'lbcTeamCode.List(lbcTeam.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", SPerson
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmSlf.iMnfSlsTeam = CInt(slCode)
        Else
            tmSlf.iMnfSlsTeam = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(STATEINDEX).iChg Then
        Select Case imState
            Case 0  'Active
                tmSlf.sState = "A"
            Case 1  'Dormant
                tmSlf.sState = "D"
            Case Else
                tmSlf.sState = "A"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(GOALINDEX).iChg Then
        slStr = edcGoal.Text
        tmSlf.lSalesGoal = gStrDecToLong(slStr, 0)
    End If
    If Not ilTestChg Or tmCtrls(UNDERINDEX).iChg Then
        slStr = edcUnderComm.Text
        tmSlf.iUnderComm = gStrDecToInt(slStr, 2)
    End If
    If Not ilTestChg Or tmCtrls(OVERINDEX).iChg Then
        slStr = edcOverComm.Text
        tmSlf.iOverComm = gStrDecToInt(slStr, 2)
    End If
    If Not ilTestChg Or tmCtrls(REMUNDERINDEX).iChg Then
        slStr = edcRemUnderComm.Text
        tmSlf.iRemUnderComm = gStrDecToInt(slStr, 2)
    End If
    If Not ilTestChg Or tmCtrls(REMOVERINDEX).iChg Then
        slStr = edcRemOverComm.Text
        tmSlf.iRemOverComm = gStrDecToInt(slStr, 2)
    End If
    If Not ilTestChg Or tmCtrls(NEWCLIENTINDEX).iChg Then
        slStr = edcNewClientComm.Text
        tmSlf.iNewClientComm = gStrDecToInt(slStr, 2)
    End If
    If Not ilTestChg Or tmCtrls(INCCLIENTINDEX).iChg Then
        slStr = edcIncClientComm.Text
        tmSlf.iIncClientComm = gStrDecToInt(slStr, 2)
    End If
    If Not ilTestChg Or tmCtrls(COMMPAIDINDEX).iChg Then
        slStr = edcCommPaid.Text
        tmSlf.lStartCommPaid = gStrDecToLong(slStr, 2)
    End If
    If Not ilTestChg Or tmCtrls(COMMSALESINDEX).iChg Then
        slStr = edcSales.Text
        tmSlf.lStartSales = gStrDecToLong(slStr, 2)
    End If
    If Not ilTestChg Or tmCtrls(COMMDATEINDEX).iChg Then
        slStr = edcCommDate.Text
        If gValidDate(slStr) Then
            gPackDate slStr, tmSlf.iStartCommDate(0), tmSlf.iStartCommDate(1)
        End If
    End If
    Exit Sub
mMoveCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim slRecCode As String
    Dim slNameCode As String  'Name and code
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'Sales source code number
    Dim slStr As String
    edcFirstName.Text = Trim$(tmSlf.sFirstName)
    edcLastName.Text = Trim$(tmSlf.sLastName)
    'look up sales office name from code number
    lbcSOffice.ListIndex = 0
    smSOffice = ""
    slRecCode = Trim$(str$(tmSlf.iSofCode))
    For ilLoop = 0 To UBound(tmSOfficeCode) - 1 Step 1  'lbcSOfficeCode.ListCount - 1 Step 1
        slNameCode = tmSOfficeCode(ilLoop).sKey    'lbcSOfficeCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", SPerson
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcSOffice.ListIndex = ilLoop + 1
            smSOffice = lbcSOffice.List(ilLoop + 1)
            Exit For
        End If
    Next ilLoop
    edcStationCode.Text = Trim$(tmSlf.sCodeStn)
    gSetPhoneNo tmSlf.sPhone, mkcPhone
    gSetPhoneNo tmSlf.sFax, mkcFax
    Select Case tmSlf.sJobTitle
        Case "S"
            lbcTitle.ListIndex = 0
        Case "N"
            lbcTitle.ListIndex = 1
        Case "P"
            lbcTitle.ListIndex = 2
        Case "M"
            lbcTitle.ListIndex = 3
        Case Else
            lbcTitle.ListIndex = -1
    End Select
    'look up team from code number
    lbcTeam.ListIndex = 1
    smTeam = ""
    slRecCode = Trim$(str$(tmSlf.iMnfSlsTeam))
    For ilLoop = 0 To UBound(tmTeamCode) - 1 Step 1 'lbcTeamCode.ListCount - 1 Step 1
        slNameCode = tmTeamCode(ilLoop).sKey   'lbcTeamCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", SPerson
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcTeam.ListIndex = ilLoop + 2
            smTeam = lbcTeam.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    slStr = gLongToStrDec(tmSlf.lSalesGoal, 0)
    edcGoal.Text = slStr
    slStr = gIntToStrDec(tmSlf.iUnderComm, 2)
    edcUnderComm.Text = slStr
    slStr = gIntToStrDec(tmSlf.iOverComm, 2)
    edcOverComm.Text = slStr
    slStr = gIntToStrDec(tmSlf.iRemUnderComm, 2)
    edcRemUnderComm.Text = slStr
    slStr = gIntToStrDec(tmSlf.iRemOverComm, 2)
    edcRemOverComm.Text = slStr
    slStr = gIntToStrDec(tmSlf.iNewClientComm, 2)
    edcNewClientComm.Text = slStr
    slStr = gIntToStrDec(tmSlf.iIncClientComm, 2)
    edcIncClientComm.Text = slStr
    'gPDNToStr tmSlf.sComm, 4, slStr
    'edcComm.Text = slStr
    'gPDNToStr tmSlf.sDrawAmt, 2, slStr
    'edcDraw.Text = slStr
    If tmSlf.sState = "D" Then
        imState = 1 'Dormant
    Else
        imState = 0 'Active
    End If
    slStr = gLongToStrDec(tmSlf.lStartCommPaid, 2)
    edcCommPaid.Text = slStr
    slStr = gLongToStrDec(tmSlf.lStartSales, 2)
    edcSales.Text = slStr
    gUnpackDate tmSlf.iStartCommDate(0), tmSlf.iStartCommDate(1), slStr
    edcCommDate.Text = slStr
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    Exit Sub
mMoveRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
    Dim slStr As String
    Dim slName As String
    If (edcFirstName.Text <> "") And (edcLastName.Text <> "") Then    'Test name
        If igSlfFirstNameFirst Then
            slName = Trim$(edcFirstName.Text) & " " & Trim$(edcLastName.Text)
        Else
            slName = Trim$(edcLastName.Text) & ", " & Trim$(edcFirstName.Text)
        End If
        gFindMatch slName, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If igSlfFirstNameFirst Then
                    slStr = Trim$(edcFirstName.Text) & " " & Trim$(edcLastName.Text)
                Else
                    slStr = Trim$(edcLastName.Text) & ", " & Trim$(edcFirstName.Text)
                End If
                If slStr = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Salesperson already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcFirstName.Text = Trim$(tmSlf.sFirstName) 'Reset text
                    edcLastName.Text = Trim$(tmSlf.sLastName)
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = FIRSTNAMEINDEX
                    mEnableBox imBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
    End If
    mOKName = True
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
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
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
    '    imShowHelpMsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone SPerson, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igSlfCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igSlfCallSource = CALLNONE
    'End If
    If igSlfCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgSlfName = slStr
        Else
            sgSlfName = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status

    imPopReqd = False
    'ilRet = gPopSalespersonBox(SPerson, 0, True, True, cbcSelect, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(SPerson, 0, True, True, cbcSelect, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", SPerson
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        imPopReqd = True
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer)
'
'   iRet = mReadRec()
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    slNameCode = tgSalesperson(ilSelectIndex - 1).sKey 'Traffic!lbcSalesperson.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", SPerson
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmSlfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", SPerson
    On Error GoTo 0
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim slSyncDate As String
    Dim slSyncTime As String
    mSetShow imBoxNo
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    gGetSyncDateTime slSyncDate, slSyncTime
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Slf.btr")
        'If Len(Traffic!lbcSalesperson.Tag) > Len(slStamp) Then
        '    slStamp = slStamp & Right$(Traffic!lbcSalesperson.Tag, Len(Traffic!lbcSalesperson.Tag) - Len(slStamp))
        'End If
        If Len(sgSalespersonTag) > Len(slStamp) Then
            slStamp = slStamp & right$(sgSalespersonTag, Len(sgSalespersonTag) - Len(slStamp))
        End If
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mMoveCtrlToRec True
        tmSlf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
        If imSelectedIndex = 0 Then 'New selected
            tmSlf.iCode = 0  'Autoincrement
            tmSlf.iRemoteID = tgUrf(0).iRemoteUserID
            tmSlf.iAutoCode = tmSlf.iCode
            ilRet = btrInsert(hmSlf, tmSlf, imSlfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            gPackDate slSyncDate, tmSlf.iSyncDate(0), tmSlf.iSyncDate(1)
            gPackTime slSyncTime, tmSlf.iSyncTime(0), tmSlf.iSyncTime(1)
            ilRet = btrUpdate(hmSlf, tmSlf, imSlfRecLen)
            slMsg = "mSaveRec (btrUpdate)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, SPerson
    On Error GoTo 0
    If imSelectedIndex = 0 Then 'New selected
        Do
            'tmSlfSrchKey.iCode = tmSlf.iCode
            'ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            'slMsg = "mSaveRec (btrGetEqual:Salesperson)"
            'On Error GoTo mSaveRecErr
            'gBtrvErrorMsg ilRet, slMsg, SPerson
            'On Error GoTo 0
            tmSlf.iRemoteID = tgUrf(0).iRemoteUserID
            tmSlf.iAutoCode = tmSlf.iCode
            gPackDate slSyncDate, tmSlf.iSyncDate(0), tmSlf.iSyncDate(1)
            gPackTime slSyncTime, tmSlf.iSyncTime(0), tmSlf.iSyncTime(1)
            ilRet = btrUpdate(hmSlf, tmSlf, imSlfRecLen)
            slMsg = "mSaveRec (btrUpdate:Salesperson)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, SPerson
        On Error GoTo 0
    End If
'    'If Traffic!lbcSalesperson.Tag <> "" Then
'    '    If slStamp = Traffic!lbcSalesperson.Tag Then
'    '        Traffic!lbcSalesperson.Tag = FileDateTime(sgDBPath & "Slf.btr")
'    '        If Len(slStamp) > Len(Traffic!lbcSalesperson.Tag) Then
'    '            Traffic!lbcSalesperson.Tag = Traffic!lbcSalesperson.Tag & Right$(slStamp, Len(slStamp) - Len(Traffic!lbcSalesperson.Tag))
'    '        End If
'    '    End If
'    'End If
'    If sgSalespersonTag <> "" Then
'        If slStamp = sgSalespersonTag Then
'            sgSalespersonTag = gFileDateTime(sgDBPath & "Slf.btr")
'            If Len(slStamp) > Len(sgSalespersonTag) Then
'                sgSalespersonTag = sgSalespersonTag & right$(slStamp, Len(slStamp) - Len(sgSalespersonTag))
'            End If
'        End If
'    End If
'    If imSelectedIndex <> 0 Then
'        'Traffic!lbcSalesperson.RemoveItem imSelectedIndex - 1
'        gRemoveItemFromSortCode imSelectedIndex - 1, tgSalesperson()
'        cbcSelect.RemoveItem imSelectedIndex
'    End If
'    cbcSelect.RemoveItem 0 'Remove [New]
'    If igSlfFirstNameFirst Then
'        slName = Trim$(tmSlf.sFirstName) & " " & Trim$(tmSlf.sLastName)
'    Else
'        slName = Trim$(tmSlf.sLastName) & ", " & Trim$(tmSlf.sFirstName)
'    End If
'    cbcSelect.AddItem slName
'    Do While Len(slName) < Len(tmSlf.sFirstName) + Len(tmSlf.sLastName) + 2
'        slName = slName & " "
'    Loop
'    slName = slName + "\" + LTrim$(Str$(tmSlf.iCode))
'    'Traffic!lbcSalesperson.AddItem slName
'    gAddItemToSortCode slName, tgSalesperson(), True
'    cbcSelect.AddItem "[New]", 0
    imChgSaveFlag = True
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if record altered and*
'*                      requires updating              *
'*                                                     *
'*******************************************************
Private Function mSaveRecChg(ilAsk As Integer) As Integer
'
'   iAsk = True
'   iRet = mSaveRecChg(iAsk)
'   Where:
'       iAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRes As Integer
    Dim slMess As String
    Dim ilAltered As Integer
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If mTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
        If ilAltered = YES Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & edcFirstName.Text & " " & edcLastName.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcSlfID_Paint
                    Exit Function
                End If
                If ilRes = vbYes Then
                    ilRes = mSaveRec()
                    mSaveRecChg = ilRes
                    Exit Function
                End If
                If ilRes = vbNo Then
                    cbcSelect.ListIndex = 0
                End If
            Else
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
        End If
    End If
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mSetChg(ilBoxNo As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim slStr As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case FIRSTNAMEINDEX 'First Name
            gSetChgFlag tmSlf.sFirstName, edcFirstName, tmCtrls(ilBoxNo)
        Case LASTNAMEINDEX 'Last Name
            gSetChgFlag tmSlf.sLastName, edcLastName, tmCtrls(ilBoxNo)
        Case SOFFICEINDEX   'Sales Source
            gSetChgFlag smSOffice, lbcSOffice, tmCtrls(ilBoxNo)
        Case SCODEINDEX 'Station code
            gSetChgFlag tmSlf.sCodeStn, edcStationCode, tmCtrls(ilBoxNo)
        Case PHONEINDEX 'Phone number plus extension
            gSetChgFlag tmSlf.sPhone, mkcPhone, tmCtrls(ilBoxNo)
        Case FAXINDEX 'Fax number
            gSetChgFlag tmSlf.sFax, mkcFax, tmCtrls(ilBoxNo)
        Case TITLEINDEX
            slStr = ""
            Select Case tmSlf.sJobTitle
                Case "S"
                    slStr = lbcTitle.List(0)
                Case "N"
                    slStr = lbcTitle.List(1)
                Case "P"
                    slStr = lbcTitle.List(2)
                Case "M"
                    slStr = lbcTitle.List(3)
            End Select
            gSetChgFlag slStr, lbcTitle, tmCtrls(ilBoxNo)
        Case TEAMINDEX   'Sales Source
            gSetChgFlag smTeam, lbcTeam, tmCtrls(ilBoxNo)
        'Case COMMISSIONINDEX
        '    gPDNToStr tmSlf.sComm, 4, slStr
        '    gSetChgFlag slStr, edcComm, tmCtrls(ilBoxNo)
        Case STATEINDEX
        Case GOALINDEX
            'gPDNToStr tmSlf.sDrawAmt, 2, slStr
            slStr = gLongToStrDec(tmSlf.lSalesGoal, 0)
            gSetChgFlag slStr, edcGoal, tmCtrls(ilBoxNo)
        Case UNDERINDEX
            'gPDNToStr tmSlf.sDrawAmt, 2, slStr
            slStr = gIntToStrDec(tmSlf.iUnderComm, 2)
            gSetChgFlag slStr, edcUnderComm, tmCtrls(ilBoxNo)
        Case OVERINDEX
            'gPDNToStr tmSlf.sDrawAmt, 2, slStr
            slStr = gIntToStrDec(tmSlf.iOverComm, 2)
            gSetChgFlag slStr, edcOverComm, tmCtrls(ilBoxNo)
        Case REMUNDERINDEX
            'gPDNToStr tmSlf.sDrawAmt, 2, slStr
            slStr = gIntToStrDec(tmSlf.iRemUnderComm, 2)
            gSetChgFlag slStr, edcRemUnderComm, tmCtrls(ilBoxNo)
        Case REMOVERINDEX
            'gPDNToStr tmSlf.sDrawAmt, 2, slStr
            slStr = gIntToStrDec(tmSlf.iRemOverComm, 2)
            gSetChgFlag slStr, edcRemOverComm, tmCtrls(ilBoxNo)
        Case NEWCLIENTINDEX
            'gPDNToStr tmSlf.sDrawAmt, 2, slStr
            slStr = gIntToStrDec(tmSlf.iNewClientComm, 2)
            gSetChgFlag slStr, edcNewClientComm, tmCtrls(ilBoxNo)
        Case INCCLIENTINDEX
            'gPDNToStr tmSlf.sDrawAmt, 2, slStr
            slStr = gIntToStrDec(tmSlf.iIncClientComm, 2)
            gSetChgFlag slStr, edcIncClientComm, tmCtrls(ilBoxNo)
        Case COMMPAIDINDEX
            slStr = gLongToStrDec(tmSlf.lStartCommPaid, 2)
            gSetChgFlag slStr, edcCommPaid, tmCtrls(ilBoxNo)
        Case COMMSALESINDEX
            slStr = gLongToStrDec(tmSlf.lStartSales, 2)
            gSetChgFlag slStr, edcSales, tmCtrls(ilBoxNo)
        Case COMMDATEINDEX
            gUnpackDate tmSlf.iStartCommDate(0), tmSlf.iStartCommDate(1), slStr
            gSetChgFlag slStr, edcCommDate, tmCtrls(ilBoxNo)
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    Dim ilAltered As Integer
    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    'Revert button set if any field changed
    If ilAltered Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    'If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) Then
    If ((imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO)) And (tgSpf.sRemoteUsers <> "Y") Then
        If imUpdateAllowed Then
            cmcErase.Enabled = True
        Else
            cmcErase.Enabled = False
        End If
    Else
        cmcErase.Enabled = False
    End If
    'Merge set only if change mode
    'If (imSelectedIndex > 0) And (tgUrf(0).sMerge = "I") Then
    If (Not ilAltered) And (tgUrf(0).sMerge = "I") And (tgUrf(0).iRemoteID = 0) And (imUpdateAllowed) Then
        cmcMerge.Enabled = True
    Else
        cmcMerge.Enabled = False
    End If
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
'    If cbcSelect.ListCount <= 2 Then
'        cmcCombo.Enabled = False
'    Else
'        cmcCombo.Enabled = True
'    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case FIRSTNAMEINDEX 'Name
            edcFirstName.SetFocus
        Case LASTNAMEINDEX 'Name
            edcLastName.SetFocus
        Case SOFFICEINDEX   'Sales Office
            edcDropDown.SetFocus
        Case SCODEINDEX 'Name
            edcStationCode.SetFocus
        Case PHONEINDEX 'Phone and extension
            mkcPhone.SetFocus
        Case FAXINDEX 'Fax
            mkcFax.SetFocus
        Case TITLEINDEX
            edcDropDown.SetFocus
        Case TEAMINDEX   'Sales Team
            edcDropDown.SetFocus
        Case STATEINDEX   'Active/Dormant
            pbcState.SetFocus
        Case GOALINDEX 'Sales Goal
            edcGoal.SetFocus
        Case UNDERINDEX 'Under commission
            edcUnderComm.SetFocus
        Case OVERINDEX 'Over commission
            edcOverComm.SetFocus
        Case REMUNDERINDEX 'Under commission
            edcRemUnderComm.SetFocus
        Case REMOVERINDEX 'Over commission
            edcRemOverComm.SetFocus
        Case NEWCLIENTINDEX 'Under commission
            edcNewClientComm.SetFocus
        Case INCCLIENTINDEX 'Over commission
            edcIncClientComm.SetFocus
        Case COMMPAIDINDEX 'Startup Commission Paid
            edcCommPaid.SetFocus
        Case COMMSALESINDEX 'Startup Sales
            edcSales.SetFocus
        Case COMMDATEINDEX 'Start Commission Date
            edcCommDate.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilPos As Integer
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    '2/4/16: Add filter to handle the case where the name has illegal characters and it was pasted into the field
    If (ilBoxNo = FIRSTNAMEINDEX) Then
        slStr = gReplaceIllegalCharacters(edcFirstName.Text)
        edcFirstName.Text = slStr
    End If
    If (ilBoxNo = LASTNAMEINDEX) Then
        slStr = gReplaceIllegalCharacters(edcLastName.Text)
        edcLastName.Text = slStr
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case FIRSTNAMEINDEX 'Name
            edcFirstName.Visible = False  'Set visibility
            slStr = edcFirstName.Text
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case LASTNAMEINDEX 'Name
            edcLastName.Visible = False  'Set visibility
            slStr = edcLastName.Text
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case SOFFICEINDEX   'Sales office
            lbcSOffice.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcSOffice.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcSOffice.List(lbcSOffice.ListIndex)
            End If
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case SCODEINDEX 'Station code
            edcStationCode.Visible = False  'Set visibility
            slStr = edcStationCode.Text
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case PHONEINDEX 'Phone number plus extension
            mkcPhone.Visible = False  'Set visibility
            If mkcPhone.Text = smPhoneImage Then
                slStr = ""
            Else
                slStr = mkcPhone.Text
            End If
            If slStr <> "" Then
                If InStr(slStr, "(____)") <> 0 Then
                    ilPos = InStr(slStr, "Ext(")
                    slStr = Left$(slStr, ilPos - 1)
                End If
            End If
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case FAXINDEX 'Fax number
            mkcFax.Visible = False  'Set visibility
            If mkcFax.Text = smFaxImage Then
                slStr = ""
            Else
                slStr = mkcFax.Text
            End If
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case TITLEINDEX
            lbcTitle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcTitle.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcTitle.List(lbcTitle.ListIndex)
            End If
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case TEAMINDEX   'Sales team
            lbcTeam.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcTeam.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcTeam.List(lbcTeam.ListIndex)
            End If
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case STATEINDEX   'Active/Dormant
            pbcState.Visible = False  'Set visibility
            If imState = 0 Then
                slStr = "Active"
            ElseIf imState = 1 Then
                slStr = "Dormant"
            Else
                slStr = ""
            End If
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case GOALINDEX
            edcGoal.Visible = False
            slStr = edcGoal.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 0, slStr
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case UNDERINDEX
            edcUnderComm.Visible = False
            slStr = edcUnderComm.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case OVERINDEX
            edcOverComm.Visible = False
            slStr = edcOverComm.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case REMUNDERINDEX
            edcRemUnderComm.Visible = False
            slStr = edcRemUnderComm.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case REMOVERINDEX
            edcRemOverComm.Visible = False
            slStr = edcRemOverComm.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case NEWCLIENTINDEX
            edcNewClientComm.Visible = False
            slStr = edcNewClientComm.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case INCCLIENTINDEX
            edcIncClientComm.Visible = False
            slStr = edcIncClientComm.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case COMMPAIDINDEX
            edcCommPaid.Visible = False
            slStr = edcCommPaid.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 2, slStr
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case COMMSALESINDEX 'Name
            edcSales.Visible = False
            slStr = edcSales.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 2, slStr
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
        Case COMMDATEINDEX
            edcCommDate.Visible = False
            cmcDropDown.Visible = False
            plcCalendar.Visible = False
            slStr = edcCommDate.Text
            If Not gValidDate(slStr) Then
                slStr = ""
            End If
            gSetShow pbcSlfID, slStr, tmCtrls(ilBoxNo)
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSOfficeBranch                  *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to sales  *
'*                      office and process             *
'*                      communication back from sales  *
'*                      office                         *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mSOfficeBranch() As Integer
'
'   ilRet = mSSourceBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    Dim ilPos As Integer
    Dim ilSvPos As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcSOffice, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mSOfficeBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(SALESOFFICESLIST)) Then
    '    imDoubleClickName = False
    '    mSOfficeBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    igSofCallSource = CALLSOURCESALESPERSON
    If edcDropDown.Text = "[New]" Then
        sgSofName = ""
    Else
        ilSvPos = 0
        ilPos = 1
        Do While ilPos > 0
            ilPos = InStr(ilPos, slStr, "/")
            If ilPos = 0 Then
                Exit Do
            End If
            ilSvPos = ilPos
            ilPos = ilPos + 1
        Loop
        If ilSvPos > 0 Then
            sgSofName = Left$(slStr, ilSvPos - 1)
        Else
            sgSofName = ""
        End If
        'ilRet = gParseItem(slStr, 1, "/", sgSofName)
        'If ilRet <> CP_MSG_NONE Then
        '    sgSofName = slStr
        'End If
    End If
    ilUpdateAllowed = imUpdateAllowed

'    SPerson.Enabled = False
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "SPerson^Test\" & sgUserName & "\" & Trim$(str$(igSofCallSource)) & "\" & sgSofName
        Else
            slStr = "SPerson^Prod\" & sgUserName & "\" & Trim$(str$(igSofCallSource)) & "\" & sgSofName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "SPerson^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igSofCallSource)) & "\" & sgSofName
    '    Else
    '        slStr = "SPerson^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igSofCallSource)) & "\" & sgSofName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "SOffice.Exe " & slStr, 1)
    'SPerson.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    SOffice.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgSofName)
    igSofCallSource = Val(sgSofName)
    ilParse = gParseItem(slStr, 2, "\", sgSofName)
    'SPerson.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mSOfficeBranch = True
    imUpdateAllowed = ilUpdateAllowed
'    gShowBranner
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igSofCallSource = CALLDONE Then  'Done
        igSofCallSource = CALLNONE
'        gSetMenuState True
        lbcSOffice.Clear
        smSOfficeCodeTag = ""
        mSOfficePop
        If imTerminate Then
            mSOfficeBranch = False
            Exit Function
        End If
        gFindMatch sgSofName, 1, lbcSOffice
        sgSofName = ""
        If gLastFound(lbcSOffice) > 0 Then
            imChgMode = True
            lbcSOffice.ListIndex = gLastFound(lbcSOffice)
            edcDropDown.Text = lbcSOffice.List(lbcSOffice.ListIndex)
            imChgMode = False
            mSOfficeBranch = False
            mSetChg SOFFICEINDEX
        Else
            imChgMode = True
            lbcSOffice.ListIndex = 0
            edcDropDown.Text = lbcSOffice.List(0)
            imChgMode = False
            mSetChg SOFFICEINDEX
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igSofCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igSofCallSource = CALLNONE
        sgSofName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igSofCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igSofCallSource = CALLNONE
        sgSofName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSOfficePop                     *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Sales office list box *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mSOfficePop()
'
'   mSOfficePop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcSOffice.ListIndex
    If ilIndex > 0 Then
        slName = lbcSOffice.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopOfficeSourceBox(SPerson, lbcSOffice, lbcSOfficeCode)
    ilRet = gPopOfficeSourceBox(SPerson, lbcSOffice, tmSOfficeCode(), smSOfficeCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSOfficePopErr
        gCPErrorMsg ilRet, "mSOfficePop (gIMoveListBox)", SPerson
        On Error GoTo 0
        lbcSOffice.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcSOffice
            If gLastFound(lbcSOffice) > 0 Then
                lbcSOffice.ListIndex = gLastFound(lbcSOffice)
            Else
                lbcSOffice.ListIndex = -1
            End If
        Else
            lbcSOffice.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mSOfficePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTeamBranch                     *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to sales  *
'*                      team and process               *
'*                      communication back from sales  *
'*                      team                           *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mTeamBranch() As Integer
'
'   ilRet = mSSourceBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcDropDown, lbcTeam, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mTeamBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(SALESTEAMSLIST)) Then
    '    imDoubleClickName = False
    '    mTeamBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "T"
    igMNmCallSource = CALLSOURCESALESPERSON
    If edcDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "SPerson^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "SPerson^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "SPerson^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "SPerson^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'SPerson.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'SPerson.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mTeamBranch = True
    imUpdateAllowed = ilUpdateAllowed
'    gShowBranner
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcTeam.Clear
        smTeamCodeTag = ""
        mTeamPop
        If imTerminate Then
            mTeamBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcTeam
        sgMNmName = ""
        If gLastFound(lbcTeam) > 0 Then
            imChgMode = True
            lbcTeam.ListIndex = gLastFound(lbcTeam)
            edcDropDown.Text = lbcTeam.List(lbcTeam.ListIndex)
            imChgMode = False
            mTeamBranch = False
            mSetChg TEAMINDEX
        Else
            imChgMode = True
            lbcTeam.ListIndex = 1
            edcDropDown.Text = lbcTeam.List(1)
            imChgMode = False
            mSetChg TEAMINDEX
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTeamPop                        *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Sales team list box   *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mTeamPop()
'
'   mTeamPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcTeam.ListIndex
    If ilIndex > 1 Then
        slName = lbcTeam.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "T"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(SPerson, lbcTeam, lbcTeamCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(SPerson, lbcTeam, tmTeamCode(), smTeamCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mTeamPopErr
        gCPErrorMsg ilRet, "mTeamPop (gIMoveListBox)", SPerson
        On Error GoTo 0
        lbcTeam.AddItem "[None]", 0
        lbcTeam.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcTeam
            If gLastFound(lbcTeam) > 1 Then
                lbcTeam.ListIndex = gLastFound(lbcTeam)
            Else
                lbcTeam.ListIndex = -1
            End If
        Else
            lbcTeam.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mTeamPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    sgDoneMsg = Trim$(str$(igSlfCallSource)) & "\" & sgSlfName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload SPerson
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim slStr As String

    If (ilCtrlNo = FIRSTNAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcFirstName, "", "First name must be specified", tmCtrls(FIRSTNAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = FIRSTNAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = LASTNAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcLastName, "", "Last name must be specified", tmCtrls(LASTNAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = LASTNAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SOFFICEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcSOffice, "", "Sales office must be specified", tmCtrls(SOFFICEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SOFFICEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SCODEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcStationCode, "", "Station Salesperson Code must be specified", tmCtrls(SCODEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SCODEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PHONEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(mkcPhone, smPhoneImage, "Phone # must be specified", tmCtrls(PHONEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PHONEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = FAXINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(mkcFax, smFaxImage, "Fax # must be specified", tmCtrls(FAXINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = FAXINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TITLEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcTitle, "", "Job Title must be specified", tmCtrls(TITLEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TITLEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TEAMINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcTeam, "", "Sales team must be specified", tmCtrls(TEAMINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TEAMINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = STATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imState = 0 Then
            slStr = "Active"
        ElseIf imState = 1 Then
            slStr = "Dormant"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Active/Dormant must be specified", tmCtrls(STATEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STATEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = GOALINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcGoal, "", "Sales Goal must be specified", tmCtrls(GOALINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = GOALINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = UNDERINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcUnderComm, "", "Under Commission must be specified", tmCtrls(UNDERINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = UNDERINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = OVERINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcOverComm, "", "Over Commission must be specified", tmCtrls(OVERINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = OVERINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = REMUNDERINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcRemUnderComm, "", "Remnant Under Commission must be specified", tmCtrls(REMUNDERINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = REMUNDERINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = REMOVERINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcRemOverComm, "", "Remnant Over Commission must be specified", tmCtrls(REMOVERINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = REMOVERINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = NEWCLIENTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcNewClientComm, "", "New Client Commission must be specified", tmCtrls(NEWCLIENTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NEWCLIENTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = INCCLIENTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcIncClientComm, "", "Increased Client Commission must be specified", tmCtrls(INCCLIENTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = INCCLIENTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = COMMPAIDINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcCommPaid, "", "Startup Commission Paid must be specified", tmCtrls(COMMPAIDINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMMPAIDINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = COMMSALESINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcSales, "", "Startup Sales must be specified", tmCtrls(COMMSALESINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMMSALESINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = COMMDATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcCommDate, "", "Startup Commission Date must be specified", tmCtrls(COMMDATEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMMDATEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    mTestFields = YES
End Function
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
                edcCommDate.Text = Format$(llDate, "m/d/yy")
                edcCommDate.SelStart = 0
                edcCommDate.SelLength = Len(edcCommDate.Text)
                imBypassFocus = True
                edcCommDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcCommDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcSlfID_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    If (imBoxNo = FIRSTNAMEINDEX) Or (imBoxNo = LASTNAMEINDEX) Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                If (ilBox = SCODEINDEX) And (tgSpf.sAStnCodes = "N") Then
                    Beep
                    mSetFocus imBoxNo
                    Exit Sub
                End If
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcSlfID_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        pbcSlfID.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcSlfID.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcSlfID.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If (imBoxNo = FIRSTNAMEINDEX) Or (imBoxNo = LASTNAMEINDEX) Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If imBoxNo = SOFFICEINDEX Then
        If mSOfficeBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = TEAMINDEX Then
        If mTeamBranch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
        If (imBoxNo <> FIRSTNAMEINDEX) Or (Not cbcSelect.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    imTabDirection = -1  'Set-right to left
    Select Case imBoxNo
        Case -1
            imTabDirection = 0  'Set-Left to right
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                ilBox = 1
                mSetCommands
            Else
                mSetChg 1
                ilBox = 3
            End If
        Case 1 'Name (first control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            ilBox = 1
        Case PHONEINDEX
            If tgSpf.sAStnCodes = "N" Then
                ilBox = SOFFICEINDEX
            Else
                ilBox = SCODEINDEX
            End If
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcState_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcState_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If imState <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imState = 0
        pbcState_Paint
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If imState <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imState = 1
        pbcState_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imState = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imState = 1
            pbcState_Paint
        ElseIf imState = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imState = 0
            pbcState_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imState = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imState = 1
    ElseIf imState = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imState = 0
    End If
    pbcState_Paint
    mSetCommands
End Sub
Private Sub pbcState_Paint()
    pbcState.Cls
    pbcState.CurrentX = fgBoxInsetX
    pbcState.CurrentY = 0 'fgBoxInsetY
    If imState = 0 Then
        pbcState.Print "Active"
    ElseIf imState = 1 Then
        pbcState.Print "Dormant"
    Else
        pbcState.Print "   "
    End If
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If (imBoxNo = FIRSTNAMEINDEX) Or (imBoxNo = LASTNAMEINDEX) Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If imBoxNo = SOFFICEINDEX Then
        If mSOfficeBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = TEAMINDEX Then
        If mTeamBranch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    Select Case imBoxNo
        Case -1
            imTabDirection = -1  'Set-Right to left
            ilBox = UBound(tmCtrls)
        Case SOFFICEINDEX
            If tgSpf.sAStnCodes = "N" Then
                ilBox = PHONEINDEX
            Else
                ilBox = SCODEINDEX
            End If
        Case UBound(tmCtrls) 'last control
            mSetShow imBoxNo
            imBoxNo = -1
            If (cmcUpdate.Enabled) And (igSlfCallSource = CALLNONE) Then
                cmcUpdate.SetFocus
            Else
                cmcDone.SetFocus
            End If
            Exit Sub
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSlfID_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case SOFFICEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcSOffice, edcDropDown, imChgMode, imLbcArrowSetting
        Case TEAMINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcTeam, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
