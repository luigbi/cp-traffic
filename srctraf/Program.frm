VERSION 5.00
Begin VB.Form Program 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6015
   ClientLeft      =   225
   ClientTop       =   1590
   ClientWidth     =   9360
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6015
   ScaleWidth      =   9360
   Begin VB.PictureBox plcDragTime 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   735
      ScaleHeight     =   210
      ScaleWidth      =   1260
      TabIndex        =   39
      Top             =   465
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.ComboBox cbcVeh 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   5745
      TabIndex        =   6
      Top             =   0
      Width           =   3360
   End
   Begin VB.CommandButton cmcLink 
      Appearance      =   0  'Flat
      Caption         =   "Lin&ks"
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
      HelpContextID   =   5
      Left            =   4635
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   75
      Width           =   945
   End
   Begin VB.CommandButton cmcDupl 
      Appearance      =   0  'Flat
      Caption         =   "D&uplicate"
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
      HelpContextID   =   3
      Left            =   7710
      TabIndex        =   22
      Top             =   5430
      Width           =   1005
   End
   Begin VB.PictureBox plcLibInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   1800
      ScaleHeight     =   600
      ScaleWidth      =   7440
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3705
      Visible         =   0   'False
      Width           =   7470
      Begin VB.Label lacLibInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Library Name xxxxxxxxxxxxxxxxxx  Version xx"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   32
         Top             =   90
         Width           =   7305
      End
      Begin VB.Label lacLibInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Start Time xx:xx:xxam  Length xx:xx:xx"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   33
         Top             =   315
         Width           =   7305
      End
   End
   Begin VB.PictureBox pbcLibrary 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
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
      Height          =   4545
      Left            =   7215
      Picture         =   "Program.frx":0000
      ScaleHeight     =   4545
      ScaleWidth      =   7155
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   7155
      Begin VB.Label lacLibFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   1395
         TabIndex        =   29
         Top             =   975
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4620
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcCount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
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
      Height          =   4545
      Left            =   6555
      Picture         =   "Program.frx":11C92
      ScaleHeight     =   4545
      ScaleWidth      =   7155
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   7155
   End
   Begin VB.PictureBox pbcLayout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
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
      Height          =   4545
      Left            =   6105
      Picture         =   "Program.frx":23924
      ScaleHeight     =   4545
      ScaleWidth      =   7155
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   7155
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   4620
      ScaleHeight     =   165
      ScaleWidth      =   60
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   60
   End
   Begin VB.PictureBox plcResol 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3825
      ScaleHeight     =   285
      ScaleWidth      =   1755
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5715
      Width           =   1755
      Begin VB.PictureBox pbcResolType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   840
         ScaleHeight     =   180
         ScaleWidth      =   840
         TabIndex        =   17
         Top             =   30
         Width           =   870
      End
   End
   Begin VB.PictureBox pbcCurrent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4785
      ScaleHeight     =   180
      ScaleWidth      =   720
      TabIndex        =   14
      Top             =   5460
      Width           =   750
   End
   Begin VB.ListBox lbcEvents 
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
      Height          =   450
      ItemData        =   "Program.frx":471D6
      Left            =   75
      List            =   "Program.frx":471D8
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5430
      Width           =   3570
   End
   Begin VB.VScrollBar vbcLayout 
      Height          =   4530
      LargeChange     =   23
      Left            =   8955
      Max             =   24
      Min             =   1
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   555
      Value           =   1
      Width           =   240
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
      ScaleWidth      =   1245
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   1245
   End
   Begin VB.CheckBox ckcShowVersion 
      Caption         =   "Show All Versions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   75
      TabIndex        =   18
      Top             =   240
      Width           =   1710
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
      HelpContextID   =   1
      Left            =   5685
      TabIndex        =   20
      Top             =   5430
      Width           =   885
   End
   Begin VB.CommandButton cmcPrgName 
      Appearance      =   0  'Flat
      Caption         =   "Prg Name"
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
      HelpContextID   =   5
      Left            =   6645
      TabIndex        =   21
      Top             =   5430
      Width           =   1005
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
      HelpContextID   =   6
      Left            =   7710
      TabIndex        =   25
      Top             =   5745
      Width           =   1005
   End
   Begin VB.CommandButton cmcDated 
      Appearance      =   0  'Flat
      Caption         =   "D&ated"
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
      HelpContextID   =   6
      Left            =   5685
      TabIndex        =   23
      Top             =   5745
      Width           =   885
   End
   Begin VB.CommandButton cmcSchedule 
      Appearance      =   0  'Flat
      Caption         =   "&Schedule"
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
      HelpContextID   =   7
      Left            =   6645
      TabIndex        =   24
      Top             =   5745
      Width           =   1005
   End
   Begin VB.PictureBox plcLib 
      ForeColor       =   &H00000000&
      Height          =   4905
      Left            =   75
      ScaleHeight     =   4845
      ScaleWidth      =   1575
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   495
      Width           =   1635
      Begin VB.CommandButton cmcDefSchd 
         Appearance      =   0  'Flat
         Caption         =   "Define S&chedule"
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
         HelpContextID   =   3
         Left            =   45
         TabIndex        =   41
         Top             =   4485
         Width           =   1485
      End
      Begin VB.CommandButton cmcLib 
         Appearance      =   0  'Flat
         Caption         =   "Define &Library"
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
         HelpContextID   =   3
         Left            =   45
         TabIndex        =   8
         Top             =   4140
         Width           =   1485
      End
      Begin VB.ListBox lbcLib 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4020
         ItemData        =   "Program.frx":471DA
         Left            =   45
         List            =   "Program.frx":471DC
         TabIndex        =   7
         Top             =   60
         Width           =   1500
      End
   End
   Begin VB.PictureBox plcType 
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   1860
      ScaleHeight     =   360
      ScaleWidth      =   2550
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   2610
      Begin VB.OptionButton rbcType 
         Caption         =   "Counts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   1680
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Width           =   810
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Layout"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   885
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   90
         Width           =   810
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Library"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   90
         Width           =   840
      End
   End
   Begin VB.PictureBox plcLayout 
      ForeColor       =   &H00000000&
      Height          =   4905
      Left            =   1755
      ScaleHeight     =   4845
      ScaleWidth      =   7440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   495
      Width           =   7500
      Begin VB.HScrollBar hbcDate 
         Height          =   255
         Left            =   990
         Min             =   1
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4605
         Value           =   1
         Width           =   4965
      End
      Begin VB.Label plcWkCnt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5985
         TabIndex        =   38
         Top             =   4590
         Width           =   1455
      End
      Begin VB.Label plcDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   60
         TabIndex        =   37
         Top             =   4590
         Width           =   930
      End
   End
   Begin VB.Timer tmcClick 
      Interval        =   2000
      Left            =   6300
      Top             =   60
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6225
      Top             =   165
   End
   Begin VB.TextBox edcLinkSrceHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4710
      TabIndex        =   34
      Top             =   75
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4290
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lacType 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6120
      TabIndex        =   40
      Top             =   315
      Width           =   2970
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8850
      Top             =   5475
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacLibFrame 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1125
      TabIndex        =   30
      Top             =   1335
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8820
      Picture         =   "Program.frx":471DE
      Top             =   5430
      Width           =   480
   End
End
Attribute VB_Name = "Program"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Program.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmLcfSrchKey1                 tmGhfSrchKey0                 tmGsfSrchKey0             *
'*  tmGsfSrchKey1                                                                         *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Program.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Programming input screen code
'
Option Explicit
Option Compare Text
'Library title
Dim hmLtf As Integer            'Log library file handle
Dim tmLtf As LTF                'LTF record image
Dim tmLtfSrchKey As INTKEY0     'LTF key record image
Dim imLtfRecLen As Integer         'LTF record length
'Library version
Dim hmLvf As Integer            'Log library file handle
Dim tmLvf As LVF                'LVF record image
Dim tmLvfSrchKey As LONGKEY0     'LVF key record image
Dim imLvfRecLen As Integer         'LVF record length
'Library event
Dim hmLef As Integer            'Log event file handle
Dim tmLef As LEF              'Lef record images
Dim tmLefSrchKey As LEFKEY0     'Lef key record image
Dim imLefRecLen As Integer         'Lef record length
'Library calendar
Dim hmLcf As Integer            'Log calendar library file handle
Dim tmCLcf As LCF               'LCF record image-current
Dim tmPLcf As LCF               'LCF record image-pending
Dim tmDLcf As LCF               'LCF record image-delete
Dim tmLcfSrchKey As LCFKEY0     'LCF key record image
Dim imLcfRecLen As Integer         'LCF record length
'Event names
Dim hmEnf As Integer            'Event name file handle
Dim tmEnf As ENF                'Enf record images
Dim tmEnfSrchKey As INTKEY0     'Enf key record image
Dim imEnfRecLen As Integer         'Enf record length

Dim hmGhf As Integer
Dim tmGhf As GHF        'GHF record image
Dim tmGhfSrchKey1 As GHFKEY1    'GHF key record image
Dim imGhfRecLen As Integer        'GHF record length

Dim hmGsf As Integer
Dim tmGsf As GSF        'GSF record image
Dim imGsfRecLen As Integer        'GSF record length
'Vehicle
Dim hmVef As Integer    'Vehicle file handle
Dim tmVef As VEF        'VEF record image
Dim tmVefSrchKey As INTKEY0    'VEF key record image
Dim imVefRecLen As Integer        'CEF record length
Dim tmETypeCode() As SORTCODE
Dim smETypeCodeTag As String
Dim tmLibName() As SORTCODE
Dim smLibNameTag As String
Dim imUpdateAllowed As Integer
Dim imSvWinStatus As Integer  'igWinStatus save value: 0=Hide; 1= View; 2=Input
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer 'True=First focus, then set focus to cbcVeh
Dim imFirstTime As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imComboBoxIndex As Integer
Dim imVefCode As Integer
Dim imVpfIndex As Integer   'Vehicle option index
Dim imCurrent As Integer    '0=Current; 1=Pending
Dim imResol As Integer  '0=Hour; 1= 1/2 hour; 2= 15 Mins
Dim imLibLayCnt As Integer
Dim imDoubleClick As Integer
Dim lmCEarliestDate As Long 'Earliest current date
Dim lmCLatestDate As Long   'Latest current date
Dim lmPEarliestDate As Long 'Earliest pending date
Dim lmPLatestDate As Long   'Latest pending date
Dim fmPaintHeight As Single
Dim imIgnoreChg As Integer
Dim implcLayoutTop As Integer
Dim imNoEvt As Integer
Dim imUnits As Integer
Dim cmSec As Currency
Dim imDatePaint As Integer
Dim imSelectDelay As Integer    'True=cbcSelect change mode
Dim imStartMode As Integer
Dim lmPrgLength As Long         'imDragSource = 0; this contains the library length of time
Dim imDragSource As Integer     '0= Library list box; 1= Library Picture box
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer  'Shift state when mouse down event occurrs
Dim imLCDDragIndex As Integer
Dim imCurrentPending As Integer
Dim imIgnoreRightMove As Integer
Dim imButtonIndex As Integer
Dim imUsingTFNForDate As Integer    'True=Moved Lib to day that does not exist- used TFN for Current
Dim imShowHelpMsg As Integer    'True=Show help messages; False= Ignore help message system
Dim smDragTime As String
Dim tmGsfInfo() As GSFINFO
Dim tmCLLC() As LLC           'Current
Dim tmCDLLC() As LLC           'Current with deletes removed
Dim tmPLLC() As LLC           'Pending
Dim tmPDLLC() As LLC           'Pending with deletes removed
Dim tmCTFN() As LLC             'Current TFN
Dim tmCDTFN() As LLC            'Current TFN with deletes removed
Dim tmPTFN() As LLC             'Pending TFN
Dim tmPDTFN() As LLC            'Pendinf TFN with deletes removed
Dim tmLCD() As LCD
Dim tmPrgVehicle() As SORTCODE
Dim smPrgVehicleTag As String

Dim tmTeam() As MNF
Dim smTeamTag As String

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Height adjustment factor
Dim imStartXFirstColumn As Integer
Dim imWidthToNextColumn As Integer
Dim imWidthWithinColumn As Integer
Dim imStartYFirstColumn As Integer
Dim imHeightWithinColumn As Integer

Const LBONE = 1


Private Sub cbcVeh_Change()
    If imStartMode Then
        imStartMode = False
        mCbcVehChange
        Exit Sub
    End If
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        If cbcVeh.Text <> "" Then
            gManLookAhead cbcVeh, imBSMode, imComboBoxIndex
            lacType.Caption = ""
            tmcClick.Enabled = False
            imSelectDelay = True
            tmcClick.Interval = 2000    '2 seconds
            tmcClick.Enabled = True
        End If
    End If
    Exit Sub
End Sub
Private Sub cbcVeh_Click()
    imComboBoxIndex = cbcVeh.ListIndex
'    igVehIndexViaPrg = imComboBoxIndex
'    mLibPop
    cbcVeh_Change
'    pbcCount.Cls
'    pbcLayout.Cls
End Sub
Private Sub cbcVeh_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub cbcVeh_DropDown()
    tmcClick.Enabled = False
    imSelectDelay = False
End Sub
Private Sub cbcVeh_GotFocus()
    If imFirstTime Then
        imFirstTime = False
    End If
    imFirstFocus = False
    If cbcVeh.Text = "" Then
        gFindMatch sgUserDefVehicleName, 1, cbcVeh
        If gLastFound(cbcVeh) >= 1 Then
            cbcVeh.ListIndex = gLastFound(cbcVeh)
        Else
            If cbcVeh.ListCount > 1 Then
                cbcVeh.ListIndex = 1
            End If
        End If
        imComboBoxIndex = cbcVeh.ListIndex
        igVehIndexViaPrg = imComboBoxIndex - 1
    End If
    imComboBoxIndex = igVehIndexViaPrg + 1
    gCtrlGotFocus cbcVeh
End Sub
Private Sub cbcVeh_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcVeh_KeyPress(KeyAscii As Integer)
    tmcClick.Enabled = False
    imSelectDelay = False
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcVeh.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcVeh_LostFocus()
    If imSelectDelay Then
        tmcClick.Enabled = False
        imSelectDelay = False
        mCbcVehChange
    End If
End Sub
Private Sub ckcShowVersion_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcShowVersion.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    mLibPop
'    mDateSpan
    pbcCount.Cls
    pbcLayout.Cls
    pbcLibrary.Cls
    If rbcType(2).Value Then
        pbcCount_Paint
    ElseIf rbcType(1).Value Then
        pbcLayout_Paint
    Else
        pbcLibrary_Paint
    End If
End Sub
Private Sub ckcShowVersion_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub ckcShowVersion_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDated_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub cmcDated_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcDefSchd_Click()
    Dim ilLoop As Integer
    Dim ilType As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim slStr As String

    If tmVef.sType = "G" Then
        ReDim tgPrg(0 To 1) As PRGDATE  'Time/Dates
        slNameCode = tmPrgVehicle(igVehIndexViaPrg).sKey  'Traffic!lbcVehicle.List(igVehIndexViaPrg)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        slStr = slStr & slName & "\" & slCode
        sgCommandStr = slStr
        GameSchd.Show vbModal
        tmGhfSrchKey1.iVefCode = tmVef.iCode
        ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet <> BTRV_ERR_NONE Then
            tmGhf.lCode = -1
        End If
        smTeamTag = ""
        mTeamPop
        mDateSpan ""
    Else
        ReDim tgPrg(0 To 1) As PRGDATE  'Time/Dates
        tgPrg(0).sStartTime = ""
        tgPrg(0).sStartDate = ""
        For ilLoop = 0 To 6 Step 1
            tgPrg(0).iDay(ilLoop) = 0
        Next ilLoop
        tgPrg(0).sEndDate = ""  '"TFN"
        gUnpackTimeLong tmLvf.iLen(0), tmLvf.iLen(1), False, lmPrgLength
        lgLibLength = lmPrgLength
        PrgDates.Show vbModal
        If UBound(tgPrg) > 0 Then
            Screen.MousePointer = vbHourglass
            ilType = 0
            ilRet = btrBeginTrans(hmLvf, 1000)
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Insert Not Completed, Try Later", vbOKOnly + vbExclamation, "Program")
                Exit Sub
            End If
            ilRet = gPrgToPend(Program, tmLvf, ilType)
            If Not ilRet Then
                ilRet = btrAbortTrans(hmLvf)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Insert Not Completed, Try Later", vbOKOnly + vbExclamation, "Program")
                Exit Sub
            End If
            ilRet = btrEndTrans(hmLvf)
            imCurrent = 1   'Force to pending
            pbcCurrent_Paint
            Screen.MousePointer = vbHourglass
            imDatePaint = False
            mDateSpan tgPrg(0).sStartDate
            Screen.MousePointer = vbHourglass  'Wait
            pbcCount.Cls
            pbcLayout.Cls
            pbcLibrary.Cls
            If rbcType(2).Value Then
                pbcCount_Paint
            ElseIf rbcType(1).Value Then
                If vbcLayout.Value <> vbcLayout.Min Then
                    vbcLayout.Value = vbcLayout.Min
                Else
                    pbcLayout_Paint
                End If
            Else
                If vbcLayout.Value <> vbcLayout.Min Then
                    vbcLayout.Value = vbcLayout.Min
                Else
                    pbcLibrary_Paint
                End If
            End If
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub cmcDone_Click()
    'If Not imUpdateAllowed Then
    '    cmcCancel_Click
    '    Exit Sub
    'End If
    mTerminate
End Sub
Private Sub cmcDone_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDupl_Click()
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    'If Not gWinRoom(igNoExeWinRes(PEVENTEXE)) Then
    '    Exit Sub
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    igPrgDupl = True
    PrgDupl.Show vbModal
    If Not igPrgDupl Then
        Exit Sub
    End If
    'PEvent.Show vbModal
    If Not mStartPEvent() Then
        Exit Sub
    End If
    'Screen.MousePointer = vbDefault    'Default
    'cmcDupl.Enabled = False
    If tmVef.sType <> "G" Then
        cmcDefSchd.Enabled = False
    End If
    cmcLib.Enabled = False
    'imCurrent = 1   'Force to pending
    Screen.MousePointer = vbHourglass
    pbcCurrent_Paint
    mLibPop
    Screen.MousePointer = vbHourglass  'Wait
    imDatePaint = False
    mDateSpan ""
    Screen.MousePointer = vbHourglass  'Wait
'    mCreateLLC "TFN"
    pbcCount.Cls
    pbcLayout.Cls
    pbcLibrary.Cls
    If rbcType(2).Value Then
        pbcCount_Paint
    ElseIf rbcType(1).Value Then
        If vbcLayout.Value <> vbcLayout.Min Then
            vbcLayout.Value = vbcLayout.Min
        Else
            pbcLayout_Paint
        End If
    Else
        If vbcLayout.Value <> vbcLayout.Min Then
            vbcLayout.Value = vbcLayout.Min
        Else
            pbcLibrary_Paint
        End If
    End If
    Screen.MousePointer = vbDefault    'Default
End Sub
Private Sub cmcDupl_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub cmcDupl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcLib_Click()
    'Screen.MousePointer = vbHourGlass  'Wait
    igPrgDupl = False
    'PEvent.Show vbModal
    If Not mStartPEvent() Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault    'Default
    'cmcDupl.Enabled = False
    If tmVef.sType <> "G" Then
        cmcDefSchd.Enabled = False
    End If
    cmcLib.Enabled = False
    'imCurrent = 1   'Force to pending
    Screen.MousePointer = vbHourglass
    pbcCurrent_Paint
    mLibPop
    Screen.MousePointer = vbHourglass
    imDatePaint = False
    mDateSpan ""
    Screen.MousePointer = vbHourglass  'Wait
    pbcCount.Cls
    pbcLayout.Cls
    pbcLibrary.Cls
    If rbcType(2).Value Then
        pbcCount_Paint
    ElseIf rbcType(1).Value Then
        If vbcLayout.Value <> vbcLayout.Min Then
            vbcLayout.Value = vbcLayout.Min
        Else
            pbcLayout_Paint
        End If
    Else
        If vbcLayout.Value <> vbcLayout.Min Then
            vbcLayout.Value = vbcLayout.Min
        Else
            pbcLibrary_Paint
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcLib_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub cmcLink_Click()
    Dim slStr As String

    Dim ilShell As Integer
    Dim dlShellRet As Double
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    Dim ilLoop As Integer

'    'Unload IconTraf
'    'If Not gWinRoom(igNoExeWinRes(LINKSEXE)) Then
'    '    Exit Sub
'    'End If
'    ''Screen.MousePointer = vbHourGlass  'Wait
'    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'    'edcLinkSrceDoneMsg.Text = ""
'    If (Not igStdAloneMode) And (imShowHelpMsg) Then
'        If igTestSystem Then
'            slStr = "Program^Test\" & sgUserName
'        Else
'            slStr = "Program^Prod\" & sgUserName
'        End If
'    Else
'        If igTestSystem Then
'            slStr = "Program^Test^NOHELP\" & sgUserName
'        Else
'            slStr = "Program^Prod^NOHELP\" & sgUserName
'        End If
'    End If
'    'lgShellRet = Shell(sgExePath & "Links.Exe " & slStr, 1)
'    'Program.Enabled = False
'    ''Screen.MousePointer = vbDefault  'Wait
'    'Do While Not igChildDone
'    '    DoEvents
'    'Loop
'    sgCommandStr = slStr
'    Links.Show vbModal
'    slStr = sgDoneMsg
'    'Program.Enabled = True
'    'edcLinkSrceDoneMsg.Text = "Ok"
'    'For ilLoop = 0 To 10
'    '    DoEvents
'    'Next ilLoop
'    ''Screen.MousePointer = vbDefault    'Default
    If igTestSystem Then
        slCommandStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    Else
        slCommandStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    End If
    If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
        ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
        If ilPos > 0 Then
            slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
        End If
    End If
    'Dan M 9/20/10 problems in v57 reports.exe running GetCsiName
    'slDate = Trim$(gGetCSIName("SYSDate"))
    slDate = gCSIGetName()
    If slDate <> "" Then
        'use slDate when writing to file later
        slDate = " /D:" & slDate
        'slCommandStr = slCommandStr & " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
    slCommandStr = slCommandStr & "/IniLoc:" & CurDir$ & " /UserInput"
    Screen.MousePointer = vbDefault
    Traffic.WindowState = vbMinimized
    gShellAndWait Traffic, sgExePath & "Links.exe " & slCommandStr, vbNormalFocus, False    'vbFalse
    Traffic.WindowState = vbMaximized
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcLink_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub cmcLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = PROGRAMMINGJOB
    igRptType = 3               'set flag to send to rptsel
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Program^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        Else
            slStr = "Program^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Program^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    Else
    '        slStr = "Program^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'Program.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'Program.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    sgCommandStr = slStr
    RptList.Show vbModal
    'Screen.MousePointer = vbDefault    'Default
End Sub
Private Sub cmcReport_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub cmcReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcSchedule_Click()
    Dim slStr As String

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(PRGSCHEXE)) Then
    '    Exit Sub
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    edcLinkSrceDoneMsg.Text = ""
    If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Program^Test\" & sgUserName
        Else
            slStr = "Program^Prod\" & sgUserName
        End If
    Else
        If igTestSystem Then
            slStr = "Program^Test^NOHELP\" & sgUserName
        Else
            slStr = "Program^Prod^NOHELP\" & sgUserName
        End If
    End If
    'lgShellRet = Shell(sgExePath & "PrgSch.Exe " & slStr, 1)
    'Program.Enabled = False
    ''Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    PrgSch.Show vbModal
    slStr = sgDoneMsg
    'Program.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    imCurrent = 0   'Force to current
    pbcCurrent_Paint
    Screen.MousePointer = vbHourglass  'Wait
'    mCreateLLC "TFN"
    imDatePaint = False
    mDateSpan ""  'This does not force a paint (imIgnoreChg = True)
    Screen.MousePointer = vbHourglass  'Wait
    pbcCount.Cls
    pbcLayout.Cls
    pbcLibrary.Cls
    pbcResolType_Paint
    imNoEvt = 0
    imUnits = 0
    cmSec = 0
    If rbcType(2).Value Then
'            pbcCount_Paint
        If vbcLayout.Value <> vbcLayout.Min Then
            vbcLayout.Value = vbcLayout.Min
        Else
            pbcCount_Paint
        End If
    ElseIf rbcType(1).Value Then
        If vbcLayout.Value <> vbcLayout.Min Then
            vbcLayout.Value = vbcLayout.Min
        Else
            pbcLayout_Paint
        End If
    Else
        If vbcLayout.Value <> vbcLayout.Min Then
            vbcLayout.Value = vbcLayout.Min
        Else
            pbcLibrary_Paint
        End If
    End If
    Screen.MousePointer = vbDefault    'Default
End Sub
Private Sub cmcSchedule_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub cmcSchedule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcPrgName_Click()
    igPrgNameVefCode = tmVef.iCode
    PrgAirInfo.Show vbModal
End Sub

Private Sub cmcPrgName_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub cmcPrgName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub edcLinkDestDoneMsg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub Form_Activate()
    'If Not imFirstActivate Then
    '    DoEvents    'Process events so pending keys are not sent to this
    '                'form when keypreview turn on
    '    Program.KeyPreview = True  'To get Alt J and Alt L keys
    '    Exit Sub
    'End If
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True  'To get Alt J and Alt L keys
        Exit Sub
    End If
    imFirstActivate = False
    rbcType(0).Value = True
    If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
        pbcCount.Enabled = False
        pbcLayout.Enabled = False
        pbcLibrary.Enabled = False
    Else
        imUpdateAllowed = True
        pbcCount.Enabled = True
        pbcLayout.Enabled = True
        pbcLibrary.Enabled = True
    End If
    gShowBranner imUpdateAllowed
    DoEvents    'Process events so pending keys are not sent to this
                'form when keypreview turn on
    'Program.KeyPreview = True  'To get Alt J and Alt L keys
    pbcCurrent_Paint    'Force Paint
    pbcResolType_Paint
    Me.KeyPreview = True
    Me.ZOrder 0 'Send to front
    Program.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub Form_Deactivate()
    'Program.KeyPreview = False
    Me.KeyPreview = False
End Sub
Private Sub Form_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
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
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100) / Me.height
        Me.height = (lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100
    End If
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmEnf)
    btrDestroy hmEnf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmLef)
    btrDestroy hmLef
    ilRet = btrClose(hmLvf)
    btrDestroy hmLvf
    ilRet = btrClose(hmLtf)
    btrDestroy hmLtf
    Erase tmTeam
    smTeamTag = ""
    Erase tmGsfInfo
    Erase tmCLLC
    Erase tmPLLC
    Erase tmCTFN
    Erase tmPTFN
    Erase tmCDLLC
    Erase tmPDLLC
    Erase tmCDTFN
    Erase tmPDTFN
    Erase tmLCD
    Erase tmETypeCode
    Erase tmLibName
    Erase tmPrgVehicle
    
    Set Program = Nothing   'Remove data segment
    
    igJobShowing(PROGRAMMINGJOB) = False
End Sub
Private Sub hbcDate_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         slNameCode                    slCode                    *
'*  ilTeam                        ilRet                                                   *
'******************************************************************************************

    Dim slDate As String
    Dim llDate As Long
    Dim ilIndex As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilLoop As Integer
    Dim ilLLC As Integer
    Dim ilLowerLLC As Integer

    If imCurrentPending Then
        Exit Sub
    End If
    If tmVef.sType <> "G" Then
        If imCurrent = 0 Then   'Current
            If hbcDate.Value = hbcDate.Max Then
                plcDate.Caption = "TFN"
                slDate = "TFN"
            Else
                llDate = 7 * (hbcDate.Value - 1) + lmCEarliestDate
                slDate = Format(llDate, "m/d/yy")
                slDate = gFormatDate(slDate)
                plcDate.Caption = slDate
            End If
        Else
            If hbcDate.Value = hbcDate.Max Then
                plcDate.Caption = "TFN"
                slDate = "TFN"
            Else
                llDate = 7 * (hbcDate.Value - 1) + lmPEarliestDate
                slDate = Format(llDate, "m/d/yy")
                slDate = gFormatDate(slDate)
                plcDate.Caption = slDate
            End If
        End If
        Screen.MousePointer = vbHourglass  'Wait
        If slDate <> "TFN" Then
            mCreateLLC slDate
        Else
            mCreateLLC "TFN"
        End If
    Else
        ilIndex = 7 * (hbcDate.Value - hbcDate.Min)
        llDate = tmGsfInfo(ilIndex).lGameDate
        slDate = Format$(llDate, "m/d/yy")
        slDate = gFormatDate(slDate)
        plcDate.Caption = Trim$(str$(tmGsfInfo(ilIndex).iGameNo)) & "-" & slDate
        ReDim tmCLLC(0 To 0) As LLC
        ReDim tmCDLLC(0 To 0) As LLC
        ReDim tmPLLC(0 To 0) As LLC
        ReDim tmPDLLC(0 To 0) As LLC
        tmCLLC(0).iDay = -1
        tmCDLLC(0).iDay = -1
        tmPLLC(0).iDay = -1
        tmPDLLC(0).iDay = -1
        imUsingTFNForDate = False
        For ilLoop = 1 To 7 Step 1
            llDate = tmGsfInfo(ilIndex).lGameDate
            slDate = Format$(llDate, "m/d/yy")
            slDate = gFormatDate(slDate)
            gPackDateLong llDate, ilLogDate0, ilLogDate1
            ilLowerLLC = UBound(tmCLLC)
            mReadLcfLnfLef tmGsfInfo(ilIndex).iGameNo, ilLoop, "C", ilLogDate0, ilLogDate1, tmCLLC(), tmCDLLC()
            For ilLLC = ilLowerLLC To UBound(tmCLLC) - 1 Step 1
                tmCLLC(ilLLC).sVisitName = tmGsfInfo(ilIndex).sVisitName
                tmCLLC(ilLLC).sHomeName = tmGsfInfo(ilIndex).sHomeName
            Next ilLLC
            ilIndex = ilIndex + 1
            If ilIndex >= UBound(tmGsfInfo) Then
                Exit For
            End If
        Next ilLoop
        'Exit Sub
    End If
    If imIgnoreChg Then
        imIgnoreChg = False
        imDatePaint = True
        Screen.MousePointer = vbDefault    'Default
        Exit Sub
    End If
    If imDatePaint Then
        If rbcType(2).Value Then
            pbcCount.Cls
            pbcCount_Paint
        ElseIf rbcType(1).Value Then
            pbcLayout.Cls
            pbcLayout_Paint
        Else
            pbcLibrary.Cls
            pbcLibrary_Paint
        End If
    End If
    imDatePaint = True
    imIgnoreChg = False
    Screen.MousePointer = vbDefault    'Default
End Sub
Private Sub hbcDate_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub hbcDate_Scroll()
    Dim slDate As String
    Dim llDate As Long
    Dim ilIndex As Integer

    If tmVef.sType <> "G" Then
        If imCurrent = 0 Then   'Current
            If hbcDate.Value = hbcDate.Max Then
                plcDate.Caption = "TFN"
            Else
                llDate = 7 * (hbcDate.Value - 1) + lmCEarliestDate
                slDate = Format(llDate, "m/d/yy")
                slDate = gFormatDate(slDate)
                plcDate.Caption = slDate
            End If
        Else
            If hbcDate.Value = hbcDate.Max Then
                plcDate.Caption = "TFN"
            Else
                llDate = 7 * (hbcDate.Value - 1) + lmPEarliestDate
                slDate = Format(llDate, "m/d/yy")
                slDate = gFormatDate(slDate)
                plcDate.Caption = slDate
            End If
        End If
    Else
        ilIndex = 7 * (hbcDate.Value - hbcDate.Min)
        llDate = tmGsfInfo(ilIndex).lGameDate
        slDate = Format$(llDate, "m/d/yy")
        slDate = gFormatDate(slDate)
        plcDate.Caption = Trim$(str$(tmGsfInfo(ilIndex).iGameNo)) & "-" & slDate
    End If
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub imcTrash_Click()
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilRet As Integer
    ReDim tgRPrg(0 To 1) As PRGDATE  'Time/Dates
    'Ask which days are to be removed
    'If Not gWinRoom(igNoExeWinRes(PRGDELEXE)) Then
    '    lacLibFrame(imDragSource).Visible = False
    '    imDragType = -1
    '    imcTrash.Visible = False
    '    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    '    imcHelp.Visible = True
    '    Exit Sub
    'End If
    sgRemLibName = tmLCD(imLCDDragIndex).sName
    If imCurrent = 0 Then
        If tmLCD(imLCDDragIndex).iCurOrPend = 0 Then
            tgRPrg(0).sStartTime = tmCLLC(tmLCD(imLCDDragIndex).iLLCIndex).sStartTime
        Else
            tgRPrg(0).sStartTime = tmPLLC(tmLCD(imLCDDragIndex).iLLCIndex).sStartTime
        End If
    Else
        If tmLCD(imLCDDragIndex).iCurOrPend = 0 Then
            tgRPrg(0).sStartTime = tmCDLLC(tmLCD(imLCDDragIndex).iLLCIndex).sStartTime
        Else
            tgRPrg(0).sStartTime = tmPDLLC(tmLCD(imLCDDragIndex).iLLCIndex).sStartTime
        End If
    End If
    For ilLoop = 0 To 6 Step 1
        tgRPrg(0).iDay(ilLoop) = 0
    Next ilLoop
    tgRPrg(0).iDay(tmLCD(imLCDDragIndex).iDay - 1) = 1
    slDate = plcDate.Caption
    llDate = gDateValue(slDate) + tmLCD(imLCDDragIndex).iDay - 1
    tgRPrg(0).sStartDate = Format$(llDate, "m/d/yy")
    tgRPrg(0).sEndDate = "" 'TFN
    'Screen.MousePointer = vbHourGlass  'Wait
    PrgDel.Show vbModal
    'Screen.MousePointer = vbDefault    'Default
    If igRemLibReturn = 1 Then  'Library removed
        Screen.MousePointer = vbHourglass
        If imCurrent = 0 Then
            If tmLCD(imLCDDragIndex).iCurOrPend = 0 Then
                tmLvfSrchKey.lCode = tmCLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            Else
                tmLvfSrchKey.lCode = tmPLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            End If
        Else
            If tmLCD(imLCDDragIndex).iCurOrPend = 0 Then
                tmLvfSrchKey.lCode = tmCDLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            Else
                tmLvfSrchKey.lCode = tmPDLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            End If
        End If
        ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrBeginTrans(hmLvf, 1000)
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                lacLibFrame(imDragSource).Visible = False
                imDragType = -1
                imcTrash.Visible = False
                imcTrash.Picture = IconTraf!imcTrashClosed.Picture
                'imcHelp.Visible = True
                ilRet = MsgBox("Delete Not Completed, Try Later", vbOKOnly + vbExclamation, "Program")
                Exit Sub
            End If
            ilRet = gPrgToDelete(Program, tmLvf, 0)
            If Not ilRet Then
                ilRet = btrAbortTrans(hmLvf)
                Screen.MousePointer = vbDefault
                lacLibFrame(imDragSource).Visible = False
                imDragType = -1
                imcTrash.Visible = False
                imcTrash.Picture = IconTraf!imcTrashClosed.Picture
                'imcHelp.Visible = True
                ilRet = MsgBox("Delete Not Completed, Try Later", vbOKOnly + vbExclamation, "Program")
                Exit Sub
            End If
            ilRet = btrEndTrans(hmLvf)
        End If
        'imCurrent = 1   'Force to pending
        pbcCurrent_Paint
        mLibPop
        Screen.MousePointer = vbHourglass
        imDatePaint = False
        mDateSpan tgRPrg(0).sStartDate  '""
        Screen.MousePointer = vbHourglass  'Wait
        pbcCount.Cls
        pbcLayout.Cls
        pbcLibrary.Cls
        If rbcType(2).Value Then
            pbcCount_Paint
        ElseIf rbcType(1).Value Then
            If vbcLayout.Value <> vbcLayout.Min Then
                vbcLayout.Value = vbcLayout.Min
            Else
                pbcLayout_Paint
            End If
        Else
            If vbcLayout.Value <> vbcLayout.Min Then
                vbcLayout.Value = vbcLayout.Min
            Else
                pbcLibrary_Paint
            End If
        End If
        Screen.MousePointer = vbDefault
    End If
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    'imcHelp.Visible = True
End Sub
Private Sub imcTrash_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacLibFrame(imDragSource).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
'        lacEvtFrame.DragIcon = IconTraf!imcIconTrash.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        lacLibFrame(imDragSource).DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub lbcLib_Click()
    If rbcType(0).Value Then    'Library
        mEventPop
        'If (lbcLib.ListIndex > 0) And (imUpdateAllowed) Then
        If (imUpdateAllowed) Then
            cmcDupl.Enabled = True
        Else
            cmcDupl.Enabled = False
        End If
        If lbcLib.ListIndex >= 0 Then
            If tmVef.sType <> "G" Then
                cmcDefSchd.Enabled = True
            End If
            cmcLib.Enabled = True
        Else
            If tmVef.sType <> "G" Then
                cmcDefSchd.Enabled = False
            End If
            cmcLib.Enabled = False
        End If
    Else
        If rbcType(2).Value Then
            pbcCount.Cls
            pbcCount_Paint
        ElseIf rbcType(1).Value Then
            pbcLayout.Cls
            pbcLayout_Paint
        Else
            pbcLibrary.Cls
            pbcLibrary_Paint
        End If
    End If
End Sub
Private Sub lbcLib_DblClick()
    If rbcType(0).Value Then    'Library
        If igVehIndexViaPrg >= 0 Then
            imDoubleClick = True    'Double click event is followed by a mouse up event
                                    'Process the double click event in the mouse up event
                                    'to avoid the mouse up event being in next form
        End If
    End If
End Sub
Private Sub lbcLib_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub lbcLib_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If plcDate.Caption = "TFN" Then
    '    Exit Sub
    'End If
    If imUpdateAllowed Then
        imDragSource = 0
        fmDragX = X
        fmDragY = Y
        imDragType = 0
        imDragShift = Shift
        tmcDrag.Enabled = True  'Start timer to see if drag or click
    End If
End Sub
Private Sub lbcLib_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If rbcType(0).Value Then    'Library
        If imDoubleClick Then
            imDoubleClick = False
            'If Not gWinRoom(igNoExeWinRes(PEVENTEXE)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourGlass  'Wait
            igPrgDupl = False
            'PEvent.Show vbModal
            If Not mStartPEvent() Then
                Exit Sub
            End If
            'Screen.MousePointer = vbDefault    'Default
            'cmcDupl.Enabled = False
            If tmVef.sType <> "G" Then
                cmcDefSchd.Enabled = False
            End If
            cmcLib.Enabled = False
            'imCurrent = 1   'Force to pending
            Screen.MousePointer = vbHourglass
            pbcCurrent_Paint
            mLibPop
            Screen.MousePointer = vbHourglass
            imDatePaint = False
            mDateSpan ""
            Screen.MousePointer = vbHourglass  'Wait
'            mCreateLLC "TFN"
            pbcCount.Cls
            pbcLayout.Cls
            pbcLibrary.Cls
            If rbcType(2).Value Then
                pbcCount_Paint
            ElseIf rbcType(1).Value Then
                If vbcLayout.Value <> vbcLayout.Min Then
                    vbcLayout.Value = vbcLayout.Min
                Else
                    pbcLayout_Paint
                End If
            Else
                If vbcLayout.Value <> vbcLayout.Min Then
                    vbcLayout.Value = vbcLayout.Min
                Else
                    pbcLibrary_Paint
                End If
            End If
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCbcVehChange                   *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process vehicle change         *
'*                                                     *
'*******************************************************
Private Sub mCbcVehChange()
    Dim ilLoopCount As Integer

    If imChgMode = False Then
        imChgMode = True
        If mVehBranch() Then
            imChgMode = False
            cbcVeh.SetFocus
            Exit Sub
        End If
        ilLoopCount = 0
        tmcClick.Enabled = False    'For safety
        Do
            If ilLoopCount > 0 Then
                If cbcVeh.ListIndex >= 0 Then
                    cbcVeh.Text = cbcVeh.List(cbcVeh.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            If cbcVeh.Text <> "" Then
                gManLookAhead cbcVeh, imBSMode, imComboBoxIndex
            End If
            igVehIndexViaPrg = cbcVeh.ListIndex - 1
            mLibPop
            If imTerminate Then
                imChgMode = False
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass  'Wait
    '        mCreateLLC "TFN"
            imDatePaint = False
            mDateSpan "" 'This does not force a paint (imIgnoreChg = True)
            Screen.MousePointer = vbHourglass  'Wait
            If tgVpf(imVpfIndex).sGGridRes = "H" Then
                imResol = 1
                vbcLayout.Max = 25
            ElseIf tgVpf(imVpfIndex).sGGridRes = "Q" Then
                imResol = 2
                vbcLayout.Max = 73
            Else
                imResol = 0
                vbcLayout.Max = 1
            End If
            pbcCount.Cls
            pbcLayout.Cls
            pbcLibrary.Cls
            pbcResolType_Paint
            imNoEvt = 0
            imUnits = 0
            cmSec = 0
            If rbcType(2).Value Then
    '            pbcCount_Paint
                If vbcLayout.Value <> vbcLayout.Min Then
                    vbcLayout.Value = vbcLayout.Min
                Else
                    pbcCount_Paint
                End If
            ElseIf rbcType(1).Value Then
                If vbcLayout.Value <> vbcLayout.Min Then
                    vbcLayout.Value = vbcLayout.Min
                Else
                    pbcLayout_Paint
                End If
            Else
                If vbcLayout.Value <> vbcLayout.Min Then
                    vbcLayout.Value = vbcLayout.Min
                Else
                    pbcLibrary_Paint
                End If
            End If
            If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
                cmcPrgName.Enabled = True
            Else
                cmcPrgName.Enabled = False
            End If
        Loop While igVehIndexViaPrg <> cbcVeh.ListIndex - 1
        Screen.MousePointer = vbDefault    'Default
        imChgMode = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateLLC                      *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create Library/Layout/Count    *
'*                      records for current/pending    *
'*                                                     *
'*******************************************************
Private Sub mCreateLLC(slDate As String)
'
'   mCreate slDate
'   Where:
'       slDate (I)- Start date of week or TFN
'
    Dim ilLoop As Integer
    Dim slRunDate As String
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim llDate As Long
    If slDate <> "TFN" Then
        ReDim tmCLLC(0 To 0) As LLC
        ReDim tmCDLLC(0 To 0) As LLC
        ReDim tmPLLC(0 To 0) As LLC
        ReDim tmPDLLC(0 To 0) As LLC
        tmCLLC(0).iDay = -1
        tmCDLLC(0).iDay = -1
        tmPLLC(0).iDay = -1
        tmPDLLC(0).iDay = -1
    Else
        ReDim tmCTFN(0 To 0) As LLC
        ReDim tmCDTFN(0 To 0) As LLC
        ReDim tmPTFN(0 To 0) As LLC
        ReDim tmPDTFN(0 To 0) As LLC
        tmCTFN(0).iDay = -1
        tmCDTFN(0).iDay = -1
        tmPTFN(0).iDay = -1
        tmPDTFN(0).iDay = -1
    End If
    imUsingTFNForDate = False
    For ilLoop = 1 To 7 Step 1
        If slDate <> "TFN" Then
            llDate = gDateValue(slDate) + ilLoop - 1
            If llDate <= lmCLatestDate Then
                slRunDate = Format(llDate, "m/d/yy")
                gPackDate slRunDate, ilLogDate0, ilLogDate1
                mReadLcfLnfLef 0, ilLoop, "C", ilLogDate0, ilLogDate1, tmCLLC(), tmCDLLC()
                imUsingTFNForDate = False
            Else    'Use TFN current for current since not build
                ilLogDate0 = ilLoop  'TFN
                ilLogDate1 = 0
                mReadLcfLnfLef 0, ilLoop, "C", ilLogDate0, ilLogDate1, tmCLLC(), tmCDLLC()
                If UBound(tmCLLC) > LBound(tmCLLC) Then
                    imUsingTFNForDate = True
                End If
            End If
            llDate = gDateValue(slDate) + ilLoop - 1
            slRunDate = Format(llDate, "m/d/yy")
            gPackDate slRunDate, ilLogDate0, ilLogDate1
            mReadLcfLnfLef 0, ilLoop, "P", ilLogDate0, ilLogDate1, tmPLLC(), tmPDLLC()
        Else
            ilLogDate0 = ilLoop  'TFN
            ilLogDate1 = 0
            mReadLcfLnfLef 0, ilLoop, "C", ilLogDate0, ilLogDate1, tmCTFN(), tmCDTFN()
            ilLogDate0 = ilLoop  'TFN
            ilLogDate1 = 0
            mReadLcfLnfLef 0, ilLoop, "P", ilLogDate0, ilLogDate1, tmPTFN(), tmPDTFN()
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDateSpan                       *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine date span of current *
'*                      pending                        *
'*                                                     *
'*******************************************************
Private Sub mDateSpan(slInitDate As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilUpper                       slSortDate                    llSortTime                *
'*  slSortTime                    slSortGameNo                  ilTeam                    *
'*                                                                                        *
'******************************************************************************************

'
'   mDateSpan slInitDate
'   Where:
'       slIniDate(I)- initial date to select or ""
'
    Dim ilType As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilWkIndex As Integer
    Dim llDLatestDate As Long
    Dim llDEarliestDate As Long
    Dim ilInitDate As Integer

    lmCEarliestDate = -1
    lmCLatestDate = -1
    lmPEarliestDate = -1
    lmPLatestDate = -1
    If tmVef.sType <> "G" Then
        tmLcfSrchKey.iType = 0
        ilType = 0
        tmLcfSrchKey.sStatus = "C"
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = 257  'Year 1/1/2100
        tmLcfSrchKey.iLogDate(1) = 2100
        tmLcfSrchKey.iSeqNo = 1
        ilRet = btrGetLessOrEqual(hmLcf, tmCLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) And (tmCLcf.sStatus = "C") And (tmCLcf.iVefCode = imVefCode) And (tmCLcf.iType = ilType) Then
            gUnpackDate tmCLcf.iLogDate(0), tmCLcf.iLogDate(1), slDate
            lmCLatestDate = gDateValue(gObtainNextSunday(slDate))
            tmLcfSrchKey.iType = ilType
            tmLcfSrchKey.sStatus = "C"
            tmLcfSrchKey.iVefCode = imVefCode
            tmLcfSrchKey.iLogDate(0) = 257  'Year 1/1/1900
            tmLcfSrchKey.iLogDate(1) = 1900
            tmLcfSrchKey.iSeqNo = 1
            ilRet = btrGetGreaterOrEqual(hmLcf, tmCLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tmCLcf.sStatus = "C") And (tmCLcf.iVefCode = imVefCode) And (tmCLcf.iType = ilType) Then
                gUnpackDate tmCLcf.iLogDate(0), tmCLcf.iLogDate(1), slDate
                lmCEarliestDate = gDateValue(gObtainPrevMonday(slDate))
            Else
                slDate = Format$(gNow(), "m/d/yy")
                slDate = gObtainNextMonday(slDate)
                slDate = gDecOneWeek(slDate)
                lmCEarliestDate = gDateValue(slDate)
            End If
        Else
            slDate = Format$(gNow(), "m/d/yy")
            slDate = gObtainNextMonday(slDate)
            lmCEarliestDate = gDateValue(slDate)
            lmCLatestDate = lmCEarliestDate + 6
        End If
        tmLcfSrchKey.iType = ilType
        tmLcfSrchKey.sStatus = "P"
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = 257  'Year 1/1/2100
        tmLcfSrchKey.iLogDate(1) = 2100
        tmLcfSrchKey.iSeqNo = 1
        ilRet = btrGetLessOrEqual(hmLcf, tmPLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) And (tmPLcf.sStatus = "P") And (tmPLcf.iVefCode = imVefCode) And (tmPLcf.iType = ilType) Then
            gUnpackDate tmPLcf.iLogDate(0), tmPLcf.iLogDate(1), slDate
            lmPLatestDate = gDateValue(gObtainNextSunday(slDate))
            tmLcfSrchKey.iType = ilType
            tmLcfSrchKey.sStatus = "P"
            tmLcfSrchKey.iVefCode = imVefCode
            tmLcfSrchKey.iLogDate(0) = 257  'Year 1/1/1900
            tmLcfSrchKey.iLogDate(1) = 1900
            tmLcfSrchKey.iSeqNo = 1
            ilRet = btrGetGreaterOrEqual(hmLcf, tmPLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tmPLcf.sStatus = "P") And (tmPLcf.iVefCode = imVefCode) And (tmPLcf.iType = ilType) Then
                gUnpackDate tmPLcf.iLogDate(0), tmPLcf.iLogDate(1), slDate
                lmPEarliestDate = gDateValue(gObtainPrevMonday(slDate))
            Else
                slDate = Format$(gNow(), "m/d/yy")
                slDate = gObtainNextMonday(slDate)
                slDate = gDecOneWeek(slDate)
                lmPEarliestDate = gDateValue(slDate)
            End If
        'Else
        '    slDate = Format$(gNow(), "m/d/yy")
        '    slDate = gObtainNextMonday(slDate)
        '    lmPEarliestDate = gDateValue(slDate)
        '    lmPLatestDate = lmPEarliestDate + 6
        End If
        tmLcfSrchKey.iType = ilType
        tmLcfSrchKey.sStatus = "D"
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = 257  'Year 1/1/2100
        tmLcfSrchKey.iLogDate(1) = 2100
        tmLcfSrchKey.iSeqNo = 1
        ilRet = btrGetLessOrEqual(hmLcf, tmPLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) And (tmPLcf.sStatus = "D") And (tmPLcf.iVefCode = imVefCode) And (tmPLcf.iType = ilType) Then
            gUnpackDate tmPLcf.iLogDate(0), tmPLcf.iLogDate(1), slDate
            llDLatestDate = gDateValue(gObtainNextSunday(slDate))
            If lmPLatestDate <> -1 Then
                If llDLatestDate > lmPLatestDate Then
                    lmPLatestDate = llDLatestDate
                End If
            Else
                lmPLatestDate = llDLatestDate
            End If
            tmLcfSrchKey.iType = ilType
            tmLcfSrchKey.sStatus = "D"
            tmLcfSrchKey.iVefCode = imVefCode
            tmLcfSrchKey.iLogDate(0) = 257  'Year 1/1/1900
            tmLcfSrchKey.iLogDate(1) = 1900
            tmLcfSrchKey.iSeqNo = 1
            ilRet = btrGetGreaterOrEqual(hmLcf, tmPLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tmPLcf.sStatus = "D") And (tmPLcf.iVefCode = imVefCode) And (tmPLcf.iType = ilType) Then
                gUnpackDate tmPLcf.iLogDate(0), tmPLcf.iLogDate(1), slDate
                llDEarliestDate = gDateValue(gObtainPrevMonday(slDate))
                If lmPEarliestDate <> -1 Then
                    If llDEarliestDate < lmPEarliestDate Then
                        lmPEarliestDate = llDEarliestDate
                    End If
                Else
                    lmPEarliestDate = llDEarliestDate
                End If
            Else
                If lmPEarliestDate = -1 Then
                    slDate = Format$(gNow(), "m/d/yy")
                    slDate = gObtainNextMonday(slDate)
                    slDate = gDecOneWeek(slDate)
                    lmPEarliestDate = gDateValue(slDate)
                End If
            End If
        Else
            If lmPEarliestDate = -1 Then
                slDate = Format$(gNow(), "m/d/yy")
                slDate = gObtainNextMonday(slDate)
                lmPEarliestDate = gDateValue(slDate)
                lmPLatestDate = lmPEarliestDate + 6
            End If
        End If
    Else
'        ReDim tmGsfInfo(0 To 0) As GSFINFO
'        tmLcfSrchKey1.iVefCode = tmVef.iCode
'        tmLcfSrchKey1.iType = 0
'        ilRet = btrGetGreaterOrEqual(hmLcf, tmCLcf, imLcfRecLen, tmLcfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
'        Do While (ilRet = BTRV_ERR_NONE) And (tmCLcf.iVefCode = tmVef.iCode)
'            ilUpper = UBound(tmGsfInfo)
'            tmGsfInfo(ilUpper).iGameNo = tmCLcf.iType
'            gUnpackDateLong tmCLcf.iLogDate(0), tmCLcf.iLogDate(1), tmGsfInfo(ilUpper).lGameDate
'            gUnpackDateForSort tmCLcf.iLogDate(0), tmCLcf.iLogDate(1), slSortDate
'            gUnpackTimeLong tmCLcf.iTime(0, LBound(tmCLcf.lLvfCode)), tmCLcf.iTime(1, LBound(tmCLcf.lLvfCode)), False, llSortTime
'            slSortTime = Trim$(Str$(llSortTime))
'            Do While Len(slSortTime) < 6
'                slSortTime = "0" & slSortTime
'            Loop
'            slSortGameNo = Trim$(Str$(tmCLcf.iType))
'            Do While Len(slSortGameNo) < 3
'                slSortGameNo = "0" & slSortGameNo
'            Loop
'            tmGsfInfo(ilUpper).sKey = slSortDate & slSortTime & slSortGameNo
'            'Get Names
'            If tmGhf.lCode > 0 Then
'                tmGsfSrchKey1.lGhfCode = tmGhf.lCode
'                tmGsfSrchKey1.iGameNo = tmGsfInfo(ilUpper).iGameNo
'                ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
'                If ilRet = BTRV_ERR_NONE Then
'                    For ilTeam = 1 To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
'                        If tmGsf.iVisitMnfCode = tmTeam(ilTeam).iCode Then
'                            tmGsfInfo(ilUpper).sVisitName = tmTeam(ilTeam).sName
'                            If Trim$(tmTeam(ilTeam).sUnitType) <> "" Then
'                                tmGsfInfo(ilUpper).sVisitName = tmTeam(ilTeam).sUnitType
'                            End If
'                            Exit For
'                        End If
'                    Next ilTeam
'                    For ilTeam = 1 To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
'                        If tmGsf.iHomeMnfCode = tmTeam(ilTeam).iCode Then
'                            tmGsfInfo(ilUpper).sHomeName = tmTeam(ilTeam).sName
'                            If Trim$(tmTeam(ilTeam).sUnitType) <> "" Then
'                                tmGsfInfo(ilUpper).sHomeName = tmTeam(ilTeam).sUnitType
'                            End If
'                            Exit For
'                        End If
'                    Next ilTeam
'                Else
'                    tmGsfInfo(ilUpper).sVisitName = ""
'                    tmGsfInfo(ilUpper).sHomeName = ""
'                End If
'            Else
'                tmGsfInfo(ilUpper).sVisitName = ""
'                tmGsfInfo(ilUpper).sHomeName = ""
'            End If
'            ReDim Preserve tmGsfInfo(0 To ilUpper + 1) As GSFINFO
'            ilRet = btrGetNext(hmLcf, tmCLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'        If UBound(tmGsfInfo) - 1 > 0 Then
'            ArraySortTyp fnAV(tmGsfInfo(), 0), UBound(tmGsfInfo), 0, LenB(tmGsfInfo(0)), 0, LenB(tmGsfInfo(0).sKey), 0
'        End If
        ilRet = gGetGameDates(hmLcf, hmGhf, hmGsf, tmVef.iCode, tmTeam(), tmGsfInfo())
        hbcDate.Min = 1
        hbcDate.Max = UBound(tmGsfInfo) \ 7 + 1
        If hbcDate.Value <> hbcDate.Min Then
            hbcDate.Value = hbcDate.Min
        Else
            hbcDate_Change
        End If
        Exit Sub
    End If
    hbcDate.Min = 1
    imIgnoreChg = True
    If imCurrent = 0 Then   'Current
        hbcDate.Max = (lmCLatestDate - lmCEarliestDate) \ 7 + 2 '1 for adjusting and 1 for TFN
        imIgnoreChg = True
        ilInitDate = False
        If slInitDate <> "" Then
            If (gDateValue(slInitDate) >= lmCEarliestDate) And (gDateValue(slInitDate) <= lmCLatestDate) Then
                ilInitDate = True
            End If
        End If
        If Not ilInitDate Then
            If hbcDate.Max > 2 Then
                If (lgEarliestDateViaPrg > 0) And (lgEarliestDateViaPrg >= lmCEarliestDate) Then
                    ilWkIndex = (lgEarliestDateViaPrg - lmCEarliestDate) \ 7
                    If hbcDate.Min + ilWkIndex >= hbcDate.Max Then
                        ilWkIndex = hbcDate.Max - 2
                    End If
                Else
                    ilWkIndex = 1
                End If
                If hbcDate.Value <> hbcDate.Min + ilWkIndex Then
                    hbcDate.Value = hbcDate.Min + ilWkIndex
                Else
                    hbcDate_Change
                End If
            Else
                If hbcDate.Value <> hbcDate.Min Then
                    hbcDate.Value = hbcDate.Min
                Else
                    hbcDate_Change
                End If
            End If
        Else
            ilWkIndex = (gDateValue(slInitDate) - lmCEarliestDate) \ 7
            If hbcDate.Value <> hbcDate.Min + ilWkIndex Then
                hbcDate.Value = hbcDate.Min + ilWkIndex
            Else
                hbcDate_Change
            End If
        End If
    Else    'Pending
        hbcDate.Max = (lmPLatestDate - lmPEarliestDate) \ 7 + 2
        imIgnoreChg = True
        ilInitDate = False
        If slInitDate <> "" Then
            If (gDateValue(slInitDate) >= lmPEarliestDate) And (gDateValue(slInitDate) <= lmPLatestDate) Then
                ilInitDate = True
            End If
        End If
        If Not ilInitDate Then
            If hbcDate.Max > 2 Then
                If hbcDate.Value <> hbcDate.Min + 1 Then
                    hbcDate.Value = hbcDate.Min + 1
                Else
                    hbcDate_Change
                End If
            Else
                If hbcDate.Value <> hbcDate.Min Then
                    hbcDate.Value = hbcDate.Min
                Else
                    hbcDate_Change
                End If
            End If
        Else
            ilWkIndex = (gDateValue(slInitDate) - lmPEarliestDate) \ 7
            If hbcDate.Value <> hbcDate.Min + ilWkIndex Then
                hbcDate.Value = hbcDate.Min + ilWkIndex
            Else
                hbcDate_Change
            End If
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEventPop                       *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate event list box        *
'*                                                     *
'*******************************************************
Private Sub mEventPop()
'
'   iRet = mEventPop
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slLength As String
    lbcEvents.Clear
    If lbcLib.ListIndex <= 0 Then
        Exit Sub
    End If
    slNameCode = tmLibName(lbcLib.ListIndex - 1).sKey  'lbcLibName.List(lbcLib.ListIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mEventPopErr
    gCPErrorMsg ilRet, "mEventPop (gParseItem field 2)", Program
    On Error GoTo 0
    tmLvfSrchKey.lCode = CLng(slCode)
    ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
    tmLefSrchKey.lLvfCode = CLng(slCode)
    tmLefSrchKey.iStartTime(0) = 0
    tmLefSrchKey.iStartTime(1) = 0
    tmLefSrchKey.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hmLef, tmLef, imLefRecLen, tmLefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmLef.lLvfCode = CLng(slCode))
        gUnpackLength tmLef.iStartTime(0), tmLef.iStartTime(1), "3", False, slStr
        If (tmLef.iEtfCode < 10) Or (tmLef.iEtfCode > 13) Then
            If tmLef.iEnfCode <> tmEnf.iCode Then
                If tmLef.iEnfCode > 0 Then
                    tmEnfSrchKey.iCode = tmLef.iEnfCode
                    ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Select Case tmLef.iEtfCode
                            Case 1  'Program
                                tmEnf.sName = "Program"
                            Case 2  'Contract Avail
                                tmEnf.sName = "Avail"
                            Case 3
                                tmEnf.sName = "Open BB Avail"
                            Case 4
                                tmEnf.sName = "Floating Avail"
                            Case 5
                                tmEnf.sName = "Close BB Avail"
                            Case 6  'Cmml Promo
                                tmEnf.sName = "Cmml Promo"
                            Case 7  'Feed avail
                                tmEnf.sName = "Feed Avail"
                            Case 8  'PSA/Promo (Avail)
                                tmEnf.sName = "PSA Avail"
                            Case 9
                                tmEnf.sName = "Promo Avail"
                            Case 10  'Page eject, Line space 1, 2 or 3
                                tmEnf.sName = "Page Skip"
                            Case 11
                                tmEnf.sName = "1 Line Space"
                            Case 12
                                tmEnf.sName = "2 Line Spaces"
                            Case 13
                                tmEnf.sName = "3 Line Spaces"
                            Case Else   'Other
                                tmEnf.sName = "Other Events"
                        End Select
                    End If
                Else
                    tmEnf.sName = ""
                End If
            End If
        End If
        Select Case tmLef.iEtfCode
            Case 1  'Program
                gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", False, slLength
                'slStr = slStr & " Program " & slLength
                slStr = slStr & " " & Trim$(tmEnf.sName) & " " & slLength
            Case 2  'Contract Avail
                gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "2", True, slLength
                If Len(slLength) > 0 Then
                    slLength = Left$(slLength, Len(slLength) - 1)  'Remove "
                End If
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    'slStr = slStr & " Avail " & Trim(Str$(tmLef.iMaxUnits)) & "/" & slLength
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & "/" & slLength
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    'slStr = slStr & " Avail " & Trim(Str$(tmLef.iMaxUnits)) & "/" & slLength
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & "/" & slLength
                Else
                    'slStr = slStr & " Avail " & Trim(Str$(tmLef.iMaxUnits)) & " Units"
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & " Units"
                End If
            Case 3
                'slStr = slStr & " Open BB Avail"
                slStr = slStr & " " & Trim$(tmEnf.sName)
            Case 4
                'slStr = slStr & " Floating Avail"
                slStr = slStr & " " & Trim$(tmEnf.sName)
            Case 5
                'slStr = slStr & " Close BB Avail"
                slStr = slStr & " " & Trim$(tmEnf.sName)
            Case 6  'Cmml Promo
                gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "2", True, slLength
                If Len(slLength) > 0 Then
                    slLength = Left$(slLength, Len(slLength) - 1)  'Remove "
                End If
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    'slStr = slStr & " Cmml Promo " & Trim(Str$(tmLef.iMaxUnits)) & "/" & slLength
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & "/" & slLength
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    'slStr = slStr & " Cmml Promo " & Trim(Str$(tmLef.iMaxUnits)) & "/" & slLength
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & "/" & slLength
                Else
                    'slStr = slStr & " Cmml Promo " & Trim(Str$(tmLef.iMaxUnits)) & " Units"
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & " Units"
                End If
            Case 7  'Feed avail
                gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "2", True, slLength
                If Len(slLength) > 0 Then
                    slLength = Left$(slLength, Len(slLength) - 1)  'Remove "
                End If
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    'slStr = slStr & " Feed Avail " & Trim(Str$(tmLef.iMaxUnits)) & "/" & slLength
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & "/" & slLength
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    'slStr = slStr & " Feed Avail " & Trim(Str$(tmLef.iMaxUnits)) & "/" & slLength
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & "/" & slLength
                Else
                    'slStr = slStr & " Feed Avail " & Trim(Str$(tmLef.iMaxUnits)) & " Units"
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & " Units"
                End If
            Case 8  'PSA/Promo (Avail)
                gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "2", True, slLength
                If Len(slLength) > 0 Then
                    slLength = Left$(slLength, Len(slLength) - 1)  'Remove "
                End If
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    'slStr = slStr & " PSA Avail " & Trim(Str$(tmLef.iMaxUnits)) & "/" & slLength
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & "/" & slLength
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    'slStr = slStr & " PSA Avail " & Trim(Str$(tmLef.iMaxUnits)) & "/" & slLength
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & "/" & slLength
                Else
                    'slStr = slStr & " PSA Avail " & Trim(Str$(tmLef.iMaxUnits)) & " Units"
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & " Units"
                End If
            Case 9
                gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "2", True, slLength
                If Len(slLength) > 0 Then
                    slLength = Left$(slLength, Len(slLength) - 1)  'Remove "
                End If
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    'slStr = slStr & " Promo Avail " & Trim(Str$(tmLef.iMaxUnits)) & "/" & slLength
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & "/" & slLength
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    'slStr = slStr & " Promo Avail " & Trim(Str$(tmLef.iMaxUnits)) & "/" & slLength
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & "/" & slLength
                Else
                    'slStr = slStr & " Promo Avail " & Trim(Str$(tmLef.iMaxUnits)) & " Units"
                    slStr = slStr & " " & Trim$(tmEnf.sName) & " " & Trim(str$(tmLef.iMaxUnits)) & " Units"
                End If
            Case 10  'Page eject, Line space 1, 2 or 3
                slStr = slStr & " Page Skip"
            Case 11
                slStr = slStr & " 1 Line Space"
            Case 12
                slStr = slStr & " 2 Line Spaces"
            Case 13
                slStr = slStr & " 3 Line Spaces"
            Case Else   'Other
                gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, slLength
                'slStr = slStr & " Other " & slLength
                slStr = slStr & " " & Trim$(tmEnf.sName) & " " & slLength
        End Select
        lbcEvents.AddItem slStr
        ilRet = btrGetNext(hmLef, tmLef, imLefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_KEY_NOT_FOUND) Then
        On Error GoTo mEventPopErr
        gBtrvErrorMsg ilRet, "mEventPop (btrGetNext):" & "Lef.Btr", Program
        On Error GoTo 0
    End If
    Exit Sub
mEventPopErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
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
    imFirstActivate = True
    igJobShowing(PROGRAMMINGJOB) = True
    imTerminate = False         'terminate if true
    'Program.Height = cmcReport.Top + 5 * cmcReport.Height / 3 - 45
    'gCenterForm Program
    imSvWinStatus = igWinStatus(PROGRAMMINGJOB)
    'fmPaintHeight = 4320
    igLibType = 0
    imFirstFocus = True
    imFirstTime = True
    imSelectDelay = False
    imStartMode = True
    igVehIndexViaPrg = -1
    imVefCode = -1
    imLibLayCnt = -1
    imButtonIndex = -1
    imIgnoreRightMove = False
    imIgnoreChg = False
    imFirstActivate = True
    imUsingTFNForDate = True
    imNoEvt = 0
    imUnits = 0
    cmSec = 0
    imcTrash.Visible = False
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    'imcHelp.Visible = True
    imDatePaint = True
    ReDim tmGsfInfo(0 To 0) As GSFINFO
    ReDim tmCLLC(0 To 0) As LLC
    ReDim tmCDLLC(0 To 0) As LLC
    ReDim tmPLLC(0 To 0) As LLC
    ReDim tmPDLLC(0 To 0) As LLC
    tmCLLC(0).iDay = -1
    tmCDLLC(0).iDay = -1
    tmPLLC(0).iDay = -1
    tmPDLLC(0).iDay = -1
    ReDim tmCTFN(0 To 0) As LLC
    ReDim tmCDTFN(0 To 0) As LLC
    ReDim tmPTFN(0 To 0) As LLC
    ReDim tmPDTFN(0 To 0) As LLC
    tmCTFN(0).iDay = -1
    tmCDTFN(0).iDay = -1
    tmPTFN(0).iDay = -1
    tmPDTFN(0).iDay = -1
    If (tgSpf.sSSellNet = "Y") Or (tgSpf.sSDelNet = "Y") Then
        cmcLink.Visible = True
    Else
        cmcLink.Visible = False
    End If
'    pbcLibType.Enabled = False
    'cmcDupl.Enabled = False
    cmcDefSchd.Enabled = False
    cmcLib.Enabled = False
'    cmcSchedule.Enabled = False
    imCurrent = 0   'Current
    imResol = 0     'Hour
'    pbcResolType.Enabled = False
    imDoubleClick = False
'    plcDate.Caption = "TFN"

    Screen.MousePointer = vbHourglass
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    mInitBox
    gCenterForm Program
    fmPaintHeight = pbcLayout.height - 210 - 30 '4320
    'Program.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    'gCenterForm Program
    'Program.Show vbModeless
    cbcVeh.Clear
    mVehPop
    If imTerminate Then
        Exit Sub
    End If

    'Populate lbcETypeCode-> required by week count (pbcLibrary_Paint)
    'ilRet = gPopEvtNmByTypeBox(Program, True, True, lbcLib, lbcETypeCode)
    ilRet = gPopEvtNmByTypeBox(Program, True, True, lbcLib, tmETypeCode(), smETypeCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mInitErr
        gCPErrorMsg ilRet, "mInit (gPopEvtNmByTypeBox: Event Type)", Program
        On Error GoTo 0
    End If

    smTeamTag = ""
    mTeamPop

    hmLtf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLtf, "", sgDBPath & "Ltf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ltf.Btr)", Program
    On Error GoTo 0
    hmLvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lvf.Btr)", Program
    On Error GoTo 0
    hmLef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLef, "", sgDBPath & "Lef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lef.btr)", Program
    On Error GoTo 0
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", Program
    On Error GoTo 0
    hmEnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmEnf, "", sgDBPath & "Enf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Enf.Btr)", Program
    On Error GoTo 0

    hmGsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gsf.Btr)", Program
    On Error GoTo 0

    hmGhf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", Program
    On Error GoTo 0

    hmVef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", Program
    On Error GoTo 0
    imLtfRecLen = Len(tmLtf)  'Get and save LTF record length
    imLvfRecLen = Len(tmLvf)  'Get and save LTF record length
    imLefRecLen = Len(tmLef)  'Get and save LEF record length
    imLcfRecLen = Len(tmCLcf)  'Get and save LEF record length
    imEnfRecLen = Len(tmEnf)  'Get and save ENF record length
    imGsfRecLen = Len(tmGsf)  'Get and save ENF record length
    imGhfRecLen = Len(tmGhf)  'Get and save ENF record length
    imVefRecLen = Len(tmVef)  'Get and save VEF record length
    'Force vehicle to show since focus is not set to combo first
    gFindMatch sgUserDefVehicleName, 1, cbcVeh
    If gLastFound(cbcVeh) >= 1 Then
        cbcVeh.ListIndex = gLastFound(cbcVeh)
    Else
        If cbcVeh.ListCount > 1 Then
            cbcVeh.ListIndex = 1
        End If
    End If
'    cbcVeh_Change
    cbcVeh.SelLength = 0
    'Traffic!plcHelp.Caption = ""
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
'*             Created:5/10/93       By:D. LeVine      *
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
    flTextHeight = pbcLayout.TextHeight("1") - 35
    'Position panel and picture areas with panel
    implcLayoutTop = 520
    plcLayout.Move 1740, implcLayoutTop, pbcLayout.Width + vbcLayout.Width + fgPanelAdj, pbcLayout.height + fgPanelAdj + hbcDate.height - 30
    pbcLayout.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
    pbcCount.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
    pbcLibrary.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
    plcDate.Move fgBevelX, fgBevelY + pbcLayout.height - 30
    plcWkCnt.Move plcLayout.Width - plcWkCnt.Width - fgBevelX, plcDate.Top
    hbcDate.Move plcDate.Left + plcDate.Width, plcDate.Top, plcWkCnt.Left - plcDate.Width
    lbcLib.Width = fmAdjFactorW * lbcLib.Width
    plcLib.Width = lbcLib.Width + 2 * fgBevelX + 45
    lbcLib.Top = fgBevelY
    lbcLib.Left = fgBevelX
    plcLib.Move 60, plcLayout.Top, plcLib.Width, plcLayout.height

    cbcVeh.Width = fmAdjFactorW * cbcVeh.Width
    imStartXFirstColumn = 420    'Start X of first column
    imWidthToNextColumn = 960
    imWidthToNextColumn = CLng(fmAdjFactorW * imWidthToNextColumn)
    Do While (imWidthToNextColumn Mod 15) <> 0
        imWidthToNextColumn = imWidthToNextColumn + 1
    Loop
    imWidthWithinColumn = imWidthToNextColumn - 30
    plcLayout.Left = plcLib.Left + plcLib.Width + 120
    plcLayout.Width = imStartXFirstColumn + 7 * imWidthToNextColumn + 2 * fgBevelX
    pbcLayout.Width = plcLayout.Width - 2 * fgBevelX
    pbcLayout.Picture = LoadPicture("")
    pbcLayout.BackColor = LIGHTYELLOW
    pbcCount.Width = pbcLayout.Width
    pbcCount.Picture = LoadPicture("")
    pbcCount.BackColor = LIGHTYELLOW
    pbcLibrary.Width = pbcLayout.Width
    pbcLibrary.Picture = LoadPicture("")
    pbcLibrary.BackColor = LIGHTYELLOW
    Program.Width = plcLayout.Left + plcLayout.Width + 2 * vbcLayout.Width

    imStartYFirstColumn = 210
    imHeightWithinColumn = fgBoxGridH
    imHeightWithinColumn = CLng(fmAdjFactorH * imHeightWithinColumn)
    Do While (imHeightWithinColumn Mod 15) <> 0
        imHeightWithinColumn = imHeightWithinColumn + 1
    Loop
    plcLayout.height = imStartYFirstColumn + 24 * (imHeightWithinColumn + 15) + 2 * fgBevelY
    pbcLayout.height = plcLayout.height - 2 * fgBevelY
    pbcCount.height = pbcLayout.height
    pbcLibrary.height = pbcLayout.height
    plcLayout.height = plcLayout.height + hbcDate.height
    vbcLayout.height = pbcLayout.height
    plcDate.Top = plcLayout.height - plcDate.height - fgBevelY
    hbcDate.Top = plcDate.Top
    plcWkCnt.Top = plcDate.Top
    lbcEvents.Top = plcLayout.Top + plcLayout.height + 30
    pbcCurrent.Top = lbcEvents.Top
    cmcDone.Top = lbcEvents.Top
    cmcPrgName.Top = lbcEvents.Top
    cmcDupl.Top = lbcEvents.Top
    cmcDated.Top = cmcDone.Top + cmcDone.height + 30
    cmcSchedule.Top = cmcDated.Top
    cmcReport.Top = cmcDated.Top
    imcTrash.Top = cmcDone.Top
    plcResol.Top = cmcDated.Top
    plcLib.height = plcLayout.height
    cmcDefSchd.Top = plcLib.height - fgBevelY - cmcDefSchd.height - 30
    cmcLib.Top = cmcDefSchd.Top - cmcLib.height - 60
    cmcDefSchd.Left = plcLib.Width / 2 - cmcDefSchd.Width / 2
    cmcLib.Left = cmcDefSchd.Left
    lbcLib.height = plcLib.height - 2 * cmcLib.height - 120
    lacLibFrame(0).Width = imWidthWithinColumn
    lacLibFrame(1).Width = imWidthWithinColumn
    Program.height = cmcReport.Top + 5 * cmcReport.height / 3 - 45

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLibPop                         *
'*                                                     *
'*             Created:10/11/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection library *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mLibPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slType As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVer As Integer
    Dim slDate As String
    Dim llLLD As Long
    Dim llNow As Long
    Dim slStr As String

    Screen.MousePointer = vbHourglass  'Wait
    lbcEvents.Clear
    lbcLib.Clear
    'lbcLibName.Clear
    'lbcLibName.Tag = ""
    ReDim tmLibName(0 To 0) As SORTCODE
    smLibNameTag = ""
    'If (igVehIndexViaPrg < 0) Or (igVehIndexViaPrg > Traffic!lbcVehicle.ListCount - 1) Then
    If (igVehIndexViaPrg < 0) Or (igVehIndexViaPrg > UBound(tmPrgVehicle) - 1) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    slNameCode = tmPrgVehicle(igVehIndexViaPrg).sKey  'Traffic!lbcVehicle.List(igVehIndexViaPrg)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mLibPopErr
    gCPErrorMsg ilRet, "mLibPop (gParseItem field 2: Vehicle)", Program
    On Error GoTo 0
    If imVefCode <> Val(slCode) Then
        imVefCode = Val(slCode)
        tmVefSrchKey.iCode = imVefCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mLibPopErr
        gBtrvErrorMsg ilRet, "mLibPop (btrGetEqual: Vef.Btr)", Program
        On Error GoTo 0
        sgVefTypeViaPrg = tmVef.sType
        imVpfIndex = gVpfFind(Program, imVefCode)
        If sgVefTypeViaPrg = "C" Then
            'If tgVpf(imVpfIndex).iGMnfNCode(1) > 0 Then
            If tgVpf(imVpfIndex).iGMnfNCode(0) > 0 Then
                sgVefTypeViaPrg = "CF"
            End If
        End If
        gUnpackDate tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), slDate
        If Trim$(slDate) = "" Then
            'Until the first log is generated-allow any date
            lgEarliestDateViaPrg = 0
            If tmVef.sType = "V" Then   'Since Virtual vehicles don't has a lld, use todays date
                slDate = Format$(gNow(), "m/d/yy")
                llNow = gDateValue(slDate)
                lgEarliestDateViaPrg = llNow + 1
            End If
        Else
            llLLD = gDateValue(slDate)
            slDate = Format$(gNow(), "m/d/yy")
            llNow = gDateValue(slDate)
            If llNow < llLLD Then
                If tgVpf(imVpfIndex).sMoveLLD <> "Y" Then
                    lgEarliestDateViaPrg = llLLD + 1
                Else
                    lgEarliestDateViaPrg = llNow + 1
                End If
            Else
                lgEarliestDateViaPrg = llNow + 1
            End If
        End If
        mSetWinStatus
    End If
    If tmVef.sState = "D" Then
        slStr = ", Dormant"
        lacType.ForeColor = RED
    Else
        slStr = ", Active"
        lacType.ForeColor = DARKGREEN
    End If
    Select Case tmVef.sType
        Case "C"
            lacType.Caption = "Conventional" & slStr
        Case "S"
            lacType.Caption = "Selling" & slStr
        Case "A"
            lacType.Caption = "Airing" & slStr
        Case "L"
            lacType.Caption = "Log" & slStr
        Case "V"
            lacType.Caption = "Virtual" & slStr
        Case "T"
            lacType.Caption = "Simulcast" & slStr
        Case "P"
            lacType.Caption = "Package" & slStr
        Case "G"
            lacType.Caption = "Sport" & slStr
    End Select
    If rbcType(0).Value Then
        'If igLibType = 3 Then 'Std Format
        '    slType = "F"
        'ElseIf igLibType = 2 Then 'Sports
        '    slType = "P"
        'ElseIf igLibType = 1 Then 'Special
        '    slType = "S"
        'Else    'Regular
            slType = "R"
        'End If
        If ckcShowVersion.Value = vbChecked Then
            ilVer = ALLLIBFRONT
        Else
            ilVer = LATESTLIB
        End If
        'ilRet = gPopProgLibBox(Program, ilVer, slType, imVefCode, lbcLib, lbcLibName)
        ilRet = gPopProgLibBox(Program, ilVer, slType, imVefCode, lbcLib, tmLibName(), smLibNameTag)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mLibPopErr
            gCPErrorMsg ilRet, "mLibPop (gPopProgLibBox: Library)", Program
            On Error GoTo 0
            lbcLib.AddItem "[New]", 0  'Force as first item on list
        End If
    Else
        'ilRet = gPopEvtNmByTypeBox(Program, True, True, lbcLib, lbcETypeCode)
        ilRet = gPopEvtNmByTypeBox(Program, True, True, lbcLib, tmETypeCode(), smETypeCodeTag)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mLibPopErr
            gCPErrorMsg ilRet, "mLibPop (gPopEvtNmByTypeBox: Event Type)", Program
            On Error GoTo 0
        End If
    End If
    If tmVef.sType = "G" Then
        If lbcLib.ListCount > 1 Then
            cmcDefSchd.Enabled = True
        Else
            cmcDefSchd.Enabled = False
        End If
        tmGhfSrchKey1.iVefCode = tmVef.iCode
        ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet <> BTRV_ERR_NONE Then
            tmGhf.lCode = -1
        End If
    Else
        cmcDefSchd.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mLibPopErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gMoveLLC                        *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Move LLC image to another LLC   *
'*                                                     *
'*******************************************************
Private Sub mMoveLLC(tlFromLLC As LLC, tlToLLC As LLC)
    tlToLLC.iDay = tlFromLLC.iDay         'Day index 1-7
    tlToLLC.sType = tlFromLLC.sType     'Library (R,P, or S) or Event type (1, 2,..D or Y)
    tlToLLC.sStartTime = tlFromLLC.sStartTime    'Start time of library or event
    tlToLLC.sLength = tlFromLLC.sLength       'Length of library or event
    tlToLLC.iUnits = tlFromLLC.iUnits       'Units if avail; MnfExcl # 1 if program
    tlToLLC.sName = tlFromLLC.sName        'Library name (version#/Name-Variation)
    tlToLLC.lLvfCode = tlFromLLC.lLvfCode        'Lvf Code (used by gBuildEventDay only)
    tlToLLC.iLtfCode = tlFromLLC.iLtfCode     'Ltf Code if avail ; MnfExcl # 1 if program (used by gBuildEventDay only)
    tlToLLC.iAvailInfo = tlFromLLC.iAvailInfo   'Avail flags (used by gBuildEventDay only); iAnfCode (used by LinksDef->mReadLcfLefLnf & mMoveEvents)
    tlToLLC.iEtfCode = tlFromLLC.iEtfCode
    tlToLLC.iEnfCode = tlFromLLC.iEnfCode
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintCounts                    *
'*                                                     *
'*             Created:10/11/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint counts within days       *
'*                                                     *
'*******************************************************
Private Sub mPaintCounts(clStartTime As Currency, clTimeInc As Currency, tlLLC() As LLC)
'
'   mPaintCounts clStartTime, clTimeInc, tmCLLC()
'   Where:
'       clStartTime (I)- Start time
'       clTimeInc (I) - Time increment (3600)
'       tmCLLC() (I)- event records to be processed
'
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilIndex As Integer
    Dim clGStartTime As Currency
    Dim clGEndTime As Currency
    Dim clEvtTime As Currency
    Dim ilNoEvt As Integer
    Dim ilUnits As Integer
    Dim clSec As Currency
    Dim ilFindEvt As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slEvtType As String
    Dim ilRet As Integer
    Dim llColor As Long
    If lbcLib.ListIndex < 0 Then
        Exit Sub
    End If
    slNameCode = tmETypeCode(lbcLib.ListIndex).sKey    'lbcETypeCode.List(lbcLib.ListIndex)
    ilRet = gParseItem(slNameCode, 3, "\", slCode)
    If ilRet <> CP_MSG_NONE Then
        Exit Sub
    End If
    llColor = pbcCount.ForeColor
    If imCurrent = 0 Then
        pbcCount.ForeColor = BLUE 'BLACK 'GREEN
    Else
        pbcCount.ForeColor = RED
    End If
    Select Case Val(slCode)
        Case 1  'Program
            slEvtType = "1"
        Case 2  'Contract Avail
            slEvtType = "2"
        Case 3
            slEvtType = "3"
        Case 4
            slEvtType = "4"
        Case 5
            slEvtType = "5"
        Case 6  'Cmml Promo
            slEvtType = "6"
        Case 7  'Feed avail
            slEvtType = "7"
        Case 8  'PSA/Promo (Avail)
            slEvtType = "8"
        Case 9
            slEvtType = "9"
        Case 10  'Page eject, Line space 1, 2 or 3
            slEvtType = "A"
        Case 11
            slEvtType = "B"
        Case 12
            slEvtType = "C"
        Case 13
            slEvtType = "D"
        Case Else   'Other
            slEvtType = "Y"
    End Select
    ilIndex = LBound(tlLLC)
    flX = imStartXFirstColumn
    For ilCol = 1 To 7 Step 1
        flY = imStartYFirstColumn   '210
        clGStartTime = clStartTime
        clGEndTime = clStartTime + clTimeInc - 1
        ilNoEvt = 0
        ilUnits = 0
        clSec = 0
        'Only required if not showing 24 hours
        If (tlLLC(ilIndex).iDay = -1) Or (ilIndex >= UBound(tlLLC)) Then
            Exit Sub
        End If
        Do While (tlLLC(ilIndex).iDay < ilCol)
            ilIndex = ilIndex + 1
            If (tlLLC(ilIndex).iDay = -1) Or (ilIndex >= UBound(tlLLC)) Then
                Exit Sub
            End If
        Loop
        If tlLLC(ilIndex).iDay = ilCol Then
            For ilRow = 1 To 24 Step 1
                ilFindEvt = True
                Do
                    clEvtTime = gTimeToCurrency(tlLLC(ilIndex).sStartTime, False)
                    If tlLLC(ilIndex).iDay = ilCol Then
                        If (clEvtTime >= clGStartTime) And (clEvtTime <= clGEndTime) Then
                            If tlLLC(ilIndex).sType = slEvtType Then
                                Select Case slEvtType
                                    Case "1"  'Program
                                        ilNoEvt = ilNoEvt + 1
                                    Case "2", "3", "4", "5"  'Contract Avail
                                        ilUnits = ilUnits + tlLLC(ilIndex).iUnits
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            clSec = clSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            clSec = clSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                        End If
'                                    Case "3"    'Open BB
'                                        ilNoEvt = ilNoEvt + 1
'                                    Case "4"    'Floating BB
'                                        ilNoEvt = ilNoEvt + 1
'                                    Case "5"    'Close BB
'                                        ilNoEvt = ilNoEvt + 1
                                    Case "6"  'Cmml Promo
                                        ilUnits = ilUnits + tlLLC(ilIndex).iUnits
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            clSec = clSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            clSec = clSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                        End If
                                    Case "7"  'Feed avail
                                        ilUnits = ilUnits + tlLLC(ilIndex).iUnits
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            clSec = clSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            clSec = clSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                        End If
                                    Case "8"  'PSA/Promo (Avail)
                                        ilUnits = ilUnits + tlLLC(ilIndex).iUnits
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            clSec = clSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            clSec = clSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                        End If
                                    Case "9"
                                        ilUnits = ilUnits + tlLLC(ilIndex).iUnits
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            clSec = clSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            clSec = clSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                        End If
                                    Case "A"  'Page eject, Line space 1, 2 or 3
                                        ilNoEvt = ilNoEvt + 1
                                    Case "B"
                                        ilNoEvt = ilNoEvt + 1
                                    Case "C"
                                        ilNoEvt = ilNoEvt + 1
                                    Case "D"
                                        ilNoEvt = ilNoEvt + 1
                                    Case "Y"   'Other
                                        ilNoEvt = ilNoEvt + 1
                                End Select
                            End If
                            ilIndex = ilIndex + 1
                            If ilIndex >= UBound(tlLLC) Then
                                ilFindEvt = False
                            End If
                        Else
                            If clEvtTime > clGEndTime Then
                                ilFindEvt = False
                            Else
                                ilIndex = ilIndex + 1
                                If ilIndex >= UBound(tlLLC) Then
                                    ilFindEvt = False
                                End If
                            End If
                        End If
                    Else
                        ilFindEvt = False
                    End If
                Loop While ilFindEvt
                pbcCount.CurrentX = flX
                pbcCount.CurrentY = flY
                Select Case slEvtType
                    Case "1"  'Program
                        If ilNoEvt > 0 Then
                            pbcCount.Print Trim$(str$(ilNoEvt))
                        End If
                    Case "2", "3", "4", "5"  'Contract Avail
                        If ilUnits > 0 Then
                            slStr = Trim$(str$(ilUnits))
                            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                slStr = slStr & "/" & gCurrencyToLength(clSec)
                            ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                slStr = slStr & "/" & gCurrencyToLength(clSec)
                            End If
                            pbcCount.Print slStr
                        End If
'                    Case "3"    'Open BB
'                        If ilNoEvt > 0 Then
'                            slStr = Trim$(Str$(ilNoEvt))
'                            pbcCount.Print slStr
'                        End If
'                    Case "4"    'Floating BB
'                        If ilNoEvt > 0 Then
'                            slStr = Trim$(Str$(ilNoEvt))
'                            pbcCount.Print slStr
'                        End If
'                    Case "5"    'Close BB
'                        If ilNoEvt > 0 Then
'                            slStr = Trim$(Str$(ilNoEvt))
'                            pbcCount.Print slStr
'                        End If
                    Case "6"  'Cmml Promo
                        If ilUnits > 0 Then
                            slStr = Trim$(str$(ilUnits))
                            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                slStr = slStr & "/" & gCurrencyToLength(clSec)
                            ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                slStr = slStr & "/" & gCurrencyToLength(clSec)
                            End If
                            pbcCount.Print slStr
                        End If
                    Case "7"  'Feed avail
                        If ilUnits > 0 Then
                            slStr = Trim$(str$(ilUnits))
                            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                slStr = slStr & "/" & gCurrencyToLength(clSec)
                            ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                slStr = slStr & "/" & gCurrencyToLength(clSec)
                            End If
                            pbcCount.Print slStr
                        End If
                    Case "8"  'PSA/Promo (Avail)
                        If ilUnits > 0 Then
                            slStr = Trim$(str$(ilUnits))
                            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                slStr = slStr & "/" & gCurrencyToLength(clSec)
                            ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                slStr = slStr & "/" & gCurrencyToLength(clSec)
                            End If
                            pbcCount.Print slStr
                        End If
                    Case "9"
                        If ilUnits > 0 Then
                            slStr = Trim$(str$(ilUnits))
                            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                slStr = slStr & "/" & gCurrencyToLength(clSec)
                            ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                slStr = slStr & "/" & gCurrencyToLength(clSec)
                            End If
                            pbcCount.Print slStr
                        End If
                    Case "A"  'Page eject, Line space 1, 2 or 3
                        If ilNoEvt > 0 Then
                            slStr = Trim$(str$(ilNoEvt))
                            pbcCount.Print slStr
                        End If
                    Case "B"
                        If ilNoEvt > 0 Then
                            slStr = Trim$(str$(ilNoEvt))
                            pbcCount.Print slStr
                        End If
                    Case "C"
                        If ilNoEvt > 0 Then
                            slStr = Trim$(str$(ilNoEvt))
                            pbcCount.Print slStr
                        End If
                    Case "D"
                        If ilNoEvt > 0 Then
                            slStr = Trim$(str$(ilNoEvt))
                            pbcCount.Print slStr
                        End If
                    Case "Y"   'Other
                        If ilNoEvt > 0 Then
                            slStr = Trim$(str$(ilNoEvt))
                            pbcCount.Print slStr
                        End If
                End Select
                If ilIndex >= UBound(tlLLC) Then
                    Exit Sub
                End If
                ilNoEvt = 0
                ilUnits = 0
                clSec = 0
                flY = flY + imHeightWithinColumn + 15  'fgBoxGridH
                clGStartTime = clGStartTime + clTimeInc
                clGEndTime = clGStartTime + clTimeInc - 1
            Next ilRow
        End If
        flX = flX + imWidthToNextColumn
    Next ilCol
    pbcCount.ForeColor = llColor
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintDates                     *
'*                                                     *
'*             Created:10/11/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint dates as titles          *
'*                                                     *
'*******************************************************
Private Sub mPaintDates(pbcCtrl As control)
    Dim ilCol As Integer
    Dim llColor As Long
    Dim flX As Single
    Dim flY As Single
    Dim slDate As String
    Dim llDate As Long
    Dim ilIndex As Integer

    llColor = pbcCtrl.ForeColor
    pbcCtrl.ForeColor = BLUE
    flX = imStartXFirstColumn
    flY = 30
    If plcDate.Caption = "" Then
        Exit Sub
    End If
    For ilCol = 1 To 7 Step 1
        If plcDate.Caption = "TFN" Then
            Select Case ilCol
                Case 1
                    slDate = "Monday"
                Case 2
                    slDate = "Tuesday"
                Case 3
                    slDate = "Wednesday"
                Case 4
                    slDate = "Thursday"
                Case 5
                    slDate = "Friday"
                Case 6
                    slDate = "Saturday"
                Case 7
                    slDate = "Sunday"
            End Select
        Else
            If tmVef.sType <> "G" Then
                slDate = plcDate.Caption
                llDate = gDateValue(slDate) + ilCol - 1
                If lgEarliestDateViaPrg > 0 Then
                    If llDate < lgEarliestDateViaPrg Then
                        pbcCtrl.ForeColor = RED
                    End If
                End If
                slDate = Format$(llDate, "m/d/yy")
                Select Case ilCol
                    Case 1
                        slDate = "Mo: " & slDate
                    Case 2
                        slDate = "Tu: " & slDate
                    Case 3
                        slDate = "We: " & slDate
                    Case 4
                        slDate = "Th: " & slDate
                    Case 5
                        slDate = "Fr: " & slDate
                    Case 6
                        slDate = "Sa: " & slDate
                    Case 7
                        slDate = "Su: " & slDate
                End Select
            Else
                ilIndex = 7 * (hbcDate.Value - hbcDate.Min) + ilCol - 1
                If ilIndex < UBound(tmGsfInfo) Then
                    llDate = tmGsfInfo(ilIndex).lGameDate
                    slDate = Format(llDate, "m/d/yy")
                    slDate = Trim$(str$(tmGsfInfo(ilIndex).iGameNo)) & "-" & slDate
                Else
                    slDate = ""
                End If
            End If
        End If
        pbcCtrl.CurrentX = flX
        pbcCtrl.CurrentY = flY
'        gPaintArea pbcCtrl, flX, flY, 945 - 15, 165 - 15, LIGHTYELLOW
        pbcCtrl.CurrentX = flX
        pbcCtrl.CurrentY = flY
        pbcCtrl.Print slDate
        flX = flX + imWidthToNextColumn
        pbcCtrl.ForeColor = BLUE
    Next ilCol
    pbcCtrl.ForeColor = llColor
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintLayout                    *
'*                                                     *
'*             Created:10/11/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint layout within days       *
'*                                                     *
'*******************************************************
Private Sub mPaintLayout(clStartTime As Currency, clTimeInc As Currency, tlLLC() As LLC)
'
'   mPaintLayout clStartTime, clTimeInc, tmCLLC()
'   Where:
'       clStartTime (I)- Start time
'       clTimeInc (I) - Time increment (3600)
'       tmCLLC() (I)- event records to be processed
'
    Dim ilCol As Integer
    Dim ilIndex As Integer
    Dim clGStartTime As Currency
    Dim clGEndTime As Currency
    Dim clEvtTime As Currency
    Dim clEvtEndTime As Currency
    Dim slTime As String
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slEvtType As String
    Dim ilRet As Integer
    Dim llColor As Long
    Dim flY1 As Single
    Dim flY2 As Single
    Dim flX1 As Single
    Dim flX2 As Single
    Dim slXMid As String
    Dim llY As Long

    If lbcLib.ListIndex < 0 Then
        Exit Sub
    End If
    slNameCode = tmETypeCode(lbcLib.ListIndex).sKey    'lbcETypeCode.List(lbcLib.ListIndex)
    ilRet = gParseItem(slNameCode, 3, "\", slCode)
    If ilRet <> CP_MSG_NONE Then
        Exit Sub
    End If
    llColor = pbcLayout.ForeColor
    If imCurrent = 0 Then
        pbcLayout.ForeColor = BLUE 'BLACK 'GREEN
    Else
        pbcLayout.ForeColor = RED
    End If
    Select Case Val(slCode)
        Case 1  'Program
            slEvtType = "1"
        Case 2  'Contract Avail
            slEvtType = "2"
        Case 3
            slEvtType = "3"
        Case 4
            slEvtType = "4"
        Case 5
            slEvtType = "5"
        Case 6  'Cmml Promo
            slEvtType = "6"
        Case 7  'Feed avail
            slEvtType = "7"
        Case 8  'PSA/Promo (Avail)
            slEvtType = "8"
        Case 9
            slEvtType = "9"
        Case 10  'Page eject, Line space 1, 2 or 3
            slEvtType = "A"
        Case 11
            slEvtType = "B"
        Case 12
            slEvtType = "C"
        Case 13
            slEvtType = "D"
        Case Else   'Other
            slEvtType = "Y"
    End Select
    ilIndex = LBound(tlLLC)
    flX1 = imStartXFirstColumn
    flX2 = flX1 + imWidthWithinColumn
    For ilCol = 1 To 7 Step 1
        clGStartTime = clStartTime
        clGEndTime = clStartTime + 24 * clTimeInc - 1
        'Only required if not showing 24 hours
        If (tlLLC(ilIndex).iDay = -1) Or (ilIndex >= UBound(tlLLC)) Then
            Exit Sub
        End If
        Do While (tlLLC(ilIndex).iDay < ilCol)
            ilIndex = ilIndex + 1
            If (tlLLC(ilIndex).iDay = -1) Or (ilIndex >= UBound(tlLLC)) Then
                Exit Sub
            End If
        Loop
        Do While tlLLC(ilIndex).iDay = ilCol
            If tmVef.sType <> "G" Then
                clEvtTime = gTimeToCurrency(tlLLC(ilIndex).sStartTime, False)
                gAddTimeLength tlLLC(ilIndex).sStartTime, tlLLC(ilIndex).sLength, "A", "1", slTime, slXMid
            Else
                clEvtTime = gTimeToCurrency("12am", False)
                gAddTimeLength "12am", tlLLC(ilIndex).sLength, "A", "1", slTime, slXMid
            End If
            clEvtEndTime = gTimeToCurrency(slTime, True)
            If (clEvtTime >= clGStartTime) And (clEvtTime <= clGEndTime) Or (clEvtEndTime >= clGStartTime) And (clEvtEndTime <= clGEndTime) Then
                If tlLLC(ilIndex).sType = slEvtType Then
                    Select Case slEvtType
                        Case "1"  'Program
                            'flY1 = fmPaintHeight * (clEvtTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                            'If flY1 <= (imStartYFirstColumn - 15) Then
                            '    flY1 = imStartYFirstColumn
                            'End If
                            'flY2 = fmPaintHeight * (clEvtEndTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                            'If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                            '    flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                            'End If
                            mComputeXY gTimeToLong(gCurrencyToTime(clEvtTime), False), False, llY
                            If llY <> -1 Then
                                flY1 = llY
                                If flY1 <= (imStartYFirstColumn - 15) Then
                                    flY1 = imStartYFirstColumn
                                End If
                                mComputeXY gTimeToLong(gCurrencyToTime(clEvtEndTime + 1), True), True, llY
                                If llY <> -1 Then
                                    flY2 = llY
                                    If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                                        flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                                    End If
                                    pbcLayout.Line (flX1 + 15, flY1)-(flX2 - 15, flY2), , B
        '                            gPaintArea pbcLayout, flX1, flY1, flY2 - flY1, flX2 - flX2, BLUE 'BLACK
                                    pbcLayout.CurrentX = flX1 + fgBoxInsetX
                                    pbcLayout.CurrentY = flY1
                                    slStr = Trim$(tlLLC(ilIndex).sName)
                                    gAdjShowLen pbcLayout, slStr, CSng(imWidthWithinColumn)
                                    pbcLayout.Print slStr
                                End If
                            End If
                        Case "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D"
                            'flY1 = fmPaintHeight * (clEvtTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                            'If flY1 <= (imStartYFirstColumn - 15) Then
                            '    flY1 = imStartYFirstColumn
                            'End If
                            'If flY1 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                            '    flY1 = fmPaintHeight + (imStartYFirstColumn - 15)
                            'End If
                            mComputeXY gTimeToLong(gCurrencyToTime(clEvtTime), False), False, llY
                            If llY <> -1 Then
                                flY1 = llY
                                If flY1 <= (imStartYFirstColumn - 15) Then
                                    flY1 = imStartYFirstColumn
                                End If
                                If flY1 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                                    flY1 = fmPaintHeight + (imStartYFirstColumn - 15)
                                End If
                                pbcLayout.Line (flX1, flY1)-(flX2, flY1), , B
                            End If
'                            gPaintArea pbcLayout, flX1, flY1, 0, flX2 - flX2, BLUE 'BLACK
                        Case "Y"   'Other
                            'flY1 = fmPaintHeight * (clEvtTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                            'If flY1 <= (imStartYFirstColumn - 15) Then
                            '    flY1 = imStartYFirstColumn
                            'End If
                            'flY2 = fmPaintHeight * (clEvtEndTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                            'If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                            '    flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                            'End If
                            mComputeXY gTimeToLong(gCurrencyToTime(clEvtTime), False), False, llY
                            If llY <> -1 Then
                                flY1 = llY
                                If flY1 <= (imStartYFirstColumn - 15) Then
                                    flY1 = imStartYFirstColumn
                                End If
                                mComputeXY gTimeToLong(gCurrencyToTime(clEvtEndTime + 1), True), True, llY
                                If llY <> -1 Then
                                    flY2 = llY
                                    If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                                        flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                                    End If
                                    pbcLayout.Line (flX1, flY1)-(flX2, flY2), , B
'                                    gPaintArea pbcLayout, flX1, flY1, flY2 - flY1, flX2 - flX2, BLUE 'BLACK
                                    pbcLayout.CurrentX = flX1 + fgBoxInsetX \ 2
                                    pbcLayout.CurrentY = flY1
                                    slStr = Trim$(tlLLC(ilIndex).sName)
                                    gAdjShowLen pbcLayout, slStr, CSng(imWidthWithinColumn)
                                    pbcLayout.Print slStr
                                End If
                            End If
                    End Select
                End If
            End If
            ilIndex = ilIndex + 1
            If (tlLLC(ilIndex).iDay = -1) Or (ilIndex >= UBound(tlLLC)) Then
                Exit Sub
            End If
        Loop
        flX1 = flX1 + imWidthToNextColumn
        flX2 = flX1 + imWidthWithinColumn
    Next ilCol
    pbcLayout.ForeColor = llColor
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintLibrary                   *
'*                                                     *
'*             Created:10/11/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint linbrary within days     *
'*                                                     *
'*******************************************************
Private Sub mPaintLibrary(clStartTime As Currency, clTimeInc As Currency, tlCLLC() As LLC, tlPLLC() As LLC)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llDelta                                                                               *
'******************************************************************************************

'
'   mPaintLibrary clStartTime, clTimeInc, tmCLLC(), tmPLLC()
'   Where:
'       clStartTime (I)- Start time
'       clTimeInc (I) - Time increment (3600)
'       tmCLLC() (I)- event records to be processed
'
    Dim ilCol As Integer
    Dim ilCIndex As Integer
    Dim ilPIndex
    Dim clGStartTime As Currency
    Dim clGEndTime As Currency
    Dim clCEvtTime As Currency
    Dim clCEvtEndTime As Currency
    Dim clPEvtTime As Currency
    Dim clPEvtEndTime As Currency
    Dim slTime As String
    Dim ilCFindEvt As Integer
    Dim ilPFindEvt As Integer
    Dim slStr As String
    Dim llColor As Long
    Dim flY1 As Single
    Dim flY2 As Single
    Dim flX1 As Single
    Dim flX2 As Single
    Dim ilPos As Integer
    Dim ilUpper As Integer
    Dim slGameTime As String
    Dim slXMid As String
    Dim llY As Long
    Dim ilIndex As Integer
    ReDim tmLCD(0 To 1) As LCD
    ilUpper = 1
    llColor = pbcLibrary.ForeColor
    ilCIndex = LBound(tlCLLC)
    ilPIndex = LBound(tlPLLC)
    flX1 = imStartXFirstColumn
    flX2 = flX1 + imWidthWithinColumn
    For ilCol = 1 To 7 Step 1
        clGStartTime = clStartTime
        clGEndTime = clStartTime + 24 * clTimeInc - 1
        Do While (tlCLLC(ilCIndex).iDay = ilCol) Or ((tlPLLC(ilPIndex).iDay = ilCol) And (imCurrent <> 0))
            ilCFindEvt = False
            Do While tlCLLC(ilCIndex).iDay = ilCol
                If (tlCLLC(ilCIndex).sType = "R") Or (tlCLLC(ilCIndex).sType = "S") Or (tlCLLC(ilCIndex).sType = "P") Then
                    If tmVef.sType <> "G" Then
                        clCEvtTime = gTimeToCurrency(tlCLLC(ilCIndex).sStartTime, False)
                        gAddTimeLength tlCLLC(ilCIndex).sStartTime, tlCLLC(ilCIndex).sLength, "A", "1", slTime, slXMid
                    Else
                        clCEvtTime = gTimeToCurrency("12am", False)
                        gAddTimeLength "12am", tlCLLC(ilCIndex).sLength, "A", "1", slTime, slXMid
                    End If
                    clCEvtEndTime = gTimeToCurrency(slTime, True) - 1
                    If (clCEvtEndTime < clGStartTime) Or (clCEvtTime > clGEndTime) Then
                        If (tlCLLC(ilCIndex).iDay = -1) Or (ilCIndex >= UBound(tlCLLC)) Then
                            Exit Do
                        End If
                        ilCIndex = ilCIndex + 1
                    Else
                        ilCFindEvt = True
                        Exit Do
                    End If
                Else
                    If (tlCLLC(ilCIndex).iDay = -1) Or (ilCIndex >= UBound(tlCLLC)) Then
                        Exit Do
                    End If
                    ilCIndex = ilCIndex + 1
                End If
            Loop

            ilPFindEvt = False
            If imCurrent <> 0 Then
                Do While tlPLLC(ilPIndex).iDay = ilCol
                    If (tlPLLC(ilPIndex).sType = "R") Or (tlPLLC(ilPIndex).sType = "S") Or (tlPLLC(ilPIndex).sType = "P") Then
                        clPEvtTime = gTimeToCurrency(tlPLLC(ilPIndex).sStartTime, False)
                        gAddTimeLength tlPLLC(ilPIndex).sStartTime, tlPLLC(ilPIndex).sLength, "A", "1", slTime, slXMid
                        clPEvtEndTime = gTimeToCurrency(slTime, True) - 1
                        If (clPEvtEndTime < clGStartTime) Or (clPEvtTime > clGEndTime) Then
                            If (tlPLLC(ilPIndex).iDay = -1) Or (ilPIndex >= UBound(tlPLLC)) Then
                                Exit Do
                            End If
                            ilPIndex = ilPIndex + 1
                        Else
                            ilPFindEvt = True
                            Exit Do
                        End If
                    Else
                        If (tlPLLC(ilPIndex).iDay = -1) Or (ilPIndex >= UBound(tlPLLC)) Then
                            Exit Do
                        End If
                        ilPIndex = ilPIndex + 1
                    End If
                Loop
            End If
            If ilCFindEvt And ilPFindEvt Then
                If clCEvtEndTime < clPEvtTime Then
                    pbcLibrary.ForeColor = BLUE 'BLACK 'GREEN
                    'flY1 = fmPaintHeight * (clCEvtTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                    'Do While flY1 Mod 15 <> 0
                    '    flY1 = flY1 + 1
                    'Loop
                    'If flY1 <= (imStartYFirstColumn - 15) Then
                    '    flY1 = imStartYFirstColumn
                    'End If
                    'flY2 = fmPaintHeight * (clCEvtEndTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                    'Do While flY2 Mod 15 <> 0
                    '    flY2 = flY2 + 1
                    'Loop
                    'If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                    '    flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                    'End If
                    mComputeXY gTimeToLong(gCurrencyToTime(clCEvtTime), False), False, llY
                    If llY <> -1 Then
                        flY1 = llY
                        If flY1 <= (imStartYFirstColumn - 15) Then
                            flY1 = imStartYFirstColumn
                        End If
                        mComputeXY gTimeToLong(gCurrencyToTime(clCEvtEndTime + 1), True), True, llY
                        If llY <> -1 Then
                            flY2 = llY
                            If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                                flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                            End If
                            pbcLibrary.Line (flX1 + 15, flY1)-(flX2 - 15, flY2), , B
                            pbcLibrary.CurrentX = flX1 + fgBoxInsetX
                            pbcLibrary.CurrentY = flY1
                            slStr = Trim$(tlCLLC(ilCIndex).sName)
                            If ckcShowVersion.Value = vbUnchecked Then
                                ilPos = InStr(slStr, "/")
                                If ilPos > 0 Then
                                    slStr = Mid$(slStr, ilPos + 1)
                                End If
                            End If
                            gAdjShowLen pbcLibrary, slStr, CSng(imWidthWithinColumn)
                            pbcLibrary.Print slStr
                            tmLCD(ilUpper).sName = tlCLLC(ilCIndex).sName
                            tmLCD(ilUpper).iDay = ilCol
                            tmLCD(ilUpper).iCurOrPend = 0
                            tmLCD(ilUpper).iLLCIndex = ilCIndex
                            tmLCD(ilUpper).fX1 = flX1
                            tmLCD(ilUpper).fX2 = flX2
                            tmLCD(ilUpper).fY1 = flY1
                            tmLCD(ilUpper).fY2 = flY2
                            ilUpper = ilUpper + 1
                            ReDim Preserve tmLCD(0 To ilUpper) As LCD
                        End If
                    End If
                    ilCIndex = ilCIndex + 1
                ElseIf clPEvtEndTime < clCEvtTime Then
                    pbcLibrary.ForeColor = RED
                    'flY1 = fmPaintHeight * (clPEvtTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                    'Do While flY1 Mod 15 <> 0
                    '    flY1 = flY1 + 1
                    'Loop
                    'If flY1 <= (imStartYFirstColumn - 15) Then
                    '    flY1 = imStartYFirstColumn
                    'End If
                    'flY2 = fmPaintHeight * (clPEvtEndTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                    'Do While flY2 Mod 15 <> 0
                    '    flY2 = flY2 + 1
                    'Loop
                    'If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                    '    flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                    'End If
                    mComputeXY gTimeToLong(gCurrencyToTime(clPEvtTime), False), False, llY
                    If llY <> -1 Then
                        flY1 = llY
                        If flY1 <= (imStartYFirstColumn - 15) Then
                            flY1 = imStartYFirstColumn
                        End If
                        mComputeXY gTimeToLong(gCurrencyToTime(clPEvtEndTime + 1), True), True, llY
                        If llY <> -1 Then
                            flY2 = llY
                            If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                                flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                            End If
                            pbcLibrary.Line (flX1 + 15, flY1)-(flX2 - 15, flY2), , B
                            pbcLibrary.CurrentX = flX1 + fgBoxInsetX
                            pbcLibrary.CurrentY = flY1
                            slStr = Trim$(tlPLLC(ilPIndex).sName)
                            If ckcShowVersion.Value = vbUnchecked Then
                                ilPos = InStr(slStr, "/")
                                If ilPos > 0 Then
                                    slStr = Mid$(slStr, ilPos + 1)
                                End If
                            End If
                            gAdjShowLen pbcLibrary, slStr, CSng(imWidthWithinColumn)
                            pbcLibrary.Print slStr
                            tmLCD(ilUpper).sName = tlPLLC(ilPIndex).sName
                            tmLCD(ilUpper).iDay = ilCol
                            tmLCD(ilUpper).iCurOrPend = 1
                            tmLCD(ilUpper).iLLCIndex = ilPIndex
                            tmLCD(ilUpper).fX1 = flX1
                            tmLCD(ilUpper).fX2 = flX2
                            tmLCD(ilUpper).fY1 = flY1
                            tmLCD(ilUpper).fY2 = flY2
                            ilUpper = ilUpper + 1
                            ReDim Preserve tmLCD(0 To ilUpper) As LCD
                        End If
                    End If
                    ilPIndex = ilPIndex + 1
                Else
                    ilCIndex = ilCIndex + 1
                End If
            ElseIf ilCFindEvt And Not ilPFindEvt Then
                pbcLibrary.ForeColor = BLUE 'BLACK 'GREEN
                'flY1 = fmPaintHeight * (CDbl(clCEvtTime - clGStartTime) / (clGEndTime + 1 - clGStartTime))
                'flY2 = fmPaintHeight * (CDbl(clCEvtEndTime + 1 - clGStartTime) / (clGEndTime + 1 - clGStartTime))
                'llDelta = flY2 - flY1
                'Do While llDelta Mod 15 <> 0
                '    llDelta = llDelta + 1
                'Loop
                'flY1 = CLng(flY1)
                'Do While flY1 Mod 15 <> 0
                '    flY1 = flY1 + 1
                'Loop
                'flY1 = flY1 + (imStartYFirstColumn - 15)
                'If flY1 <= (imStartYFirstColumn - 15) Then
                '    flY1 = imStartYFirstColumn
                'End If
                'flY2 = flY1 + llDelta
                'Do While flY2 Mod 15 <> 0
                '    flY2 = flY2 + 1
                'Loop
                'If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                '    flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                'End If
                mComputeXY gTimeToLong(gCurrencyToTime(clCEvtTime), False), False, llY
                If llY <> -1 Then
                    flY1 = llY
                    If flY1 <= (imStartYFirstColumn - 15) Then
                        flY1 = imStartYFirstColumn
                    End If
                    mComputeXY gTimeToLong(gCurrencyToTime(clCEvtEndTime + 1), True), True, llY
                    If llY <> -1 Then
                        flY2 = llY
                        If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                            flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                        End If
                        ilIndex = -1
                        '12/29/12: fill background in Yellow if cancel event
                        If tmVef.sType = "G" Then
                            ilIndex = 7 * (hbcDate.Value - hbcDate.Min) + ilCol - 1
                            If ilIndex < UBound(tmGsfInfo) Then
                                '12/29/12: Show Canceled Events with an X
                                If tmGsfInfo(ilIndex).sGameStatus <> "C" Then
                                    ilIndex = -1
                                End If
                            Else
                                ilIndex = -1
                            End If
                        End If
                        If ilIndex = -1 Then
                            pbcLibrary.Line (flX1 + 15, flY1)-(flX2 - 15, flY2), , B
                        Else
                            pbcLibrary.Line (flX1 + 15, flY1)-(flX2 - 15, flY2), , B
                            pbcLibrary.Line (flX1 + 30, flY1 + 15)-(flX2 - 30, flY2 - 15), LIGHTYELLOW, BF
                            pbcLibrary.ForeColor = RED
                        End If
                        pbcLibrary.CurrentX = flX1 + fgBoxInsetX
                        pbcLibrary.CurrentY = flY1
                        slStr = Trim$(tlCLLC(ilCIndex).sName)
                        If ckcShowVersion.Value = vbUnchecked Then
                            ilPos = InStr(slStr, "/")
                            If ilPos > 0 Then
                                slStr = Mid$(slStr, ilPos + 1)
                            End If
                        End If
                        gAdjShowLen pbcLibrary, slStr, CSng(imWidthWithinColumn)
                        pbcLibrary.Print slStr
                        If tmVef.sType = "G" Then
                            pbcLibrary.CurrentX = flX1 + fgBoxInsetX
                            slStr = Trim$(tlCLLC(ilCIndex).sVisitName) & " @"
                            pbcLibrary.Print slStr
                            pbcLibrary.CurrentX = flX1 + fgBoxInsetX
                            slStr = Trim$(tlCLLC(ilCIndex).sHomeName)
                            pbcLibrary.Print slStr
                            pbcLibrary.CurrentX = flX1 + fgBoxInsetX
                            slGameTime = Trim$(tlCLLC(ilCIndex).sStartTime)
                            slGameTime = Left$(slGameTime, Len(slGameTime) - 1)
                            slGameTime = LCase$(slGameTime)
                            pbcLibrary.Print slGameTime
                            pbcLibrary.ForeColor = BLUE
                        End If
                        tmLCD(ilUpper).sName = tlCLLC(ilCIndex).sName
                        tmLCD(ilUpper).iDay = ilCol
                        tmLCD(ilUpper).iCurOrPend = 0
                        tmLCD(ilUpper).iLLCIndex = ilCIndex
                        tmLCD(ilUpper).fX1 = flX1
                        tmLCD(ilUpper).fX2 = flX2
                        tmLCD(ilUpper).fY1 = flY1
                        tmLCD(ilUpper).fY2 = flY2
                        ilUpper = ilUpper + 1
                        ReDim Preserve tmLCD(0 To ilUpper) As LCD
                    End If
                End If
                ilCIndex = ilCIndex + 1
            ElseIf Not ilCFindEvt And ilPFindEvt Then
                pbcLibrary.ForeColor = RED
                'flY1 = fmPaintHeight * (clPEvtTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                'Do While flY1 Mod 15 <> 0
                '    flY1 = flY1 + 1
                'Loop
                'If flY1 <= (imStartYFirstColumn - 15) Then
                '    flY1 = imStartYFirstColumn
                'End If
                'flY2 = fmPaintHeight * (clPEvtEndTime - clGStartTime) / (clGEndTime - clGStartTime) + (imStartYFirstColumn - 15)
                'Do While flY2 Mod 15 <> 0
                '    flY2 = flY2 + 1
                'Loop
                'If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                '    flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                'End If
                mComputeXY gTimeToLong(gCurrencyToTime(clPEvtTime), False), False, llY
                If llY <> -1 Then
                    flY1 = llY
                    If flY1 <= (imStartYFirstColumn - 15) Then
                        flY1 = imStartYFirstColumn
                    End If
                    mComputeXY gTimeToLong(gCurrencyToTime(clPEvtEndTime + 1), True), True, llY
                    If llY <> -1 Then
                        flY2 = llY
                        If flY2 >= fmPaintHeight + (imStartYFirstColumn - 15) Then
                            flY2 = fmPaintHeight + (imStartYFirstColumn - 15)
                        End If
                        pbcLibrary.Line (flX1 + 15, flY1)-(flX2 - 15, flY2), , B
                        pbcLibrary.CurrentX = flX1 + fgBoxInsetX
                        pbcLibrary.CurrentY = flY1
                        slStr = Trim$(tlPLLC(ilPIndex).sName)
                        If ckcShowVersion.Value = vbUnchecked Then
                            ilPos = InStr(slStr, "/")
                            If ilPos > 0 Then
                                slStr = Mid$(slStr, ilPos + 1)
                            End If
                        End If
                        gAdjShowLen pbcLibrary, slStr, CSng(imWidthWithinColumn)
                        pbcLibrary.Print slStr
                        tmLCD(ilUpper).sName = tlPLLC(ilPIndex).sName
                        tmLCD(ilUpper).iDay = ilCol
                        tmLCD(ilUpper).iCurOrPend = 1
                        tmLCD(ilUpper).iLLCIndex = ilPIndex
                        tmLCD(ilUpper).fX1 = flX1
                        tmLCD(ilUpper).fX2 = flX2
                        tmLCD(ilUpper).fY1 = flY1
                        tmLCD(ilUpper).fY2 = flY2
                        ilUpper = ilUpper + 1
                        ReDim Preserve tmLCD(0 To ilUpper) As LCD
                    End If
                End If
                ilPIndex = ilPIndex + 1
            End If
        Loop
        If tmVef.sType = "G" Then
            ilIndex = 7 * (hbcDate.Value - hbcDate.Min) + ilCol - 1
            If ilIndex < UBound(tmGsfInfo) Then
                '12/29/12: Show Canceled Events with an X
                If tmGsfInfo(ilIndex).sGameStatus = "C" Then
                    'Yellow Column
                    'pbcLibrary.Line (flX, fgBoxGridH + 30)-(flX1 + imWidthToNextColumn - 30, pbcLibrary.Height - fgBoxGridH - 30), LIGHTYELLOW, BF
                    'X in column
                    pbcLibrary.Line (flX1, fgBoxGridH + 15)-(flX1 + imWidthToNextColumn - 30, pbcLibrary.height - 15), RED
                    pbcLibrary.Line (flX1 + imWidthToNextColumn - 30, fgBoxGridH + 15)-(flX1, pbcLibrary.height - 15), RED
                End If
            End If
        End If
        flX1 = flX1 + imWidthToNextColumn
        flX2 = flX1 + imWidthWithinColumn
    Next ilCol
    pbcLibrary.ForeColor = llColor
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintTimes                     *
'*                                                     *
'*             Created:10/11/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint times down the rows      *
'*                                                     *
'*******************************************************
Private Sub mPaintTimes(ilStartIndex As Integer, ilEndIndex As Integer, ilStep As Integer, pbcCtrl As control)
    Dim ilRow As Integer
    Dim llColor As Long
    Dim flX As Single
    Dim flY As Single
    Dim slTime As String

    llColor = pbcCtrl.ForeColor
    pbcCtrl.ForeColor = BLUE
    flX = 30 '+ fgBoxInsetX
    flY = imStartYFirstColumn
    For ilRow = ilStartIndex To ilEndIndex Step ilStep
        If tmVef.sType <> "G" Then
            If ilRow <= 4 Then
                Select Case ilRow Mod 4
                    Case 1
                        slTime = "12AM"
                    Case 2
                        slTime = "  :15" '"12:15"
                    Case 3
                        slTime = "  :30" '"12:30"
                    Case Else
                        slTime = "  :45" '"12:45"
                End Select
            ElseIf (ilRow - 1) \ 4 < 12 Then
                Select Case ilRow Mod 4
                    Case 1
                        slTime = Trim$(str$((ilRow - 1) \ 4)) & "AM"
                    Case 2
                        slTime = "  :15" 'Trim$(Str$((ilRow - 1) \ 4)) & ":15"
                    Case 3
                        slTime = "  :30" 'Trim$(Str$((ilRow - 1) \ 4)) & ":30"
                    Case Else
                        slTime = "  :45" 'Trim$(Str$((ilRow - 1) \ 4)) & ":45"
                End Select
            ElseIf (ilRow - 1) \ 4 = 12 Then
                Select Case ilRow Mod 4
                    Case 1
                        slTime = "12PM"
                    Case 2
                        slTime = "  :15" '"12:15"
                    Case 3
                        slTime = "  :30" '"12:30"
                    Case Else
                        slTime = "  :45" '"12:45"
                End Select
            Else
                Select Case ilRow Mod 4
                    Case 1
                        slTime = Trim$(str$((ilRow - 1) \ 4 - 12)) & "PM"
                    Case 2
                        slTime = "  :15" 'Trim$(Str$((ilRow - 1) \ 4 - 12)) & ":15"
                    Case 3
                        slTime = "  :30" 'Trim$(Str$((ilRow - 1) \ 4 - 12)) & ":30"
                    Case Else
                        slTime = "  :45" 'Trim$(Str$((ilRow - 1) \ 4 - 12)) & ":45"
                End Select
            End If
        Else
            If ilRow <= 4 Then
                slTime = "0"
                Select Case ilRow Mod 4
                    Case 1
                        slTime = slTime & ":00"
                    Case 2
                        slTime = " " & slTime & ":15" '"12:15"
                    Case 3
                        slTime = " " & slTime & ":30" '"12:30"
                    Case Else
                        slTime = " " & slTime & ":45" '"12:45"
                End Select
            ElseIf (ilRow - 1) \ 4 < 12 Then
                slTime = Trim$(str$((ilRow - 1) \ 4))
                Select Case ilRow Mod 4
                    Case 1
                        slTime = slTime & ":00"
                    Case 2
                        slTime = " " & slTime & ":15" 'Trim$(Str$((ilRow - 1) \ 4)) & ":15"
                    Case 3
                        slTime = " " & slTime & ":30" 'Trim$(Str$((ilRow - 1) \ 4)) & ":30"
                    Case Else
                        slTime = " " & slTime & ":45" 'Trim$(Str$((ilRow - 1) \ 4)) & ":45"
                End Select
            ElseIf (ilRow - 1) \ 4 = 12 Then
                slTime = "12"
                Select Case ilRow Mod 4
                    Case 1
                        slTime = slTime & ":00"
                    Case 2
                        slTime = " " & slTime & ":15" '"12:15"
                    Case 3
                        slTime = " " & slTime & ":30" '"12:30"
                    Case Else
                        slTime = " " & slTime & ":45" '"12:45"
                End Select
            Else
                slTime = Trim$(str$((ilRow - 1) \ 4 - 12))
                Select Case ilRow Mod 4
                    Case 1
                        slTime = slTime & ":00"
                    Case 2
                        slTime = " " & slTime & ":15" 'Trim$(Str$((ilRow - 1) \ 4 - 12)) & ":15"
                    Case 3
                        slTime = " " & slTime & ":30" 'Trim$(Str$((ilRow - 1) \ 4 - 12)) & ":30"
                    Case Else
                        slTime = " " & slTime & ":45" 'Trim$(Str$((ilRow - 1) \ 4 - 12)) & ":45"
                End Select
            End If
        End If
        pbcCtrl.CurrentX = flX
        pbcCtrl.CurrentY = flY
'        gPaintArea pbcCtrl, flX, flY, 375 - 15, 165 - 15, LIGHTYELLOW
        pbcCtrl.CurrentX = flX + 360 - pbcCtrl.TextWidth(Trim$(slTime))
        pbcCtrl.CurrentY = flY
        pbcCtrl.Print Trim$(slTime)
        flY = flY + imHeightWithinColumn + 15  'fgBoxGridH
    Next ilRow
    pbcCtrl.ForeColor = llColor
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadLcfLnfLcf                  *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read in all events for a date  *
'*                                                     *
'*******************************************************
Private Sub mReadLcfLnfLef(ilType As Integer, ilDay As Integer, sLCP As String, ilDate0 As Integer, ilDate1 As Integer, tlLLC() As LLC, tlDLLC() As LLC)
    Dim ilUpper As Integer
    Dim ilDUpper As Integer
    Dim ilSeqNo As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim slStartTime As String
    Dim slStr As String
    Dim ilDel As Integer
    Dim ilDeleted As Integer
    Dim ilTestDel As Integer
    Dim slXMid As String
    ilUpper = UBound(tlLLC)
    ilDUpper = UBound(tlDLLC)
    'If igViewType = 1 Then
    '    slType = "A"
    'Else
    '    slType = "O"
    'End If
    ilTestDel = True
    tmLcfSrchKey.iType = ilType
    tmLcfSrchKey.sStatus = "D"
    tmLcfSrchKey.iVefCode = imVefCode
    tmLcfSrchKey.iLogDate(0) = ilDate0
    tmLcfSrchKey.iLogDate(1) = ilDate1
    tmLcfSrchKey.iSeqNo = 1
    ilRet = btrGetEqual(hmLcf, tmDLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
    If ilRet <> BTRV_ERR_NONE Then
        ilTestDel = False
    End If
    ilSeqNo = 1
    Do
        tmLcfSrchKey.iType = ilType
        tmLcfSrchKey.sStatus = sLCP
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = ilDate0
        tmLcfSrchKey.iLogDate(1) = ilDate1
        tmLcfSrchKey.iSeqNo = ilSeqNo
'        ilRet = btrGetEqual(hmLcf, tmCLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
        ilRet = btrGetEqual(hmLcf, tmCLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet = BTRV_ERR_NONE Then
            ilSeqNo = ilSeqNo + 1
            For ilIndex = LBound(tmCLcf.lLvfCode) To UBound(tmCLcf.lLvfCode) Step 1
                If tmCLcf.lLvfCode(ilIndex) <> 0 Then
                    'Test if deleted- if so ignore
                    ilDeleted = False
                    If ilTestDel Then
                        For ilDel = LBound(tmDLcf.lLvfCode) To UBound(tmDLcf.lLvfCode) Step 1
                            If tmDLcf.lLvfCode(ilDel) <> 0 Then
                                If tmCLcf.lLvfCode(ilIndex) = tmDLcf.lLvfCode(ilDel) Then
                                    'Test time
                                    If (tmCLcf.iTime(0, ilIndex) = tmDLcf.iTime(0, ilDel)) And (tmCLcf.iTime(1, ilIndex) = tmDLcf.iTime(1, ilDel)) Then
                                        ilDeleted = True
                                    End If
                                End If
                            End If
                        Next ilDel
                    End If
                    tlLLC(ilUpper).iDay = ilDay
                    gUnpackTime tmCLcf.iTime(0, ilIndex), tmCLcf.iTime(1, ilIndex), "A", "1", tlLLC(ilUpper).sStartTime
                    slStartTime = tlLLC(ilUpper).sStartTime
                    'Read in Ltf to obtain name and Lvf to obtain length
                    tmLvfSrchKey.lCode = tmCLcf.lLvfCode(ilIndex)
                    ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                    If ilRet = BTRV_ERR_NONE Then
                        tmLtfSrchKey.iCode = tmLvf.iLtfCode
                        ilRet = btrGetEqual(hmLtf, tmLtf, imLtfRecLen, tmLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                        If ilRet = BTRV_ERR_NONE Then
                            tlLLC(ilUpper).sType = tmLtf.sType
                            gUnpackLength tmLvf.iLen(0), tmLvf.iLen(1), "3", False, tlLLC(ilUpper).sLength
                            If tmLtf.iVar <> 0 Then
                                tlLLC(ilUpper).sName = Trim$(str$(tmLvf.iVersion)) & "/" & tmLtf.sName & "-" & Trim$(str$(tmLtf.iVar))
                            Else
                                tlLLC(ilUpper).sName = Trim$(str$(tmLvf.iVersion)) & "/" & tmLtf.sName
                            End If
                            tlLLC(ilUpper).lLvfCode = tmCLcf.lLvfCode(ilIndex)
                            ilUpper = ilUpper + 1
                            ReDim Preserve tlLLC(0 To ilUpper) As LLC
                            tlLLC(ilUpper).iDay = -1
                            If Not ilDeleted Then
                                mMoveLLC tlLLC(ilUpper - 1), tlDLLC(ilDUpper)
                                ilDUpper = ilDUpper + 1
                                ReDim Preserve tlDLLC(0 To ilDUpper) As LLC
                                tlDLLC(ilDUpper).iDay = -1
                            End If
                            'Read in all the event record (Lef)
                            tmLefSrchKey.lLvfCode = tmCLcf.lLvfCode(ilIndex)
                            tmLefSrchKey.iStartTime(0) = 0
                            tmLefSrchKey.iStartTime(1) = 0
                            tmLefSrchKey.iSeqNo = 0
                            ilRet = btrGetGreaterOrEqual(hmLef, tmLef, imLefRecLen, tmLefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            Do While (ilRet = BTRV_ERR_NONE) And (tmLef.lLvfCode = tmCLcf.lLvfCode(ilIndex))
                                tlLLC(ilUpper).iDay = ilDay
                                gUnpackLength tmLef.iStartTime(0), tmLef.iStartTime(1), "3", False, slStr
                                gAddTimeLength slStartTime, slStr, "A", "1", tlLLC(ilUpper).sStartTime, slXMid
                                Select Case tmLef.iEtfCode
                                    Case 1  'Program
                                        tlLLC(ilUpper).sType = "1"
                                        gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", False, tlLLC(ilUpper).sLength
                                        tmEnfSrchKey.iCode = tmLef.iEnfCode
                                        ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet = BTRV_ERR_NONE Then
                                            tlLLC(ilUpper).sName = Trim$(tmEnf.sName)
                                        Else
                                            tlLLC(ilUpper).sName = ""
                                        End If
                                    Case 2  'Contract Avail
                                        tlLLC(ilUpper).sType = "2"
                                        tlLLC(ilUpper).iUnits = tmLef.iMaxUnits
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        Else
                                            tlLLC(ilUpper).sLength = "0S"
                                        End If
                                    Case 3
                                        tlLLC(ilUpper).sType = "3"
                                    Case 4
                                        tlLLC(ilUpper).sType = "4"
                                    Case 5
                                        tlLLC(ilUpper).sType = "5"
                                    Case 6  'Cmml Promo
                                        tlLLC(ilUpper).sType = "6"
                                        tlLLC(ilUpper).iUnits = tmLef.iMaxUnits
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        Else
                                            tlLLC(ilUpper).sLength = "0S"
                                        End If
                                    Case 7  'Feed avail
                                        tlLLC(ilUpper).sType = "7"
                                        tlLLC(ilUpper).iUnits = tmLef.iMaxUnits
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        Else
                                            tlLLC(ilUpper).sLength = "0S"
                                        End If
                                    Case 8  'PSA/Promo (Avail)
                                        tlLLC(ilUpper).sType = "8"
                                        tlLLC(ilUpper).iUnits = tmLef.iMaxUnits
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        Else
                                            tlLLC(ilUpper).sLength = "0S"
                                        End If
                                    Case 9
                                        tlLLC(ilUpper).sType = "9"
                                        tlLLC(ilUpper).iUnits = tmLef.iMaxUnits
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        Else
                                            tlLLC(ilUpper).sLength = "0S"
                                        End If
                                    Case 10  'Page eject, Line space 1, 2 or 3
                                        tlLLC(ilUpper).sType = "A"
                                    Case 11
                                        tlLLC(ilUpper).sType = "B"
                                    Case 12
                                        tlLLC(ilUpper).sType = "C"
                                    Case 13
                                        tlLLC(ilUpper).sType = "D"
                                    Case Else   'Other
                                        tlLLC(ilUpper).sType = "Y"
                                        gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                        tmEnfSrchKey.iCode = tmLef.iEnfCode
                                        ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet = BTRV_ERR_NONE Then
                                            tlLLC(ilUpper).sName = Trim$(tmEnf.sName)
                                        Else
                                            tlLLC(ilUpper).sName = ""
                                        End If
                                End Select
                                ilUpper = ilUpper + 1
                                ReDim Preserve tlLLC(0 To ilUpper) As LLC
                                tlLLC(ilUpper).iDay = -1
                                If Not ilDeleted Then
                                    mMoveLLC tlLLC(ilUpper - 1), tlDLLC(ilDUpper)
                                    ilDUpper = ilDUpper + 1
                                    ReDim Preserve tlDLLC(0 To ilDUpper) As LLC
                                    tlDLLC(ilDUpper).iDay = -1
                                End If
                                ilRet = btrGetNext(hmLef, tmLef, imLefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        Else
                            tlLLC(ilUpper).iDay = -1
                        End If
                    Else
                        tlLLC(ilUpper).iDay = -1
                    End If
                Else
                    ilSeqNo = -1
                    Exit For
                End If
            Next ilIndex
        Else
            ilSeqNo = -1
        End If
    Loop While ilSeqNo > 0
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetWinStatus                   *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Use igWinStatus as a way to    *
'*                      limit changes if pending       *
'*                      linkage exist without pending  *
'*                      program libraries              *
'*                                                     *
'*******************************************************
Private Sub mSetWinStatus()
    Dim ilRet As Integer
    igWinStatus(PROGRAMMINGJOB) = imSvWinStatus
    If imSvWinStatus < 2 Then
        pbcCount.Enabled = False
        pbcLayout.Enabled = False
        pbcLibrary.Enabled = False
        Exit Sub
    End If
    'If igViewType = 1 Then
    '    tmLcfSrchKey.sType = "A"
    'Else
    '    tmLcfSrchKey.sType = "O"
    'End If
    tmLcfSrchKey.iType = 0
    tmLcfSrchKey.sStatus = "P"
    tmLcfSrchKey.iVefCode = imVefCode
    tmLcfSrchKey.iLogDate(0) = 0
    tmLcfSrchKey.iLogDate(1) = 0
    tmLcfSrchKey.iSeqNo = 1
    ilRet = btrGetGreaterOrEqual(hmLcf, tmPLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
    If ilRet = BTRV_ERR_NONE Then
        If (tmPLcf.sStatus = "P") And (tmPLcf.iVefCode = imVefCode) Then
            pbcCount.Enabled = True
            pbcLayout.Enabled = True
            pbcLibrary.Enabled = True
            Exit Sub
        End If
    End If
    'If igViewType = 1 Then
    '    tmLcfSrchKey.sType = "A"
    'Else
    '    tmLcfSrchKey.sType = "O"
    'End If
    tmLcfSrchKey.iType = 0
    tmLcfSrchKey.sStatus = "D"
    tmLcfSrchKey.iVefCode = imVefCode
    tmLcfSrchKey.iLogDate(0) = 0
    tmLcfSrchKey.iLogDate(1) = 0
    tmLcfSrchKey.iSeqNo = 1
    ilRet = btrGetGreaterOrEqual(hmLcf, tmPLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
    If ilRet = BTRV_ERR_NONE Then
        If (tmPLcf.sStatus = "D") And (tmPLcf.iVefCode = imVefCode) Then
            pbcCount.Enabled = True
            pbcLayout.Enabled = True
            pbcLibrary.Enabled = True
            Exit Sub
        End If
    End If
    'Test vehicle type
    If tmVef.sType = "S" Then
        ilRet = gCodeChrRefExist(Program, "Vlf.Btr", imVefCode, "VLFSELLCODE", "P", "VLFSTATUS")
    ElseIf tmVef.sType = "A" Then
        ilRet = gCodeChrRefExist(Program, "Vlf.Btr", imVefCode, "VLFAIRCODE", "P", "VLFSTATUS")
    Else
        ilRet = False
    End If
    If ilRet Then
        'igWinStatus(PROGRAMMINGJOB) = 1 'View only
        pbcCount.Enabled = False
        pbcLayout.Enabled = False
        pbcLibrary.Enabled = False
    Else
        pbcCount.Enabled = True
        pbcLayout.Enabled = True
        pbcLibrary.Enabled = True
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mShowlinInfo                    *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show library information for   *
'*                      right mouse                    *
'*                                                     *
'*******************************************************
Private Sub mShowLibInfo()
    Dim slLibName As String
    Dim slStartTime As String
    Dim slStatus As String
    Dim slLength As String
    Dim slVer As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilButtonIndex As Integer
    ilButtonIndex = imButtonIndex
    If (imButtonIndex < LBONE) Or (imButtonIndex > UBound(tmLCD)) Then
        plcLibInfo.Visible = False
        Exit Sub
    End If
    slStr = Trim$(tmLCD(ilButtonIndex).sName)
    If plcDate.Caption = "TFN" Then
        If imCurrent = 0 Then
            If tmLCD(ilButtonIndex).iCurOrPend = 0 Then
                slStatus = "Current"
                slStartTime = Trim$(tmCTFN(tmLCD(ilButtonIndex).iLLCIndex).sStartTime)
                slLength = Trim$(tmCTFN(tmLCD(ilButtonIndex).iLLCIndex).sLength)
                'tmLvfSrchKey.lCode = tmCLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            Else
                slStatus = "Pending"
                slStartTime = Trim$(tmPTFN(tmLCD(ilButtonIndex).iLLCIndex).sStartTime)
                slLength = Trim$(tmPTFN(tmLCD(ilButtonIndex).iLLCIndex).sLength)
                'tmLvfSrchKey.lCode = tmPLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            End If
        Else
            If tmLCD(ilButtonIndex).iCurOrPend = 0 Then
                slStatus = "Current"
                slStartTime = Trim$(tmCDTFN(tmLCD(ilButtonIndex).iLLCIndex).sStartTime)
                slLength = Trim$(tmCDTFN(tmLCD(ilButtonIndex).iLLCIndex).sLength)
                'tmLvfSrchKey.lCode = tmCLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            Else
                slStatus = "Pending"
                slStartTime = Trim$(tmPDTFN(tmLCD(ilButtonIndex).iLLCIndex).sStartTime)
                slLength = Trim$(tmPDTFN(tmLCD(ilButtonIndex).iLLCIndex).sLength)
                'tmLvfSrchKey.lCode = tmPLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            End If
        End If
    Else
        If imCurrent = 0 Then
            If tmLCD(ilButtonIndex).iCurOrPend = 0 Then
                slStatus = "Current"
                slStartTime = Trim$(tmCLLC(tmLCD(ilButtonIndex).iLLCIndex).sStartTime)
                slLength = Trim$(tmCLLC(tmLCD(ilButtonIndex).iLLCIndex).sLength)
                'tmLvfSrchKey.lCode = tmCLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            Else
                slStatus = "Pending"
                slStartTime = Trim$(tmPLLC(tmLCD(ilButtonIndex).iLLCIndex).sStartTime)
                slLength = Trim$(tmPLLC(tmLCD(ilButtonIndex).iLLCIndex).sLength)
                'tmLvfSrchKey.lCode = tmPLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            End If
        Else
            If tmLCD(ilButtonIndex).iCurOrPend = 0 Then
                slStatus = "Current"
                slStartTime = Trim$(tmCDLLC(tmLCD(ilButtonIndex).iLLCIndex).sStartTime)
                slLength = Trim$(tmCDLLC(tmLCD(ilButtonIndex).iLLCIndex).sLength)
                'tmLvfSrchKey.lCode = tmCLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            Else
                slStatus = "Pending"
                slStartTime = Trim$(tmPDLLC(tmLCD(ilButtonIndex).iLLCIndex).sStartTime)
                slLength = Trim$(tmPDLLC(tmLCD(ilButtonIndex).iLLCIndex).sLength)
                'tmLvfSrchKey.lCode = tmPLLC(tmLCD(imLCDDragIndex).iLLCIndex).lLvfCode
            End If
        End If
    End If
    ilPos = InStr(slStr, "/")
    If ilPos > 0 Then
        slLibName = Mid$(slStr, ilPos + 1)
        slVer = Left$(slStr, ilPos - 1)
    Else
        slLibName = slStr
        slVer = ""
    End If
    'ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
    lacLibInfo(0).Caption = "Library Name " & slLibName & "   Version # " & slVer & "   Status " & slStatus
    lacLibInfo(1).Caption = "Start Time " & slStartTime & "   Length " & slLength
    If (imButtonIndex < LBONE) Or (imButtonIndex > UBound(tmLCD)) Then
        plcLibInfo.Visible = False
        Exit Sub
    End If
    plcLibInfo.Visible = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mStartPEvent                    *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initiate PEvent                *
'*                                                     *
'*******************************************************
Private Function mStartPEvent() As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(PEVENTEXE)) Then
    '    mStartPEvent = False
    '    Exit Function
    'End If
    mStartPEvent = True
    'Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Program^Test\" & sgUserName
        Else
            slStr = "Program^Prod\" & sgUserName
        End If
    Else
        If igTestSystem Then
            slStr = "Program^Test^NOHELP\" & sgUserName
        Else
            slStr = "Program^Prod^NOHELP\" & sgUserName
        End If
    End If
    If igPrgDupl Then
        slStr = slStr & "\" & sgRemLibName & "\" & Trim$(str$(lgLibLength)) & "\"
        'If PrgDupl!lbcLib.ListIndex <= 0 Then
        '    slStr = slStr & "\\0\" 'This will generate a select_change event
        'Else
        '    slNameCode = PrgDupl!lbcLibName.List(PrgDupl!lbcLib.ListIndex - 1)
        '    ilRet = gParseItem(slNameCode, 1, "|", slName)
        '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
        '    slStr = slStr & "\" & slName & "\" & slCode & "\"
        'End If
    Else
        If lbcLib.ListIndex <= 0 Then
            slStr = slStr & "\\0\" 'This will generate a select_change event
        Else
            slNameCode = tmLibName(lbcLib.ListIndex - 1).sKey  'lbcLibName.List(lbcLib.ListIndex - 1)
            ilRet = gParseItem(slNameCode, 1, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            slStr = slStr & "\" & slName & "\" & slCode & "\"
        End If
    End If
    If igVehIndexViaPrg < 0 Then
        slStr = slStr & "\\" 'This will generate a select_change event
    Else
        slNameCode = tmPrgVehicle(igVehIndexViaPrg).sKey  'Traffic!lbcVehicle.List(igVehIndexViaPrg)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        slStr = slStr & slName & "\" & slCode & "\"
    End If
    'If pbcLibType.Enabled Then
        slStr = slStr & "Y\"
    'Else
    '    slStr = slStr & "N\"
    'End If
    If ckcShowVersion.Value = vbChecked Then
        slStr = slStr & "Y\"
    Else
        slStr = slStr & "N\"
    End If
    slStr = slStr & Trim$(str$(igLibType)) & "\" & Trim$(str$(igViewType)) & "\" & Trim$(str$(igPrgDupl))
    'lgShellRet = Shell(sgExePath & "PEvent.Exe " & slStr, 1)
    'Program.Enabled = False
    ''Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop

    sgCommandStr = slStr
    PEvent.Show vbModal

    slStr = sgDoneMsg
    'Program.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    mStartPEvent = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    igWinStatus(PROGRAMMINGJOB) = imSvWinStatus

    smPrgVehicleTag = ""
    imTerminate = False
    Screen.MousePointer = vbDefault
    'Unload IconTraf
    igManUnload = YES
    Unload Program
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehBranch                      *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to event  *
'*                      type and process               *
'*                      communication back from event  *
'*                      type                           *
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
Private Function mVehBranch() As Integer
'
'   ilRet = mVehBranch()
'   Where:
'       ilRet (O)- True = event type started
'                  False = Event type over or not started
'
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    gManLookAhead cbcVeh, imBSMode, imComboBoxIndex
    If cbcVeh.ListIndex > 0 Then
        mVehBranch = False
        Exit Function
    Else
        'If Not gWinRoom(igNoLJWinRes(VEHICLESLIST)) Then
        '    mVehBranch = True
        '    Exit Function
        'End If
        plcDate.Caption = ""
        ReDim tmLCD(0 To 1) As LCD
        ReDim tmCLLC(0 To 0) As LLC
        ReDim tmCDLLC(0 To 0) As LLC
        ReDim tmPLLC(0 To 0) As LLC
        ReDim tmPDLLC(0 To 0) As LLC
        tmCLLC(0).iDay = -1
        tmCDLLC(0).iDay = -1
        tmPLLC(0).iDay = -1
        tmPDLLC(0).iDay = -1
        ReDim tmCTFN(0 To 0) As LLC
        ReDim tmCDTFN(0 To 0) As LLC
        ReDim tmPTFN(0 To 0) As LLC
        ReDim tmPDTFN(0 To 0) As LLC
        tmCTFN(0).iDay = -1
        tmCDTFN(0).iDay = -1
        tmPTFN(0).iDay = -1
        tmPDTFN(0).iDay = -1
        smLibNameTag = ""
        lbcEvents.Clear
        lbcLib.Clear
        pbcCount.Cls
        pbcLayout.Cls
        pbcLibrary.Cls
        'Screen.MousePointer = vbHourGlass  'Wait
        igVehCallSource = CALLSOURCEPRG
        If cbcVeh.Text = "[New]" Then
            sgVehName = ""
        Else
            sgVehName = slStr
        End If
        ilUpdateAllowed = imUpdateAllowed
        'igChildDone = False
        'edcLinkSrceDoneMsg.Text = ""
        'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        '10071 removed
'            If igTestSystem Then
'                slStr = "Program^Test\" & sgUserName & "\" & Trim$(str$(igVehCallSource)) & "\" & sgVehName
'            Else
'                slStr = "Program^Prod\" & sgUserName & "\" & Trim$(str$(igVehCallSource)) & "\" & sgVehName
'            End If
            
            
        'Else
        '    If igTestSystem Then
        '        slStr = "Program^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igVehCallSource)) & "\" & sgVehName
        '    Else
        '        slStr = "Program^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igVehCallSource)) & "\" & sgVehName
        '    End If
        'End If
        'lgShellRet = Shell(sgExePath & "Vehicle.Exe " & slStr, 1)
        'Program.Enabled = False
        'Do While Not igChildDone
        '    DoEvents
        'Loop
        '10071
        'sgCommandStr = slStr
        'Vehicle.Show vbModal
        
        On Error Resume Next
        gCallVehicleProject Me

        '10071
'        slStr = sgDoneMsg
'        ilParse = gParseItem(slStr, 1, "\", sgVehName)
'        igVehCallSource = Val(sgVehName)
'        ilParse = gParseItem(slStr, 2, "\", sgVehName)
        
        
        'Program.Enabled = True
        'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
        'For ilLoop = 0 To 10
        '    DoEvents
        'Next ilLoop
        'Screen.MousePointer = vbDefault    'Default
        imUpdateAllowed = ilUpdateAllowed
'        gShowBranner
        'If imUpdateAllowed = False Then
        '    mSendHelpMsg "BF"
        'Else
        '    mSendHelpMsg "BT"
        'End If
        gShowBranner imUpdateAllowed
        mVehBranch = False
        '10071
        'If igVehCallSource = CALLDONE Then  'Done
            igVehCallSource = CALLNONE
            smPrgVehicleTag = ""
            cbcVeh.Clear
            mVehPop
            If imTerminate Then
                mVehBranch = False
                Exit Function
            End If
            'not looking for new.  Ok for now
            gFindMatch sgVehName, 1, cbcVeh
            sgVehName = ""
            If gLastFound(cbcVeh) > 0 Then
                imChgMode = True
                cbcVeh.ListIndex = gLastFound(cbcVeh)
                imComboBoxIndex = cbcVeh.ListIndex
                igVehIndexViaPrg = imComboBoxIndex - 1
                imChgMode = False
            Else
                imChgMode = True
                cbcVeh.ListIndex = 0
                imChgMode = False
                mVehBranch = True
                Exit Function
            End If
        End If
        '10071
'        If igVehCallSource = CALLCANCELLED Then  'Cancelled
'            igVehCallSource = CALLNONE
'            sgVehName = ""
'            mVehBranch = True
'            Exit Function
'        End If
'        If igVehCallSource = CALLTERMINATED Then
'            igVehCallSource = CALLNONE
'            sgVehName = ""
'            mVehBranch = True
'            Exit Function
'        End If
  '10071
  '  End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
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

    'ilRet = gPopUserVehicleBox(Program, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH, cbcVeh, Traffic!lbcVehicle), 1/25/21 Exclude CPM vehicles
    ilRet = gPopUserVehicleBox(Program, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + VEHSPORT + ACTIVEVEH, cbcVeh, tmPrgVehicle(), smPrgVehicleTag) '2/8/21 - Include Podcast vehicles
    If ilRet <> CP_MSG_NOPOPREQ Then
        cbcVeh.AddItem "[New]", 0  'Force as first item on list
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", Program
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mWkCnt                          *
'*                                                     *
'*             Created:10/11/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain counts within days      *
'*                                                     *
'*******************************************************
Private Sub mWkCnt(ilListIndex As Integer, ilStartDay As Integer, ilEndDay As Integer, clStartTime As Currency, clEndTime As Currency, tlLLC() As LLC)
'
'   mWkCnt clStartTime, clTimeInc, tmCLLC()
'   Where:
'       clStartTime (I)- Start time
'       clEndTime (I) - End time
'       tmCLLC() (I)- event records to be processed
'
    Dim ilCol As Integer
    Dim ilIndex As Integer
    Dim clGStartTime As Currency
    Dim clGEndTime As Currency
    Dim clEvtTime As Currency
    Dim ilFindEvt As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slEvtType As String
    Dim ilRet As Integer
    Exit Sub
    If ilListIndex < 0 Then
        Exit Sub
    End If
    slNameCode = tmETypeCode(ilListIndex).sKey 'lbcETypeCode.List(ilListIndex)
    ilRet = gParseItem(slNameCode, 3, "\", slCode)
    If ilRet <> CP_MSG_NONE Then
        Exit Sub
    End If
    Select Case Val(slCode)
        Case 1  'Program
            slEvtType = "1"
        Case 2  'Contract Avail
            slEvtType = "2"
        Case 3
            slEvtType = "3"
        Case 4
            slEvtType = "4"
        Case 5
            slEvtType = "5"
        Case 6  'Cmml Promo
            slEvtType = "6"
        Case 7  'Feed avail
            slEvtType = "7"
        Case 8  'PSA/Promo (Avail)
            slEvtType = "8"
        Case 9
            slEvtType = "9"
        Case 10  'Page eject, Line space 1, 2 or 3
            slEvtType = "A"
        Case 11
            slEvtType = "B"
        Case 12
            slEvtType = "C"
        Case 13
            slEvtType = "D"
        Case Else   'Other
            slEvtType = "Y"
    End Select
    ilIndex = LBound(tlLLC)
    clGStartTime = clStartTime
    clGEndTime = clEndTime
    For ilCol = ilStartDay To ilEndDay Step 1
        'Only required if not showing 24 hours
        If (tlLLC(ilIndex).iDay = -1) Or (ilIndex >= UBound(tlLLC)) Then
            Exit For
        End If
        Do While (tlLLC(ilIndex).iDay < ilCol)
            ilIndex = ilIndex + 1
            If (tlLLC(ilIndex).iDay = -1) Or (ilIndex >= UBound(tlLLC)) Then
                Exit For
            End If
        Loop
        If tlLLC(ilIndex).iDay = ilCol Then
            ilFindEvt = True
            Do
                clEvtTime = gTimeToCurrency(tlLLC(ilIndex).sStartTime, False)
                If tlLLC(ilIndex).iDay = ilCol Then
                    If (clEvtTime >= clGStartTime) And (clEvtTime <= clGEndTime) Then
                        If tlLLC(ilIndex).sType = slEvtType Then
                            Select Case slEvtType
                                Case "1"  'Program
                                    imNoEvt = imNoEvt + 1
                                Case "2", "3", "4", "5"  'Contract Avail
                                    imUnits = imUnits + tlLLC(ilIndex).iUnits
                                    If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                        cmSec = cmSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                    ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                        cmSec = cmSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                    End If
'                                Case "3"    'Open BB
'                                    imNoEvt = imNoEvt + 1
'                                Case "4"    'Floating BB
'                                    imNoEvt = imNoEvt + 1
'                                Case "5"    'Close BB
'                                    imNoEvt = imNoEvt + 1
                                Case "6"  'Cmml Promo
                                    imUnits = imUnits + tlLLC(ilIndex).iUnits
                                    If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                        cmSec = cmSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                    ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                        cmSec = cmSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                    End If
                                Case "7"  'Feed avail
                                    imUnits = imUnits + tlLLC(ilIndex).iUnits
                                    If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                        cmSec = cmSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                    ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                        cmSec = cmSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                    End If
                                Case "8"  'PSA/Promo (Avail)
                                    imUnits = imUnits + tlLLC(ilIndex).iUnits
                                    If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                        cmSec = cmSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                    ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                        cmSec = cmSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                    End If
                                Case "9"
                                    imUnits = imUnits + tlLLC(ilIndex).iUnits
                                    If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                        cmSec = cmSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                    ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                        cmSec = cmSec + gLengthToCurrency(tlLLC(ilIndex).sLength)
                                    End If
                                Case "A"  'Page eject, Line space 1, 2 or 3
                                    imNoEvt = imNoEvt + 1
                                Case "B"
                                    imNoEvt = imNoEvt + 1
                                Case "C"
                                    imNoEvt = imNoEvt + 1
                                Case "D"
                                    imNoEvt = imNoEvt + 1
                                Case "Y"   'Other
                                    imNoEvt = imNoEvt + 1
                            End Select
                        End If
                        ilIndex = ilIndex + 1
                        If ilIndex >= UBound(tlLLC) Then
                            ilFindEvt = False
                        End If
                    Else
                        If clEvtTime > clGEndTime Then
                            ilFindEvt = False
                        Else
                            ilIndex = ilIndex + 1
                            If ilIndex >= UBound(tlLLC) Then
                                ilFindEvt = False
                            End If
                        End If
                    End If
                Else
                    ilFindEvt = False
                End If
            Loop While ilFindEvt
            If ilIndex >= UBound(tlLLC) Then
                Exit For
            End If
        End If
    Next ilCol
End Sub
Private Sub mWkCntForLibrary(ilListIndex As Integer, clStartTime As Currency, clEndTime As Currency, tlCLLC() As LLC, tlPLLC() As LLC)
'
'   mWkCntForLibrary clStartTime, clTimeInc, tmCLLC(), tmPLLC()
'   Where:
'       clStartTime (I)- Start time
'       clTimeInc (I) - Time increment (3600)
'       tmCLLC() (I)- event records to be processed
'
    Dim ilCol As Integer
    Dim ilCIndex As Integer
    Dim ilPIndex
    Dim clGStartTime As Currency
    Dim clGEndTime As Currency
    Dim clCEvtTime As Currency
    Dim clCEvtEndTime As Currency
    Dim clPEvtTime As Currency
    Dim clPEvtEndTime As Currency
    Dim slTime As String
    Dim ilCFindEvt As Integer
    Dim ilPFindEvt As Integer
    Dim slXMid As String
    Exit Sub

    ilCIndex = LBound(tlCLLC)
    ilPIndex = LBound(tlPLLC)
    For ilCol = 1 To 7 Step 1
        clGStartTime = clStartTime
        clGEndTime = clEndTime
        Do While (tlCLLC(ilCIndex).iDay = ilCol) Or ((tlPLLC(ilPIndex).iDay = ilCol) And (imCurrent <> 0))
            ilCFindEvt = False
            Do While tlCLLC(ilCIndex).iDay = ilCol
                If (tlCLLC(ilCIndex).sType = "R") Or (tlCLLC(ilCIndex).sType = "S") Or (tlCLLC(ilCIndex).sType = "P") Then
                    clCEvtTime = gTimeToCurrency(tlCLLC(ilCIndex).sStartTime, False)
                    gAddTimeLength tlCLLC(ilCIndex).sStartTime, tlCLLC(ilCIndex).sLength, "A", "1", slTime, slXMid
                    clCEvtEndTime = gTimeToCurrency(slTime, True) - 1
                    If (clCEvtEndTime < clGStartTime) Or (clCEvtTime > clGEndTime) Then
                        If (tlCLLC(ilCIndex).iDay = -1) Or (ilCIndex >= UBound(tlCLLC)) Then
                            Exit Do
                        End If
                        ilCIndex = ilCIndex + 1
                    Else
                        ilCFindEvt = True
                        Exit Do
                    End If
                Else
                    If (tlCLLC(ilCIndex).iDay = -1) Or (ilCIndex >= UBound(tlCLLC)) Then
                        Exit Do
                    End If
                    ilCIndex = ilCIndex + 1
                End If
            Loop
            ilPFindEvt = False
            If imCurrent <> 0 Then
                Do While tlPLLC(ilPIndex).iDay = ilCol
                    If (tlPLLC(ilPIndex).sType = "R") Or (tlPLLC(ilPIndex).sType = "S") Or (tlPLLC(ilPIndex).sType = "P") Then
                        clPEvtTime = gTimeToCurrency(tlPLLC(ilPIndex).sStartTime, False)
                        gAddTimeLength tlPLLC(ilPIndex).sStartTime, tlPLLC(ilPIndex).sLength, "A", "1", slTime, slXMid
                        clPEvtEndTime = gTimeToCurrency(slTime, True) - 1
                        If (clPEvtEndTime < clGStartTime) Or (clPEvtTime > clGEndTime) Then
                            If (tlPLLC(ilPIndex).iDay = -1) Or (ilPIndex >= UBound(tlPLLC)) Then
                                Exit Do
                            End If
                            ilPIndex = ilPIndex + 1
                        Else
                            ilPFindEvt = True
                            Exit Do
                        End If
                    Else
                        If (tlPLLC(ilPIndex).iDay = -1) Or (ilPIndex >= UBound(tlPLLC)) Then
                            Exit Do
                        End If
                        ilPIndex = ilPIndex + 1
                    End If
                Loop
            End If
            If ilCFindEvt And ilPFindEvt Then
                If clCEvtEndTime < clPEvtTime Then
                    mWkCnt ilListIndex, ilCol, ilCol, clCEvtTime, clCEvtEndTime, tlCLLC()
                    ilCIndex = ilCIndex + 1
                ElseIf clPEvtEndTime < clCEvtTime Then
                    mWkCnt ilListIndex, ilCol, ilCol, clPEvtTime, clPEvtEndTime, tlPLLC()
                    ilPIndex = ilPIndex + 1
                Else
                    ilCIndex = ilCIndex + 1
                End If
            ElseIf ilCFindEvt And Not ilPFindEvt Then
                mWkCnt ilListIndex, ilCol, ilCol, clCEvtTime, clCEvtEndTime, tlCLLC()
                ilCIndex = ilCIndex + 1
            ElseIf Not ilCFindEvt And ilPFindEvt Then
                mWkCnt ilListIndex, ilCol, ilCol, clPEvtTime, clPEvtEndTime, tlPLLC()
                ilPIndex = ilPIndex + 1
            End If
        Loop
    Next ilCol
End Sub
Private Sub mWkCntPaint(ilListIndex As Integer)
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slEvtType As String
    Dim slStr As String
    plcWkCnt.Caption = ""
    Exit Sub
    If ilListIndex < 0 Then
        Exit Sub
    End If
    slNameCode = tmETypeCode(ilListIndex).sKey 'lbcETypeCode.List(ilListIndex)
    ilRet = gParseItem(slNameCode, 3, "\", slCode)
    If ilRet <> CP_MSG_NONE Then
        Exit Sub
    End If
    Select Case Val(slCode)
        Case 1  'Program
            slEvtType = "1"
        Case 2  'Contract Avail
            slEvtType = "2"
        Case 3
            slEvtType = "3"
        Case 4
            slEvtType = "4"
        Case 5
            slEvtType = "5"
        Case 6  'Cmml Promo
            slEvtType = "6"
        Case 7  'Feed avail
            slEvtType = "7"
        Case 8  'PSA/Promo (Avail)
            slEvtType = "8"
        Case 9
            slEvtType = "9"
        Case 10  'Page eject, Line space 1, 2 or 3
            slEvtType = "A"
        Case 11
            slEvtType = "B"
        Case 12
            slEvtType = "C"
        Case 13
            slEvtType = "D"
        Case Else   'Other
            slEvtType = "Y"
    End Select
    Select Case slEvtType
        Case "1"  'Program
            If imNoEvt > 0 Then
                plcWkCnt.Caption = Trim$(str$(imNoEvt))
            End If
        Case "2", "3", "4", "5"  'Contract Avail
            If imUnits > 0 Then
                slStr = Trim$(str$(imUnits))
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    slStr = slStr & "/" & gCurrencyToLength(cmSec)
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    slStr = slStr & "/" & gCurrencyToLength(cmSec)
                End If
                plcWkCnt.Caption = slStr
            End If
'        Case "3"    'Open BB
'            If imNoEvt > 0 Then
'                slStr = Trim$(Str$(imNoEvt))
'                plcWkCnt.Caption = slStr
'            End If
'        Case "4"    'Floating BB
'            If imNoEvt > 0 Then
'                slStr = Trim$(Str$(imNoEvt))
'                plcWkCnt.Caption = slStr
'            End If
'        Case "5"    'Close BB
'            If imNoEvt > 0 Then
'                slStr = Trim$(Str$(imNoEvt))
'                plcWkCnt.Caption = slStr
'            End If
        Case "6"  'Cmml Promo
            If imUnits > 0 Then
                slStr = Trim$(str$(imUnits))
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    slStr = slStr & "/" & gCurrencyToLength(cmSec)
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    slStr = slStr & "/" & gCurrencyToLength(cmSec)
                End If
                plcWkCnt.Caption = slStr
            End If
        Case "7"  'Feed avail
            If imUnits > 0 Then
                slStr = Trim$(str$(imUnits))
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    slStr = slStr & "/" & gCurrencyToLength(cmSec)
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    slStr = slStr & "/" & gCurrencyToLength(cmSec)
                End If
                plcWkCnt.Caption = slStr
            End If
        Case "8"  'PSA/Promo (Avail)
            If imUnits > 0 Then
                slStr = Trim$(str$(imUnits))
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    slStr = slStr & "/" & gCurrencyToLength(cmSec)
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    slStr = slStr & "/" & gCurrencyToLength(cmSec)
                End If
                plcWkCnt.Caption = slStr
            End If
        Case "9"
            If imUnits > 0 Then
                slStr = Trim$(str$(imUnits))
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    slStr = slStr & "/" & gCurrencyToLength(cmSec)
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    slStr = slStr & "/" & gCurrencyToLength(cmSec)
                End If
                plcWkCnt.Caption = slStr
            End If
        Case "A"  'Page eject, Line space 1, 2 or 3
            If imNoEvt > 0 Then
                slStr = Trim$(str$(imNoEvt))
                plcWkCnt.Caption = slStr
            End If
        Case "B"
            If imNoEvt > 0 Then
                slStr = Trim$(str$(imNoEvt))
                plcWkCnt.Caption = slStr
            End If
        Case "C"
            If imNoEvt > 0 Then
                slStr = Trim$(str$(imNoEvt))
                pbcCount.Print slStr
            End If
        Case "D"
            If imNoEvt > 0 Then
                slStr = Trim$(str$(imNoEvt))
                plcWkCnt.Caption = slStr
            End If
        Case "Y"   'Other
            If imNoEvt > 0 Then
                slStr = Trim$(str$(imNoEvt))
                plcWkCnt.Caption = slStr
            End If
    End Select
End Sub
Private Sub pbcClickFocus_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcClickFocus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcCount_Paint()
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim clStart As Currency
    Dim clInc As Currency
    Dim clCntStart As Currency
    Dim clCntEnd As Currency
    Dim ilListIndex As Integer
    mPaintColumns
    If imResol = 0 Then
        'Paint times
        mPaintTimes 1, 96, 4, pbcCount
        'Print Dates
        mPaintDates pbcCount
        clStart = 0
        clInc = 3600
    ElseIf imResol = 1 Then 'Half hour
        ilStart = 2 * (vbcLayout.Value - 1) + 1
        ilEnd = 2 * (vbcLayout.Value + vbcLayout.LargeChange)
        mPaintTimes ilStart, ilEnd, 2, pbcCount
        'Print Dates
        mPaintDates pbcCount
        clStart = CLng(1800) * (vbcLayout.Value - 1)
        clInc = 1800
    Else    'Quarter hour
        ilStart = vbcLayout.Value
        ilEnd = vbcLayout.Value + vbcLayout.LargeChange
        mPaintTimes ilStart, ilEnd, 1, pbcCount
        'Print Dates
        mPaintDates pbcCount
        clStart = CLng(900) * (vbcLayout.Value - 1)
        clInc = 900
    End If
    imNoEvt = 0
    imUnits = 0
    cmSec = 0
    clCntStart = 0
    clCntEnd = 86399
    ilListIndex = lbcLib.ListIndex
    If imCurrent = 0 Then
        If plcDate.Caption <> "TFN" Then
            mPaintCounts clStart, clInc, tmCLLC()
            mWkCnt ilListIndex, 1, 7, clCntStart, clCntEnd, tmCLLC()
        Else
            mPaintCounts clStart, clInc, tmCTFN()
            mWkCnt ilListIndex, 1, 7, clCntStart, clCntEnd, tmCTFN()
        End If
    Else
        If plcDate.Caption <> "TFN" Then
            mPaintCounts clStart, clInc, tmPDLLC()
            mWkCnt ilListIndex, 1, 7, clCntStart, clCntEnd, tmPDLLC()
        Else
            mPaintCounts clStart, clInc, tmPDTFN()
            mWkCnt ilListIndex, 1, 7, clCntStart, clCntEnd, tmPDTFN()
        End If
    End If
    mWkCntPaint ilListIndex
End Sub
Private Sub pbcCurrent_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub pbcCurrent_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcCurrent_KeyPress(KeyAscii As Integer)
    Dim ilIndex As Integer
    Dim slDate As String
    If KeyAscii = Asc(" ") Then
        If imCurrent = 0 Then
            imCurrent = 1
        Else
            imCurrent = 0
        End If
    Else
        Exit Sub
    End If
    'Pick same date of possible
    pbcCurrent_Paint
    vbcLayout.Value = vbcLayout.Min
    imCurrentPending = True
    If imCurrent = 0 Then   'Current
        hbcDate.Max = (lmCLatestDate - lmCEarliestDate) \ 7 + 2 '1 for adjusting and 1 for TFN
    Else    'Pending
        hbcDate.Max = (lmPLatestDate - lmPEarliestDate) \ 7 + 2
    End If
    imCurrentPending = False
    If plcDate.Caption = "" Then
        If rbcType(2).Value Then
            pbcCount.Cls
        ElseIf rbcType(1).Value Then
            pbcLayout.Cls
        Else
            pbcLibrary.Cls
        End If
        Exit Sub
    End If
    If imCurrent = 0 Then
        'Was pending
        If plcDate.Caption <> "TFN" Then
            slDate = plcDate.Caption
            If gDateValue(slDate) >= lmCEarliestDate Then
                ilIndex = (gDateValue(slDate) - lmCEarliestDate) \ 7 + 1
                If ilIndex <= hbcDate.Max Then
                    If hbcDate.Value = ilIndex Then
                        hbcDate_Change
                    Else
                        hbcDate.Value = ilIndex
                    End If
                Else
                    If hbcDate.Value = 1 Then
                        hbcDate_Change
                    Else
                        hbcDate.Value = 1
                    End If
                End If
            Else
                If hbcDate.Value = 1 Then
                    hbcDate_Change
                Else
                    hbcDate.Value = 1
                End If
            End If
        End If
    Else
        'Was current
        If plcDate.Caption <> "TFN" Then
            slDate = plcDate.Caption
            If gDateValue(slDate) >= lmPEarliestDate Then
                ilIndex = (gDateValue(slDate) - lmPEarliestDate) \ 7 + 1
                If ilIndex <= hbcDate.Max Then
                    If hbcDate.Value = ilIndex Then
                        hbcDate_Change
                    Else
                        hbcDate.Value = ilIndex
                    End If
                Else
                    If hbcDate.Value = 1 Then
                        hbcDate_Change
                    Else
                        hbcDate.Value = 1
                    End If
                End If
            Else
                If hbcDate.Value = 1 Then
                    hbcDate_Change
                Else
                    hbcDate.Value = 1
                End If
            End If
        End If
    End If
    'MousePointer = vbHourGlass
    'If rbcType(2).Value Then
    '    pbcCount.Cls
    '    pbcCount_Paint
    'ElseIf rbcType(1).Value Then
    '    pbcLayout.Cls
    '    pbcLayout_Paint
    'Else
    '    pbcLibrary.Cls
    '    pbcLibrary_Paint
    'End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub pbcCurrent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    Dim slDate As String
    If imCurrent = 0 Then
        imCurrent = 1
    Else
        imCurrent = 0
    End If
    pbcCurrent_Paint
    vbcLayout.Value = vbcLayout.Min
    'Pick same date of possible
    imCurrentPending = True
    If imCurrent = 0 Then   'Current
        hbcDate.Max = (lmCLatestDate - lmCEarliestDate) \ 7 + 2 '1 for adjusting and 1 for TFN
    Else    'Pending
        hbcDate.Max = (lmPLatestDate - lmPEarliestDate) \ 7 + 2
    End If
    imCurrentPending = False
    If plcDate.Caption = "" Then
        If rbcType(2).Value Then
            pbcCount.Cls
        ElseIf rbcType(1).Value Then
            pbcLayout.Cls
        Else
            pbcLibrary.Cls
        End If
        Exit Sub
    End If
    If imCurrent = 0 Then
        'Was pending
        If plcDate.Caption <> "TFN" Then
            slDate = plcDate.Caption
            If gDateValue(slDate) >= lmCEarliestDate Then
                ilIndex = (gDateValue(slDate) - lmCEarliestDate) \ 7 + 1
                If ilIndex <= hbcDate.Max Then
                    If hbcDate.Value = ilIndex Then
                        hbcDate_Change
                    Else
                        hbcDate.Value = ilIndex
                    End If
                Else
                    If hbcDate.Value = 1 Then
                        hbcDate_Change
                    Else
                        hbcDate.Value = 1
                    End If
                End If
            Else
                If hbcDate.Value = 1 Then
                    hbcDate_Change
                Else
                    hbcDate.Value = 1
                End If
            End If
        End If
    Else
        'Was current
        If plcDate.Caption <> "TFN" Then
            slDate = plcDate.Caption
            If gDateValue(slDate) >= lmPEarliestDate Then
                ilIndex = (gDateValue(slDate) - lmPEarliestDate) \ 7 + 1
                If ilIndex <= hbcDate.Max Then
                    If hbcDate.Value = ilIndex Then
                        hbcDate_Change
                    Else
                        hbcDate.Value = ilIndex
                    End If
                Else
                    If hbcDate.Value = 1 Then
                        hbcDate_Change
                    Else
                        hbcDate.Value = 1
                    End If
                End If
            Else
                If hbcDate.Value = 1 Then
                    hbcDate_Change
                Else
                    hbcDate.Value = 1
                End If
            End If
        End If
    End If
    'MousePointer = vbHourGlass
    'If rbcType(2).Value Then
    '    pbcCount.Cls
    '    pbcCount_Paint
    'ElseIf rbcType(1).Value Then
    '    pbcLayout.Cls
    '    pbcLayout_Paint
    'Else
    '    pbcLibrary.Cls
    '    pbcLibrary_Paint
    'End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub pbcCurrent_Paint()
    pbcCurrent.Cls
    pbcCurrent.CurrentX = fgBoxInsetX
    pbcCurrent.CurrentY = -15 'fgBoxInsetY
    If imCurrent = 0 Then
        pbcCurrent.Print "Current"
        imcTrash.Visible = False
        'imcHelp.Visible = True
    ElseIf imCurrent = 1 Then
        pbcCurrent.Print "Pending"
        'Only show trash when dragging- if reinsert trash when pending- then fix
        'code to not remove trash after drop if pending is shown
        'If imUsingTFNForDate Then
            imcTrash.Visible = False
            'imcHelp.Visible = True
        'Else
        '    imcTrash.Visible = True
        'End If
    End If
End Sub
Private Sub pbcLayout_Paint()
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim clStart As Currency
    Dim clInc As Currency
    Dim clCntStart As Currency
    Dim clCntEnd As Currency
    Dim ilListIndex As Integer
    'Paint times
'    If ((rbcType(0).Value <> False) And (imResol = 0)) Or (rbcType(0).Value = False) Then
    mPaintColumns
    If imResol = 0 Then
        mPaintTimes 1, 96, 4, pbcLayout
        'Print Dates
        mPaintDates pbcLayout
        clStart = 0
        clInc = 3600
    ElseIf imResol = 1 Then 'Half hour
        ilStart = 2 * (vbcLayout.Value - 1) + 1
        ilEnd = 2 * (vbcLayout.Value + vbcLayout.LargeChange)
        mPaintTimes ilStart, ilEnd, 2, pbcLayout
        'Print Dates
        mPaintDates pbcLayout
        clStart = CLng(1800) * (vbcLayout.Value - 1)
        clInc = 1800
    Else    'Quarter hour
        ilStart = vbcLayout.Value
        ilEnd = vbcLayout.Value + vbcLayout.LargeChange
        mPaintTimes ilStart, ilEnd, 1, pbcLayout
        'Print Dates
        mPaintDates pbcLayout
        clStart = CLng(900) * (vbcLayout.Value - 1)
        clInc = 900
    End If
    imNoEvt = 0
    imUnits = 0
    cmSec = 0
    clCntStart = 0
    clCntEnd = 86399
    ilListIndex = lbcLib.ListIndex
    If imCurrent = 0 Then
        If plcDate.Caption <> "TFN" Then
            mPaintLayout clStart, clInc, tmCLLC()
            mWkCnt ilListIndex, 1, 7, clCntStart, clCntEnd, tmCLLC()
        Else
            mPaintLayout clStart, clInc, tmCTFN()
            mWkCnt ilListIndex, 1, 7, clCntStart, clCntEnd, tmCTFN()
        End If
    Else
        If plcDate.Caption <> "TFN" Then
            mPaintLayout clStart, clInc, tmPDLLC()
            mWkCnt ilListIndex, 1, 7, clCntStart, clCntEnd, tmPDLLC()
        Else
            mPaintLayout clStart, clInc, tmPDTFN()
            mWkCnt ilListIndex, 1, 7, clCntStart, clCntEnd, tmPDTFN()
        End If
    End If
    mWkCntPaint ilListIndex
End Sub
Private Sub pbcLibrary_DragDrop(Source As control, X As Single, Y As Single)
    Dim llStart As Long
    Dim llTotalTime As Long
    Dim llTime As Long
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slTime As String
    Dim ilCol As Integer
    Dim flX1 As Single
    Dim flX2 As Single
    Dim slDate As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilDay As Integer
    'Dim slType As String
    Dim ilType As Integer
    Dim ilStartDate As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String

    plcDragTime.Visible = False
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
    If imDragSource = 0 Then
        'If Not gWinRoom(igNoExeWinRes(PRGDATESEXE)) Then
        '    Exit Sub
        'End If
        'Obtain time
        If imResol = 0 Then
            llStart = 0
            llTotalTime = 86400
        ElseIf imResol = 1 Then 'Half hour
            llStart = CLng(1800) * (vbcLayout.Value - 1)
            llTotalTime = 43200
        Else    'Quarter hour
            llStart = CLng(900) * (vbcLayout.Value - 1)
            llTotalTime = 21600
        End If
        llTime = llStart + (llTotalTime * (Y - 285)) / fmPaintHeight
        llTime = 900 * ((llTime + 450) \ 900)
        If (llTime >= 0) And (X >= imStartXFirstColumn) And (llTime + lmPrgLength <= 86400) Then
            gPackTimeLong llTime, ilHsSec, ilMinHr
            gUnpackTime ilHsSec, ilMinHr, "A", "2", slTime
        ElseIf (llTime + lmPrgLength > 86400) Then
            llTime = 86400 - lmPrgLength
            gPackTimeLong llTime, ilHsSec, ilMinHr
            gUnpackTime ilHsSec, ilMinHr, "A", "2", slTime
        Else
            Exit Sub
        End If
        'Obtain date
        flX1 = imStartXFirstColumn
        flX2 = flX1 + imWidthWithinColumn
        For ilCol = 1 To 7 Step 1
            If (X >= flX1) And (X <= flX2) Then
                If tmVef.sType = "G" Then
                    ReDim tgPrg(0 To 1) As PRGDATE  'Time/Dates
                    slNameCode = tmPrgVehicle(igVehIndexViaPrg).sKey  'Traffic!lbcVehicle.List(igVehIndexViaPrg)
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slName, 3, "|", slName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    slStr = slStr & slName & "\" & slCode
                    sgCommandStr = slStr
                    GameSchd.Show vbModal
                    tmGhfSrchKey1.iVefCode = tmVef.iCode
                    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                    If ilRet <> BTRV_ERR_NONE Then
                        tmGhf.lCode = -1
                    End If
                    smTeamTag = ""
                    mTeamPop
                    mDateSpan ""
                Else
                    ilStartDate = True
                    If plcDate.Caption <> "TFN" Then
                        slStr = plcDate.Caption
                        slDate = Format$(gDateValue(slStr) + ilCol - 1, "m/d/yy")
                        If gDateValue(slDate) < lgEarliestDateViaPrg Then
                            ilStartDate = False
                        End If
                        ilDay = gWeekDayStr(slDate)
                    Else
                        ilStartDate = False
                        ilDay = ilCol - 1
                    End If
                    'ilLibType = igLibType
                    'igLibType = imSpecSave(1)
                    'igLibType = ilLibType
                    'MousePointer = vbDefault    'Default
                    ReDim tgPrg(0 To 1) As PRGDATE  'Time/Dates
                    tgPrg(0).sStartTime = slTime
                    If ilStartDate Then
                        tgPrg(0).sStartDate = slDate
                    End If
                    For ilLoop = 0 To 6 Step 1
                        tgPrg(0).iDay(ilLoop) = 0
                    Next ilLoop
                    If igLibType = 0 Then   'TFN allowed
                        'If TV set for one day, if Radio set for TFN
                        tgPrg(0).sEndDate = ""  '"TFN"
                        If ilStartDate Then
                            If ilDay <= 4 Then
                                For ilLoop = ilDay To 4 Step 1
                                    tgPrg(0).iDay(ilLoop) = 1
                                Next ilLoop
                            Else
                                tgPrg(0).iDay(ilDay) = 1
                            End If
                        End If
                    Else    '1=Special; 2=Sports
                        'If TV set for one day, if radio set for either M-F or Sa only, Su only
                        If ilStartDate Then
                            tgPrg(0).sEndDate = slDate
                            tgPrg(0).iDay(ilDay) = 1    'Yes
                        Else
                            tgPrg(0).sEndDate = ""
                        End If
                    End If
                    lgLibLength = lmPrgLength
                    PrgDates.Show vbModal
                    If UBound(tgPrg) > 0 Then
                        Screen.MousePointer = vbHourglass
                        'If igViewType = 1 Then
                        '    slType = "A"
                        'Else
                        '    slType = "O"
                        'End If
                        ilType = 0
                        ilRet = btrBeginTrans(hmLvf, 1000)
                        If ilRet <> BTRV_ERR_NONE Then
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Insert Not Completed, Try Later", vbOKOnly + vbExclamation, "Program")
                            Exit Sub
                        End If
                        ilRet = gPrgToPend(Program, tmLvf, ilType)
                        If Not ilRet Then
                            ilRet = btrAbortTrans(hmLvf)
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Insert Not Completed, Try Later", vbOKOnly + vbExclamation, "Program")
                            Exit Sub
                        End If
                        ilRet = btrEndTrans(hmLvf)
                        imCurrent = 1   'Force to pending
                        pbcCurrent_Paint
                        Screen.MousePointer = vbHourglass
                        imDatePaint = False
                        mDateSpan tgPrg(0).sStartDate
                        Screen.MousePointer = vbHourglass  'Wait
                        pbcCount.Cls
                        pbcLayout.Cls
                        pbcLibrary.Cls
                        If rbcType(2).Value Then
                            pbcCount_Paint
                        ElseIf rbcType(1).Value Then
                            If vbcLayout.Value <> vbcLayout.Min Then
                                vbcLayout.Value = vbcLayout.Min
                            Else
                                pbcLayout_Paint
                            End If
                        Else
                            If vbcLayout.Value <> vbcLayout.Min Then
                                vbcLayout.Value = vbcLayout.Min
                            Else
                                pbcLibrary_Paint
                            End If
                        End If
                        Screen.MousePointer = vbDefault
                    End If
                End If
                Exit Sub
            End If
            flX1 = flX1 + imWidthToNextColumn
            flX2 = flX1 + imWidthWithinColumn
        Next ilCol
    End If
End Sub
Private Sub pbcLibrary_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    Dim llStart As Long
    Dim llTotalTime As Long
    Dim llTime As Long
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slTime As String
    Dim ilCol As Integer
    Dim flX1 As Single
    Dim flX2 As Single
    Dim slStr As String
    Dim slDate As String
    Dim ilFound As Integer
    If imDragSource = 0 Then
        If State = vbLeave Then
            plcDragTime.Visible = False
            Exit Sub
        Else
            If plcDate.Caption <> "TFN" Then
                ilFound = False
                flX1 = imStartXFirstColumn
                flX2 = flX1 + imWidthWithinColumn
                For ilCol = 1 To 7 Step 1
                    If (X >= flX1) And (X <= flX2) Then
                        slStr = plcDate.Caption

                        If gDateValue(slStr) + ilCol - 1 < lgEarliestDateViaPrg Then
                            'plcDragTime.Caption = ""
                            'plcDragTime.Visible = False
                            'Exit Sub
                            slDate = ""
                        Else
                            slDate = Format$(gDateValue(slStr) + ilCol - 1, "m/d/yy")
                        End If
                        ilFound = True
                        Exit For
                    End If
                    flX1 = flX1 + imWidthToNextColumn
                    flX2 = flX1 + imWidthWithinColumn
                Next ilCol
                If Not ilFound Then
                    smDragTime = ""
                    plcDragTime.Visible = False
                    Exit Sub
                End If
            End If
            If imResol = 0 Then
                llStart = 0
                llTotalTime = 86400
            ElseIf imResol = 1 Then 'Half hour
                llStart = CLng(1800) * (vbcLayout.Value - 1)
                llTotalTime = 43200
            Else    'Quarter hour
                llStart = CLng(900) * (vbcLayout.Value - 1)
                llTotalTime = 21600
            End If
            llTime = llStart + (llTotalTime * (Y - 285)) / fmPaintHeight
            llTime = 900 * ((llTime + 450) \ 900)
            If (llTime >= 0) And (X > 405) And (llTime + lmPrgLength <= 86400) Then
                gPackTimeLong llTime, ilHsSec, ilMinHr
                gUnpackTime ilHsSec, ilMinHr, "A", "2", slTime
                smDragTime = Trim$(slDate & " " & slTime)
                If plcDragTime.Visible Then
                    plcDragTime_Paint
                Else
                    plcDragTime.Visible = True
                End If
            ElseIf (llTime + lmPrgLength > 86400) Then
                llTime = 86400 - lmPrgLength
                gPackTimeLong llTime, ilHsSec, ilMinHr
                gUnpackTime ilHsSec, ilMinHr, "A", "2", slTime
                smDragTime = Trim$(slDate & " " & slTime)
                If plcDragTime.Visible Then
                    plcDragTime_Paint
                Else
                    plcDragTime.Visible = True
                End If
            Else
                smDragTime = ""
                plcDragTime.Visible = False
            End If
        End If
    End If
End Sub
Private Sub pbcLibrary_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    If Button = 2 Then  'Right Mouse
        For ilLoop = LBONE To UBound(tmLCD) - 1 Step 1
            If (Y >= tmLCD(ilLoop).fY1) And (Y <= tmLCD(ilLoop).fY2) And (X >= tmLCD(ilLoop).fX1) And (X <= tmLCD(ilLoop).fX2) Then
                imButtonIndex = ilLoop
                imIgnoreRightMove = True
                mShowLibInfo
                imIgnoreRightMove = False
                Exit Sub
            End If
        Next ilLoop
        Exit Sub
    End If
    imDragSource = 1
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcLibrary_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        For ilLoop = LBONE To UBound(tmLCD) - 1 Step 1
            If (Y >= tmLCD(ilLoop).fY1) And (Y <= tmLCD(ilLoop).fY2) And (X >= tmLCD(ilLoop).fX1) And (X <= tmLCD(ilLoop).fX2) Then
                If (imButtonIndex = ilLoop) And (plcLibInfo.Visible) Then
                    Exit Sub
                End If
                imButtonIndex = ilLoop
                imIgnoreRightMove = True
                mShowLibInfo
                imIgnoreRightMove = False
                Exit Sub
            End If
        Next ilLoop
        plcLibInfo.Visible = False
        Exit Sub
    End If
End Sub
Private Sub pbcLibrary_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        imButtonIndex = -1
        plcLibInfo.Visible = False
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcLibrary_Paint()
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim clStart As Currency
    Dim clInc As Currency
    Dim clCntStart As Currency
    Dim clCntEnd As Currency
    Dim ilListIndex As Integer
    'Paint times
'    If ((rbcType(0).Value <> False) And (imResol = 0)) Or (rbcType(0).Value = False) Then
    mPaintColumns
    If imResol = 0 Then
        mPaintTimes 1, 96, 4, pbcLibrary
        'Print Dates
        mPaintDates pbcLibrary
        clStart = 0
        clInc = 3600
    ElseIf imResol = 1 Then 'Half hour
        ilStart = 2 * (vbcLayout.Value - 1) + 1
        ilEnd = 2 * (vbcLayout.Value + vbcLayout.LargeChange)
        mPaintTimes ilStart, ilEnd, 2, pbcLibrary
        'Print Dates
        mPaintDates pbcLibrary
        clStart = 1800 * CLng((vbcLayout.Value - 1))
        clInc = 1800
    Else    'Quarter hour
        ilStart = vbcLayout.Value
        ilEnd = vbcLayout.Value + vbcLayout.LargeChange
        mPaintTimes ilStart, ilEnd, 1, pbcLibrary
        'Print Dates
        mPaintDates pbcLibrary
        clStart = 900 * CLng((vbcLayout.Value - 1))
        clInc = 900
    End If
    imNoEvt = 0
    imUnits = 0
    cmSec = 0
    clCntStart = 0
    clCntEnd = 86399
    'If lbcETypeCode.ListCount > 1 Then
    If UBound(tmETypeCode) > 1 Then
        'If Asc(lbcETypeCode.List(0)) = Asc("2") Then
        If Asc(tmETypeCode(0).sKey) = Asc("2") Then
            ilListIndex = 0
        'ElseIf Asc(lbcETypeCode.List(1)) = Asc("2") Then
        ElseIf Asc(tmETypeCode(1).sKey) = Asc("2") Then
            ilListIndex = 1
        Else
            ilListIndex = -1
        End If
    Else
        ilListIndex = -1
    End If
    If plcDate.Caption <> "TFN" Then
        If imCurrent = 0 Then
            mPaintLibrary clStart, clInc, tmCLLC(), tmPLLC()
            mWkCntForLibrary ilListIndex, clCntStart, clCntEnd, tmCLLC(), tmPLLC()
        Else
            mPaintLibrary clStart, clInc, tmCDLLC(), tmPDLLC()
            mWkCntForLibrary ilListIndex, clCntStart, clCntEnd, tmCDLLC(), tmPDLLC()
        End If
    Else
        If imCurrent = 0 Then
            mPaintLibrary clStart, clInc, tmCTFN(), tmPTFN()
            mWkCntForLibrary ilListIndex, clCntStart, clCntEnd, tmCTFN(), tmPTFN()
        Else
            mPaintLibrary clStart, clInc, tmCDTFN(), tmPDTFN()
            mWkCntForLibrary ilListIndex, clCntStart, clCntEnd, tmCDTFN(), tmPDTFN()
        End If
    End If
    mWkCntPaint ilListIndex
End Sub

Private Sub pbcResolType_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub pbcResolType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcResolType_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(" ") Then
        Screen.MousePointer = vbHourglass  'Wait
        If imResol = 0 Then
            imResol = 1
            vbcLayout.Max = 25
        ElseIf imResol = 1 Then
            imResol = 2
            vbcLayout.Max = 73
        Else
            imResol = 0
            vbcLayout.Max = 1
        End If
        If rbcType(0).Value Then
            pbcLibrary.Cls
            pbcResolType_Paint
            vbcLayout.Value = vbcLayout.Min
            pbcLibrary_Paint
        ElseIf rbcType(1).Value Then
            pbcLayout.Cls
            pbcResolType_Paint
            vbcLayout.Value = vbcLayout.Min
            pbcLayout_Paint
        ElseIf rbcType(2).Value Then
            pbcCount.Cls
            pbcResolType_Paint
            vbcLayout.Value = vbcLayout.Min
            pbcCount_Paint
        End If
        Screen.MousePointer = vbDefault    'Default
    End If
End Sub
Private Sub pbcResolType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    Screen.MousePointer = vbHourglass  'Wait
    If imResol = 0 Then
        imResol = 1
        vbcLayout.Max = 25
    ElseIf imResol = 1 Then
        imResol = 2
        vbcLayout.Max = 73
    Else
        imResol = 0
        vbcLayout.Max = 1
    End If
    If rbcType(0).Value Then
        pbcLibrary.Cls
        pbcResolType_Paint
        vbcLayout.Value = vbcLayout.Min
        pbcLibrary_Paint
    ElseIf rbcType(1).Value Then
        pbcLayout.Cls
        pbcResolType_Paint
        vbcLayout.Value = vbcLayout.Min
        pbcLayout_Paint
    ElseIf rbcType(2).Value Then
        pbcCount.Cls
        pbcResolType_Paint
        vbcLayout.Value = vbcLayout.Min
        pbcCount_Paint
    End If
    Screen.MousePointer = vbDefault    'Default
End Sub
Private Sub pbcResolType_Paint()
    If rbcType(0).Value Then
        If imResol = 0 Then
            vbcLayout.Visible = False
            plcLayout.Move plcLib.Left + plcLib.Width + 120 + vbcLayout.Width \ 2, implcLayoutTop, pbcLibrary.Width + fgPanelAdj
            pbcLibrary.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
        Else
            plcLayout.Move plcLib.Left + plcLib.Width + 120, implcLayoutTop, pbcLibrary.Width + vbcLayout.Width + fgPanelAdj
            pbcLibrary.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
            vbcLayout.Move pbcLibrary.Left + pbcLibrary.Width + 15, pbcLibrary.Top
            vbcLayout.Visible = True
        End If
    ElseIf rbcType(1).Value Then
        If imResol = 0 Then
            vbcLayout.Visible = False
            plcLayout.Move plcLib.Left + plcLib.Width + 120 + vbcLayout.Width \ 2, implcLayoutTop, pbcLayout.Width + fgPanelAdj
            pbcLayout.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
        Else
            plcLayout.Move plcLib.Left + plcLib.Width + 120, implcLayoutTop, pbcLayout.Width + vbcLayout.Width + fgPanelAdj
            pbcLayout.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
            vbcLayout.Move pbcLayout.Left + pbcLayout.Width + 15, pbcLayout.Top
            vbcLayout.Visible = True
        End If
    ElseIf rbcType(2).Value Then
        If imResol = 0 Then
            vbcLayout.Visible = False
            plcLayout.Move plcLib.Left + plcLib.Width + 120 + vbcLayout.Width \ 2, implcLayoutTop, pbcCount.Width + fgPanelAdj
            pbcCount.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
        Else
            plcLayout.Move plcLib.Left + plcLib.Width + 120, implcLayoutTop, pbcCount.Width + vbcLayout.Width + fgPanelAdj
            pbcCount.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
            vbcLayout.Move pbcCount.Left + pbcCount.Width + 15, pbcCount.Top
            vbcLayout.Visible = True
        End If
    End If
    plcWkCnt.Move plcLayout.Width - plcWkCnt.Width - fgBevelX
    hbcDate.Move hbcDate.Left, hbcDate.Top, plcWkCnt.Left - plcDate.Width
    pbcResolType.Cls
    pbcResolType.CurrentX = fgBoxInsetX
    pbcResolType.CurrentY = -15 'fgBoxInsetY
    If imResol = 0 Then
        pbcResolType.Print "Hour"
    ElseIf imResol = 1 Then
        pbcResolType.Print "1/2 Hour"
    ElseIf imResol = 2 Then
        pbcResolType.Print "15 Mins"
    End If
End Sub
Private Sub plcDate_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub plcDate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcLayout_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcLayout_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub

Private Sub plcLib_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcLib_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub plcLib_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcResol_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub plcResol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcType_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcType_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub plcType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcWkCnt_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub plcWkCnt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub rbcType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcType(Index).Value
    'End of coded added
    If Value Then
        ReDim tmLCD(0 To 1) As LCD
        'cmcDupl.Enabled = False
        If tmVef.sType <> "G" Then
            cmcDefSchd.Enabled = False
        End If
        cmcLib.Enabled = False
        Select Case Index
            Case 0  'Library
                If imLibLayCnt <> 0 Then
                    imLibLayCnt = 0
                    mLibPop
                End If
                pbcLibrary.Cls
                pbcCount.Visible = False
                pbcLayout.Visible = False
                If imResol = 0 Then
                    vbcLayout.Visible = False
                    plcLayout.Move plcLib.Left + plcLib.Width + 120 + vbcLayout.Width \ 2, implcLayoutTop, pbcLibrary.Width + fgPanelAdj
                    pbcLibrary.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
                Else
                    plcLayout.Move plcLib.Left + plcLib.Width + 120, implcLayoutTop, pbcLibrary.Width + vbcLayout.Width + fgPanelAdj
                    pbcLibrary.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
                    vbcLayout.Move pbcLibrary.Left + pbcLibrary.Width + 15, pbcLibrary.Top
                    vbcLayout.Visible = True
                End If
                plcWkCnt.Move plcLayout.Width - plcWkCnt.Width - fgBevelX
                hbcDate.Move hbcDate.Left, hbcDate.Top, plcWkCnt.Left - plcDate.Width
                plcResol.Visible = True
                ckcShowVersion.Visible = True
                pbcLibrary.Visible = True
                pbcLibrary_Paint
            Case 1  'Layout
                If (imLibLayCnt <> 1) And (imLibLayCnt <> 2) Then
                    imLibLayCnt = 1
                    mLibPop
                End If
                pbcLayout.Cls
                pbcCount.Visible = False
                pbcLibrary.Visible = False
                If imResol = 0 Then
                    vbcLayout.Visible = False
                    plcLayout.Move plcLib.Left + plcLib.Width + 120 + vbcLayout.Width \ 2, implcLayoutTop, pbcLayout.Width + fgPanelAdj
                    pbcLayout.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
                Else
                    plcLayout.Move plcLib.Left + plcLib.Width + 120, implcLayoutTop, pbcLayout.Width + vbcLayout.Width + fgPanelAdj
                    pbcLayout.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
                    vbcLayout.Move pbcLayout.Left + pbcLayout.Width + 15, pbcLayout.Top
                    vbcLayout.Visible = True
                End If
                plcWkCnt.Move plcLayout.Width - plcWkCnt.Width - fgBevelX
                hbcDate.Move hbcDate.Left, hbcDate.Top, plcWkCnt.Left - plcDate.Width
                plcResol.Visible = True
                ckcShowVersion.Visible = False
                pbcLayout.Visible = True
                pbcLayout_Paint
            Case 2  'Count
                If (imLibLayCnt <> 1) And (imLibLayCnt <> 2) Then
                    imLibLayCnt = 2
                    mLibPop
                End If
                pbcCount.Cls
                pbcLayout.Visible = False
                pbcLibrary.Visible = False
                If imResol = 0 Then
                    vbcLayout.Visible = False
                    plcLayout.Move plcLib.Left + plcLib.Width + 120 + vbcLayout.Width \ 2, implcLayoutTop, pbcCount.Width + fgPanelAdj
                    pbcCount.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
                Else
                    plcLayout.Move plcLib.Left + plcLib.Width + 120, implcLayoutTop, pbcCount.Width + vbcLayout.Width + fgPanelAdj
                    pbcCount.Move plcLayout.Left + fgBevelX, plcLayout.Top + fgBevelY
                    vbcLayout.Move pbcCount.Left + pbcCount.Width + 15, pbcCount.Top
                    vbcLayout.Visible = True
                End If
                plcWkCnt.Move plcLayout.Width - plcWkCnt.Width - fgBevelX
                hbcDate.Move hbcDate.Left, hbcDate.Top, plcWkCnt.Left - plcDate.Width
                plcResol.Visible = True
                ckcShowVersion.Visible = False
                pbcCount.Visible = True
                pbcCount_Paint
                imLibLayCnt = 2
        End Select
    End If
End Sub
Private Sub rbcType_DragDrop(Index As Integer, Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub rbcType_GotFocus(Index As Integer)
    If imFirstTime Then
        imFirstTime = False
    End If
    If imFirstFocus Then
        cbcVeh.SetFocus
        imFirstFocus = False
    End If
End Sub
Private Sub rbcType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub tmcClick_Timer()
    If imSelectDelay Then
        imSelectDelay = False
        mCbcVehChange
    End If
End Sub
Private Sub tmcDrag_Timer()
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilIndex As Integer
    Dim llTimeResol As Long
    Dim llTime As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Select Case imDragType
        Case 0  'Start Drag
            tmcDrag.Enabled = False
            If imDragSource = 0 Then
                ilIndex = fmDragY \ fgListHtArial825
                If ilIndex <> lbcLib.ListIndex - lbcLib.TopIndex Then
                    Exit Sub
                End If
                If (ilIndex = 0) And (lbcLib.TopIndex = 0) Then
                    Exit Sub
                End If
                slNameCode = tmLibName(lbcLib.ListIndex - 1).sKey   'lbcLibName.List(ilIndex - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet <> CP_MSG_NONE Then
                    Exit Sub
                End If
                If CLng(slCode) <> tmLvf.lCode Then
                    tmLvfSrchKey.lCode = CLng(slCode)
                    ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Sub
                    End If
                End If
                If imResol = 0 Then
                    llTimeResol = 86400
                ElseIf imResol = 1 Then
                    llTimeResol = 43200
                Else
                    llTimeResol = 21600
                End If
                plcDragTime.Move plcLayout.Left - (2 * plcDragTime.Width) \ 3, plcLayout.Top + 30
                DoEvents
                lacLibFrame(imDragSource).DragIcon = IconTraf!imcIconStd.DragIcon
                gUnpackTimeLong tmLvf.iLen(0), tmLvf.iLen(1), False, llTime
                lmPrgLength = llTime
                'lacLibFrame(imDragSource).Move plcLib.Left + lbcLib.Left, plcLib.Top + lbcLib.Top + fgListHtArial825 * (ilIndex), lacLibFrame(imDragSource).Width, (fmPaintHeight * llTime) \ llTimeResol
                lacLibFrame(imDragSource).Move fmDragX - lacLibFrame(imDragSource).Width \ 3, plcLib.Top + lbcLib.Top + fgListHtArial825 * (ilIndex), lacLibFrame(imDragSource).Width, (fmPaintHeight * llTime) \ llTimeResol
                lacLibFrame(imDragSource).Visible = True
                lacLibFrame(imDragSource).Drag vbBeginDrag
                lacLibFrame(imDragSource).DragIcon = IconTraf!imcIconDrag.DragIcon
                Exit Sub
            ElseIf imDragSource = 1 Then
                imDragType = -1
                'Moved the imUsingTFNForDate below- disallow TFN moved to date to be dragged to trash
                'If (plcDate.Caption = "TFN") Or (imUsingTFNForDate) Then
                If (plcDate.Caption = "TFN") Then
                    Exit Sub
                End If
                For ilLoop = LBONE To UBound(tmLCD) - 1 Step 1
                    If (fmDragY >= tmLCD(ilLoop).fY1) And (fmDragY <= tmLCD(ilLoop).fY2) And (fmDragX >= tmLCD(ilLoop).fX1) And (fmDragX <= tmLCD(ilLoop).fX2) Then
                        slDate = plcDate.Caption
                        llDate = gDateValue(slDate) + tmLCD(ilLoop).iDay - 1
                        If lgEarliestDateViaPrg > 0 Then
                            If llDate < lgEarliestDateViaPrg Then
                                Exit Sub
                            End If
                        End If
                        If (tmLCD(ilLoop).iCurOrPend = 0) And (imUsingTFNForDate) Then
                            Exit Sub
                        End If
                        imLCDDragIndex = ilLoop
                        lacLibFrame(imDragSource).DragIcon = IconTraf!imcIconStd.DragIcon
                        lacLibFrame(imDragSource).Move tmLCD(ilLoop).fX1, tmLCD(ilLoop).fY1, lacLibFrame(imDragSource).Width, tmLCD(ilLoop).fY2 - tmLCD(ilLoop).fY1
                        'If gInvertArea call then remove visible setting
                        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
                        lacLibFrame(imDragSource).Visible = True
                        imcTrash.Visible = True
                        imcTrash.Enabled = True
                        'imcHelp.Visible = False
                        lacLibFrame(imDragSource).Drag vbBeginDrag
                        lacLibFrame(imDragSource).DragIcon = IconTraf!imcIconDrag.DragIcon
                        Exit Sub
                    End If
                Next ilLoop
            End If
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub vbcLayout_Change()
    Screen.MousePointer = vbHourglass  'Wait
    If rbcType(0).Value Then
        pbcLibrary.Cls
        pbcLibrary_Paint
    ElseIf rbcType(1).Value Then
        pbcLayout.Cls
        pbcLayout_Paint
    ElseIf rbcType(2).Value Then
        pbcCount.Cls
        pbcCount_Paint
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub vbcLayout_DragDrop(Source As control, X As Single, Y As Single)
    lacLibFrame(imDragSource).Visible = False
    imDragType = -1
    imcTrash.Visible = False
    'imcHelp.Visible = True
End Sub
Private Sub plcDragTime_Paint()
    plcDragTime.Cls
    plcDragTime.CurrentX = 30
    plcDragTime.CurrentY = 0
    plcDragTime.Print smDragTime
End Sub
Private Sub plcResol_Paint()
    plcResol.CurrentX = 0
    plcResol.CurrentY = 0
    plcResol.Print "Resolution"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Programming"
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTeamPop                        *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Tema list box         *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mTeamPop()
'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer
    ilRet = gObtainMnfForType("Z", smTeamTag, tmTeam())
    Exit Sub
End Sub


Private Sub mPaintColumns()
    Dim llX1 As Long
    Dim llX2 As Long
    Dim llY As Long
    Dim ilCol As Integer
    Dim llTop As Long
    Dim ilLineCount As Integer

    llX1 = imStartXFirstColumn
    llX2 = llX1 + imWidthWithinColumn
    llY = 15
    For ilCol = 1 To 7 Step 1
        If rbcType(0).Value Then
            pbcLibrary.Line (llX1 - 15, 15)-Step(imWidthWithinColumn + 30, 165 + 15), BLUE, B
            pbcLibrary.Line (llX1, 30)-Step(imWidthWithinColumn, 150), LIGHTYELLOW, BF
            pbcLibrary.Line (llX1 - 15, 180)-Step(imWidthWithinColumn + 30, pbcLibrary.height - 195), BLUE, B
            pbcLibrary.Line (llX1, 195)-Step(imWidthWithinColumn, pbcLibrary.height - 240), LIGHTBLUE, BF
        ElseIf rbcType(1).Value Then
            pbcLayout.Line (llX1 - 15, 15)-Step(imWidthWithinColumn + 30, 165 + 15), BLUE, B
            pbcLayout.Line (llX1, 30)-Step(imWidthWithinColumn, 150), LIGHTYELLOW, BF
            pbcLayout.Line (llX1 - 15, 180)-Step(imWidthWithinColumn + 30, pbcLayout.height - 195), BLUE, B
            pbcLayout.Line (llX1, 195)-Step(imWidthWithinColumn, pbcLayout.height - 240), LIGHTYELLOW, BF
        ElseIf rbcType(2).Value Then
            pbcCount.Line (llX1 - 15, 15)-Step(imWidthWithinColumn + 30, 165), BLUE, B
            pbcCount.Line (llX1, 30)-Step(imWidthWithinColumn, 150), LIGHTYELLOW, BF
            pbcCount.Line (llX1 - 15, 180)-Step(imWidthWithinColumn + 30, pbcCount.height - 195), BLUE, B
            pbcCount.Line (llX1, 195)-Step(imWidthWithinColumn, pbcCount.height - 240), LIGHTYELLOW, BF
        End If
        llX1 = llX1 + imWidthToNextColumn
        llX2 = llX1 + imWidthWithinColumn
    Next ilCol
    If rbcType(0).Value Then
        llTop = imStartYFirstColumn
        ilLineCount = 0
        Do
            If ilLineCount Mod 2 = 0 Then
                pbcLibrary.Line (imStartXFirstColumn / 2, llTop - 30)-(imStartXFirstColumn - 15, llTop - 30), BLUE, B
            Else
                pbcLibrary.Line ((2 * imStartXFirstColumn) / 3, llTop - 30)-(imStartXFirstColumn - 15, llTop - 30), BLUE, B
            End If
            ilLineCount = ilLineCount + 1
            llTop = llTop + imHeightWithinColumn + 15
        Loop While llTop + imStartYFirstColumn < pbcLibrary.height
    ElseIf rbcType(1).Value Then
        llTop = imStartYFirstColumn
        ilLineCount = 0
        Do
            If ilLineCount Mod 2 = 0 Then
                pbcLayout.Line (imStartXFirstColumn / 2, llTop - 30)-(imStartXFirstColumn - 15, llTop - 30), BLUE, B
            Else
                pbcLayout.Line ((2 * imStartXFirstColumn) / 3, llTop - 30)-(imStartXFirstColumn - 15, llTop - 30), BLUE, B
            End If
            ilLineCount = ilLineCount + 1
            llTop = llTop + imHeightWithinColumn + 15
        Loop While llTop + imStartYFirstColumn < pbcLayout.height
    ElseIf rbcType(2).Value Then
        llTop = imStartYFirstColumn
        ilLineCount = 0
        Do
            pbcCount.Line (imStartXFirstColumn - 15, llTop - 30)-(imStartXFirstColumn + 7 * (imWidthToNextColumn + 30) - 15, llTop - 30), BLUE, B
            ilLineCount = ilLineCount + 1
            llTop = llTop + imHeightWithinColumn + 15
        Loop While llTop + imStartYFirstColumn < pbcCount.height
    End If
    'vbcLogs.LargeChange = ilLineCount - 1
End Sub

Private Sub mComputeXY(llFdTime As Long, ilEndTime As Integer, llY As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llTime2                                                                               *
'******************************************************************************************

    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim ilStep As Integer
    Dim llTime1 As Long
    Dim ilRow As Integer
    Dim slTime As String
    Dim llDelta As Long

    If imResol = 0 Then
        ilStart = 1
        ilEnd = 96
        ilStep = 4
        llDelta = 3600
    ElseIf imResol = 1 Then 'Half hour
        ilStart = 2 * (vbcLayout.Value - 1) + 1
        ilEnd = 2 * (vbcLayout.Value + vbcLayout.LargeChange)
        ilStep = 2
        llDelta = 1800
    Else    'Quarter hour
        ilStart = vbcLayout.Value
        ilEnd = vbcLayout.Value + vbcLayout.LargeChange
        ilStep = 1
        llDelta = 900
    End If
    llTime1 = -1
    llY = 0 'imStartYFirstColumn
    For ilRow = ilStart To ilEnd Step ilStep
        'If tmVef.sType <> "G" Then
            If ilRow <= 4 Then
                Select Case ilRow Mod 4
                    Case 1
                        slTime = "12AM"
                    Case 2
                        slTime = "12:15AM"
                    Case 3
                        slTime = "12:30AM"
                    Case Else
                        slTime = "12:45AM"
                End Select
            ElseIf (ilRow - 1) \ 4 < 12 Then
                Select Case ilRow Mod 4
                    Case 1
                        slTime = Trim$(str$((ilRow - 1) \ 4)) & "AM"
                    Case 2
                        slTime = Trim$(str$((ilRow - 1) \ 4)) & ":15AM"
                    Case 3
                        slTime = Trim$(str$((ilRow - 1) \ 4)) & ":30AM"
                    Case Else
                        slTime = Trim$(str$((ilRow - 1) \ 4)) & ":45AM"
                End Select
            ElseIf (ilRow - 1) \ 4 = 12 Then
                Select Case ilRow Mod 4
                    Case 1
                        slTime = "12PM"
                    Case 2
                        slTime = "12:15PM"
                    Case 3
                        slTime = "12:30PM"
                    Case Else
                        slTime = "12:45PM"
                End Select
            Else
                Select Case ilRow Mod 4
                    Case 1
                        slTime = Trim$(str$((ilRow - 1) \ 4 - 12)) & "PM"
                    Case 2
                        slTime = Trim$(str$((ilRow - 1) \ 4 - 12)) & ":15PM"
                    Case 3
                        slTime = Trim$(str$((ilRow - 1) \ 4 - 12)) & ":30PM"
                    Case Else
                        slTime = Trim$(str$((ilRow - 1) \ 4 - 12)) & ":45PM"
                End Select
            End If
            If llTime1 = -1 Then
                llTime1 = gTimeToLong(slTime, False)
                If (llFdTime < llTime1) And (ilEndTime = False) Then
                    llY = imStartYFirstColumn - 30
                    Exit Sub
                End If
                If (llFdTime < llTime1) And (ilEndTime = True) Then
                    llY = -1
                    Exit Sub
                End If
            Else
                llTime1 = gTimeToLong(slTime, ilEndTime)
            End If
        'Else
        '    If ilRow <= 4 Then
        '        slTime = "0"
        '        Select Case ilRow Mod 4
        '            Case 1
        '                slTime = slTime & ":00"
        '            Case 2
        '                slTime = " " & slTime & ":15" '"12:15"
        '            Case 3
        '                slTime = " " & slTime & ":30" '"12:30"
        '            Case Else
        '                slTime = " " & slTime & ":45" '"12:45"
        '        End Select
        '    ElseIf (ilRow - 1) \ 4 < 12 Then
        '        slTime = Trim$(Str$((ilRow - 1) \ 4))
        '        Select Case ilRow Mod 4
        '            Case 1
        '                slTime = slTime & ":00"
        '            Case 2
        '                slTime = " " & slTime & ":15" 'Trim$(Str$((ilRow - 1) \ 4)) & ":15"
        '            Case 3
        '                slTime = " " & slTime & ":30" 'Trim$(Str$((ilRow - 1) \ 4)) & ":30"
        '            Case Else
        '                slTime = " " & slTime & ":45" 'Trim$(Str$((ilRow - 1) \ 4)) & ":45"
        '        End Select
        '    ElseIf (ilRow - 1) \ 4 = 12 Then
        '        slTime = "12"
        '        Select Case ilRow Mod 4
        '            Case 1
        '                slTime = slTime & ":00"
        '            Case 2
        '                slTime = " " & slTime & ":15" '"12:15"
        '            Case 3
        '                slTime = " " & slTime & ":30" '"12:30"
        '            Case Else
        '                slTime = " " & slTime & ":45" '"12:45"
        '        End Select
        '    Else
        '        slTime = Trim$(Str$((ilRow - 1) \ 4 - 12))
        '        Select Case ilRow Mod 4
        '            Case 1
        '                slTime = slTime & ":00"
        '            Case 2
        '                slTime = " " & slTime & ":15" 'Trim$(Str$((ilRow - 1) \ 4 - 12)) & ":15"
        '            Case 3
        '                slTime = " " & slTime & ":30" 'Trim$(Str$((ilRow - 1) \ 4 - 12)) & ":30"
        '            Case Else
        '                slTime = " " & slTime & ":45" 'Trim$(Str$((ilRow - 1) \ 4 - 12)) & ":45"
        '        End Select
        '    End If
        'End If
        If (llFdTime >= llTime1) And (llFdTime <= llTime1 + llDelta) Then
            If llFdTime = llTime1 Then
                llY = llY
            ElseIf llFdTime = llTime1 + llDelta Then
                llY = llY + imHeightWithinColumn + 15
            Else
                llY = llY + (imHeightWithinColumn + 15) * (CSng(llFdTime - llTime1) / (llDelta))
                If ilEndTime Then
                    Do While llY Mod 15
                        llY = llY + 1
                    Loop
                Else
                    Do While llY Mod 15
                        llY = llY - 1
                    Loop
                End If
            End If
            llY = llY + imStartYFirstColumn - 30
            Exit Sub
        End If

        llY = llY + imHeightWithinColumn + 15  'fgBoxGridH
    Next ilRow
    If (llFdTime > llTime1 + llDelta) And (ilEndTime = False) Then
        llY = -1
        Exit Sub
    End If
    If (llFdTime > llTime1 + llDelta) And (ilEndTime = True) Then
        llY = llY + imHeightWithinColumn + 15
        Exit Sub
    End If
End Sub
