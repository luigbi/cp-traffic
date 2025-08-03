VERSION 5.00
Begin VB.Form LinkDlvy 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   1365
   ClientWidth     =   9420
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
   ScaleHeight     =   5955
   ScaleWidth      =   9420
   Begin VB.CommandButton cmcPrefeed 
      Appearance      =   0  'Flat
      Caption         =   "&Pre-Feed"
      Height          =   285
      Left            =   6885
      TabIndex        =   32
      Top             =   5580
      Width           =   1140
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Left            =   30
      Picture         =   "Linkdlvy.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1305
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1230
      Left            =   4245
      Picture         =   "Linkdlvy.frx":030A
      ScaleHeight     =   1200
      ScaleWidth      =   3030
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2985
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.ListBox lbcFeed 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4215
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   600
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1155
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   30
         Picture         =   "Linkdlvy.frx":435C
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   30
         Width           =   1020
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "Linkdlvy.frx":501A
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   345
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
      End
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
      Left            =   2310
      Picture         =   "Linkdlvy.frx":5324
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
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
      Left            =   1290
      MaxLength       =   20
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8115
      ScaleHeight     =   210
      ScaleWidth      =   375
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1695
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lbcSubfeed 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2415
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcEvtName 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   3870
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2805
      Visible         =   0   'False
      Width           =   3210
   End
   Begin VB.ListBox lbcEvtType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1350
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3210
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2790
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3585
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox plcSort 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   60
      ScaleHeight     =   255
      ScaleWidth      =   8865
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5235
      Width           =   8865
      Begin VB.OptionButton rbcSort 
         Caption         =   "Schd"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   8115
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   0
         Width           =   720
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Bus"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   7470
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   0
         Width           =   615
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Cmml"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   6645
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   0
         Width           =   795
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Prog"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   5925
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Width           =   720
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   5130
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   0
         Width           =   810
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Type"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   4410
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   765
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Feed"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   3015
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Affiliate"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   2085
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   945
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Zone"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   3705
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Width           =   750
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Air"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1560
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   600
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Vehicle"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   645
         TabIndex        =   17
         Top             =   0
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.VScrollBar vbcDelivery 
      Height          =   4860
      LargeChange     =   22
      Left            =   9030
      TabIndex        =   15
      Top             =   285
      Width           =   240
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8205
      Top             =   5535
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8445
      Top             =   5580
   End
   Begin VB.PictureBox pbcSTab 
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
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   195
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
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
      Height          =   75
      Left            =   15
      ScaleHeight     =   75
      ScaleWidth      =   60
      TabIndex        =   14
      Top             =   5655
      Width           =   60
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3015
      TabIndex        =   29
      Top             =   5580
      Width           =   1140
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1725
      TabIndex        =   28
      Top             =   5580
      Width           =   1140
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   4320
      TabIndex        =   30
      Top             =   5580
      Width           =   1140
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
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   105
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5085
      Width           =   105
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   6135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -30
      Width           =   6135
   End
   Begin VB.CommandButton cmcDupl 
      Appearance      =   0  'Flat
      Caption         =   "D&uplicate"
      Height          =   285
      Left            =   5595
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1140
   End
   Begin VB.ListBox lbcEvtNameCode 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   2925
      Sorted          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   -30
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ListBox lbcSortIndex 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4035
      Sorted          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   -75
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ListBox lbcVehCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4980
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   -150
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ListBox lbcFeedCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3735
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   -90
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox pbcDelivery 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4860
      Index           =   1
      Left            =   225
      Picture         =   "Linkdlvy.frx":541E
      ScaleHeight     =   4860
      ScaleWidth      =   8790
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   420
      Width           =   8790
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   38
         Top             =   495
         Visible         =   0   'False
         Width           =   8790
      End
   End
   Begin VB.PictureBox pbcDelivery 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4860
      Index           =   0
      Left            =   225
      Picture         =   "Linkdlvy.frx":33F2C
      ScaleHeight     =   4860
      ScaleWidth      =   8790
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   8790
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   34
         Top             =   510
         Visible         =   0   'False
         Width           =   8790
      End
   End
   Begin VB.PictureBox plcDelivery 
      ForeColor       =   &H00000000&
      Height          =   4995
      Left            =   180
      ScaleHeight     =   4935
      ScaleWidth      =   9090
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   195
      Width           =   9150
   End
   Begin VB.Label plcScroll 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   6465
      TabIndex        =   43
      Top             =   -15
      Width           =   2865
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   600
      Picture         =   "Linkdlvy.frx":62A3A
      Top             =   5625
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   5535
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8895
      Picture         =   "Linkdlvy.frx":62D44
      Top             =   5385
      Width           =   480
   End
End
Attribute VB_Name = "LinkDlvy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Linkdlvy.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: LinkDlvy.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Copy Ratio input screen code
Option Explicit
Option Compare Text
Dim tmNameCode() As SORTCODE
Dim smNameCodeTag As String
Dim tmEvtTypeCode() As SORTCODE
Dim smEvtTypeCodeTag As String
Dim tmSubFeedCode() As SORTCODE
Dim smSubFeedCodeTag As String
Dim imDelOrEngr As Integer  '0=Delivery; 1=Engineering
Dim imFirstActivate As Integer
Dim tmCtrls(0 To 12)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer       'Current event box
Dim imRowNo As Integer  'Current row number in event area (start at 0)
Dim smShow() As String  'Values shown in delivery area
Dim smSave() As String  'Values saved (1=Vehicle;2=Air time; 3=Local Time; 4=Feed Time;
                        '5=Time zone; 6=Event type; 7=Event name; 8=Program code;
                        '9=Bus; 10=Schedule
Dim imSave() As Integer 'Values saved (1= Show on Cmml sch)
Dim imSort(0 To 2) As Integer   'Major to minor
'Btrieve file variables
'Log calendar file
Dim hmLcf As Integer            'Log calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim tmLcfSrchKey As LCFKEY0     'LCF key record image
Dim imLcfRecLen As Integer         'LCF record length
'LEF Variables
Dim hmLef As Integer            'Log event file handle
Dim tmLef As LEF                'LEF record image
Dim tmLefSrchKey As LEFKEY0     'LEF Key 0 image
Dim imLefRecLen As Integer      'LEF record length
'LVF Variables
Dim hmLvf As Integer            'Log version file handle
Dim tmLvf As LVF                'LVF record image
Dim tmLvfSrchKey As LONGKEY0     'LVF Key 0 image
Dim imLvfRecLen As Integer      'LVF record length
'Delivery links and Engineering are of the same format
Dim hmDlf As Integer            'Delivery Vehicle link file handle
Dim tmDlf() As DLFLIST                'DLF record image
Dim tmDlfSrchKey As DLFKEY0            'DLF record image
Dim tmDlfSrchKey1 As LONGKEY0
Dim imDlfRecLen As Integer        'VLF record length
Dim imDlfIndex As Integer       'Index into Dlf for specified row
Dim hmPff As Integer
Dim tmPFF As PFF        'GSF record image
Dim imPffRecLen As Integer        'GSF record length
Dim tmPffSrchKey1 As PFFKEY1
Dim imSortRowIndex As Integer
Dim imPrevIndex As Integer      'Previous row index to imDlfIndex
Dim imSvRowNo As Integer        'Row number for found value
Dim imMaxRowNo As Integer
'Vehicle file
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer         'VEF record length
'Feed file
Dim hmMnf As Integer            'Multi-name file handle
Dim tmMnf As MNF                'MNF record image
Dim tmMnfSrchKey As INTKEY0            'MNF record image
Dim imMnfRecLen As Integer         'MNF record length
'ENF Variables
Dim hmEnf As Integer            'Event name file handle
Dim tmEnf As ENF                'ENF record image
Dim tmEnfSrchKey As INTKEY0     'ENF Key 0 image
Dim imEnfRecLen As Integer      'ENF record length
Dim imNoEnfSaved As Integer
Dim imEnfCode(0 To 20) As Integer   'Save last 20 used. Index zero ignored
Dim smEnfName(0 To 20) As String
'Comment (program code)
Dim hmCef As Integer    'Comment file handle
Dim tmCef As CEF        'CEF record image
Dim tmCefSrchKey As LONGKEY0    'CEF key record image
Dim imCefRecLen As Integer        'CEF record length
'Dim tmRec As LPOPREC
'Modular variables imported from Links
Dim imDelIndex As Integer       '0=Feed; 1=Vehicle
Dim imDateCode As Integer       'Date Code Active From Links module 0=M-F, 6=Sa, 7=Su
Dim smDateFilter As String      'Date filter from Links
Dim smEndDate As String         'TFN or end date
Dim imTFNDay As Integer         '0-4 from orig user date or 5= sat; 6=Sun
Dim imDate0 As Integer          'Byte 0 of smDateFilter
Dim imDate1 As Integer          'Byte 1 of smDateFilter
Dim imTermDate0 As Integer          'Byte 0 of smDateFilter-1
Dim imTermDate1 As Integer          'Byte 1 of smDateFilter-1
Dim imEndDate0 As Integer          'Byte 0 of smEndDate
Dim imEndDate1 As Integer          'Byte 1 of smEndDate
Dim imMnfFeed As Integer        'Feed code
Dim smVehName As String         'Vehicle name and type
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim imComboBoxIndex As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer   'True=Don't enable any box woth change
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer  'Shift state when mouse down event occurrs
Dim imBypassFocus As Integer
Dim imEvtNameIndex As Integer
Dim imVefCode As Integer    'Vehicle code #
Dim imVpfIndex As Integer   'Vehicle option index
Dim smEvtPrgName As String  'Program name- far coded to "Programs"
Dim smEvtAvName As String   'Contract Avail- far coded to"Contract Avails"
Dim imChg As Integer        'Any events changed
Dim imFirstTimeFocus As Integer
Dim smScreenCaption As String
Dim smScrollCaption As String
Dim imUpdateAllowed As Integer

'List Population arrays
Dim tmDEvt() As DELEVT             'Current Event image

Private bmFirstCallToVpfFind As Boolean

Const LBONE = 1

Const VEHICLEINDEX = 1      'Vehicle control/field
Const AIRTIMEINDEX = 2      'Air time control/field
Const LOCALTIMEINDEX = 3    'Local time control/field
Const FEEDTIMEINDEX = 4     'Feed time control/field
Const TIMEZONEINDEX = 5     'Time zone control/field
Const SUBFEEDINDEX = 6    'Subfeed control/field
Const EVTNAMEINDEX = 7    'Event name control/field
Const PROGCODEINDEX = 8     'Program code control/field
Const SHOWONINDEX = 9       'Show on cmml sch control/field
Const FEDINDEX = 10         'Fed
Const BUSINDEX = 11 ' 10         'Bus control/field
Const SCHINDEX = 12 '11         'Schedule control/field
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo, False
    imRowNo = -1
    imBoxNo = -1
    lacFrame(imDelIndex).Visible = False
    pbcArrow.Visible = False
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo, False
    imRowNo = -1
    imBoxNo = -1
    lacFrame(imDelIndex).Visible = False
    pbcArrow.Visible = False
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case VEHICLEINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
        Case LOCALTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case FEEDTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case TIMEZONEINDEX
        Case SUBFEEDINDEX
            lbcSubfeed.Visible = Not lbcSubfeed.Visible
        Case PROGCODEINDEX
        Case BUSINDEX
        Case SCHINDEX
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDupl_Click()
    Dim ilMax As Integer
    Dim ilLoop As Integer
    Dim ilEvt As Integer
    If Not mFindRowIndex(imRowNo) Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mSetShow imBoxNo, True
    ilMax = UBound(tmDlf) + 1
    ReDim Preserve tmDlf(0 To ilMax) As DLFLIST
    For ilLoop = ilMax - 1 To imDlfIndex + 1 Step -1
        tmDlf(ilLoop) = tmDlf(ilLoop - 1)
    Next ilLoop
    For ilEvt = lbcSortIndex.ListCount - 1 To 0 Step -1
        If lbcSortIndex.ItemData(ilEvt) >= imDlfIndex + 1 Then
           lbcSortIndex.ItemData(ilEvt) = lbcSortIndex.ItemData(ilEvt) + 1
        End If
    Next ilEvt
    tmDlf(imDlfIndex + 1).iStatus = 0
    lbcSortIndex.AddItem lbcSortIndex.List(imSortRowIndex), imSortRowIndex + 1
    lbcSortIndex.ItemData(imSortRowIndex + 1) = imDlfIndex + 1
    'mSort
    mCompMax
    imSvRowNo = -1
    If Not mFindRowIndex(imRowNo) Then
        Exit Sub
    End If
    imBoxNo = 0
    pbcDelivery(imDelIndex).Cls
    pbcDelivery_Paint imDelIndex
    pbcArrow.SetFocus
    Screen.MousePointer = vbDefault
    imChg = True
End Sub
Private Sub cmcDupl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub


Private Sub cmcPrefeed_Click()
    igPreFeedType = imDelOrEngr
    igPreFeedVefCode = imVefCode
    sgPreFeedDay = Trim$(Str$(imDateCode))
    sgPreFeedDate = smDateFilter
    sgPreFeedScreenCaption = smScreenCaption
    PreFeed.Show vbModal
End Sub

Private Sub cmcPrefeed_GotFocus()
    mSetShow imBoxNo, False
    imRowNo = -1
    imBoxNo = -1
    lacFrame(imDelIndex).Visible = False
    pbcArrow.Visible = False
End Sub

Private Sub cmcPrefeed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcUpdate_Click()
    Dim ilLoop As Integer
    Dim ilNoRemoved As Integer
    Dim ilIndex As Integer
    Dim ilMax As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    'Might want to delete old and unused records (iStatus = -1 or 2)
    ilMax = UBound(tmDlf) - 1
    ilNoRemoved = 0
    For ilLoop = ilMax To LBONE Step -1
        If (tmDlf(ilLoop).iStatus = -1) Or (tmDlf(ilLoop).iStatus = 2) Then
            For ilIndex = ilLoop To ilMax - 1 Step 1
                tmDlf(ilIndex) = tmDlf(ilIndex + 1)
            Next ilIndex
            ilNoRemoved = ilNoRemoved + 1
        End If
    Next ilLoop
    If ilNoRemoved <> 0 Then
        ilMax = ilMax - ilNoRemoved + 1
        ReDim Preserve tmDlf(0 To ilMax) As DLFLIST
        mSort
        mCompMax
        imSvRowNo = -1
    End If
    If vbcDelivery.Value <> vbcDelivery.Min Then
        vbcDelivery.Value = vbcDelivery.Min
    Else
        vbcDelivery_Change
    End If
    imChg = False
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcUpdate_GotFocus()
    mSetShow imBoxNo, False
    imRowNo = -1
    imBoxNo = -1
    lacFrame(imDelIndex).Visible = False
    pbcArrow.Visible = False
End Sub
Private Sub cmcUpdate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub edcDropDown_Change()
    Dim ilRet As Integer
    Dim slStr As String
    Select Case imBoxNo
        Case VEHICLEINDEX
            If imDelIndex = 0 Then
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
            Else
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcFeed, imBSMode, imComboBoxIndex
            End If
        Case LOCALTIMEINDEX
        Case FEEDTIMEINDEX
        Case TIMEZONEINDEX
        Case SUBFEEDINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcSubfeed, imBSMode, slStr)
            If ilRet = 1 Then
                lbcSubfeed.ListIndex = 0
            End If
        Case PROGCODEINDEX
        Case BUSINDEX
        Case SCHINDEX
    End Select
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case VEHICLEINDEX
            If imDelIndex = 0 Then
                If lbcVehicle.ListCount = 1 Then
                    lbcVehicle.ListIndex = 0
                    'If imTabDirection = -1 Then  'Right To Left
                    '    pbcSTab.SetFocus
                    'Else
                    '    pbcTab.SetFocus
                    'End If
                    'Exit Sub
                End If
            Else
                If lbcFeed.ListCount = 1 Then
                    lbcFeed.ListIndex = 0
                    'If imTabDirection = -1 Then  'Right To Left
                    '    pbcSTab.SetFocus
                    'Else
                    '    pbcTab.SetFocus
                    'End If
                    'Exit Sub
                End If
            End If
        Case LOCALTIMEINDEX
        Case FEEDTIMEINDEX
        Case TIMEZONEINDEX
        Case SUBFEEDINDEX
            If lbcSubfeed.ListCount = 1 Then
                lbcSubfeed.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case PROGCODEINDEX
        Case BUSINDEX
        Case SCHINDEX
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNo
        Case LOCALTIMEINDEX, FEEDTIMEINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                ilFound = False
                For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                    If KeyAscii = igLegalTime(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KEYDOWN) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcTab.SetFocus
        Exit Sub
    End If
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case VEHICLEINDEX
                If imDelIndex = 0 Then
                    gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
                Else
                    gProcessArrowKey Shift, KeyCode, lbcFeed, imLbcArrowSetting
                End If
            Case LOCALTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case FEEDTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case SUBFEEDINDEX
                gProcessArrowKey Shift, KeyCode, lbcSubfeed, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
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
'    gShowBranner
    If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    LinkDlvy.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        plcScroll.Visible = False
        plcScroll.Visible = True
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
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
    
    Erase tmNameCode
    Erase tmEvtTypeCode
    Erase tmSubFeedCode
    'Close Files
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmPff)
    btrDestroy hmPff
    ilRet = btrClose(hmDlf)
    btrDestroy hmDlf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmLvf)
    btrDestroy hmLvf
    ilRet = btrClose(hmLef)
    btrDestroy hmLef
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmEnf)
    btrDestroy hmEnf
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    Erase smShow
    Erase smSave
    Erase imSave
    Erase tmDlf
    Erase tmDEvt
    
    Set LinkDlvy = Nothing

End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub imcKey_Click()
    pbcKey.Visible = Not pbcKey.Visible
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pbcKey.Visible = True
End Sub
Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pbcKey.Visible = False
End Sub
Private Sub imcTrash_Click()
    Dim ilRowNo As Integer
    If (imRowNo < vbcDelivery.Value) Or (imRowNo >= vbcDelivery.Value + vbcDelivery.LargeChange + 1) Then
        Exit Sub
    End If
    ilRowNo = imRowNo
    mSetShow imBoxNo, False
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame(imDelIndex).Visible = False
    If Not mFindRowIndex(ilRowNo) Then
        Exit Sub
    End If
    gCtrlGotFocus ActiveControl
    'Change status
    If (tmDlf(imDlfIndex).iStatus = 0) Or (tmDlf(imDlfIndex).iStatus = 1) Then
        If tmDlf(imDlfIndex).iStatus = 1 Then
            tmDlf(imDlfIndex).iStatus = 2
        ElseIf tmDlf(imDlfIndex).iStatus = 0 Then
            tmDlf(imDlfIndex).iStatus = -1
        End If
        imSvRowNo = -1
        mCompMax
        imChg = True
    ElseIf tmDlf(imDlfIndex).iStatus = 2 Then
        tmDlf(imDlfIndex).iStatus = 1
        imSvRowNo = -1
        mCompMax
        imChg = True
    End If
    mSetCommands
    lacFrame(imDelIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcDelivery(imDelIndex).Cls
    pbcDelivery_Paint imDelIndex
End Sub
Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame(imDelIndex).DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacFrame(imDelIndex).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        lacFrame(imDelIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub lbcEvtName_Click(Index As Integer)
    gProcessLbcClick lbcEvtName(Index), edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcEvtName_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcEvtType_Click()
    gProcessLbcClick lbcEvtType, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcEvtType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcFeed_Click()
    gProcessLbcClick lbcFeed, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcSubfeed_Click()
    gProcessLbcClick lbcSubfeed, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcSubfeed_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehicle_Click()
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mBuildDlfRec                   *
'*                                                     *
'*             Created:7/27/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Dlf records images from   *
'*                     old dlf and events for day      *
'*                                                     *
'*******************************************************
Private Sub mBuildDlfRec()
    Dim ilLoop As Integer
    Dim ilVeh As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slDay As String * 1
    Dim ilUpperBound As Integer
    Dim ilStatus As Integer
    Dim ilDEvt As Integer
    Dim slTime As String
    Dim llTime As Long
    Dim llMatchTime As Long
    Dim ilMatchEtf As Integer
    Dim ilMatchEnf As Integer
    Dim ilExcluded As Integer
    Dim tlDlf As DLF
    ReDim tmDlf(0 To 1) As DLFLIST                'DLF record image

    imDlfRecLen = Len(tlDlf)  'Get and save DlF record length
    slDay = Trim$(Str$(imDateCode))
    For ilVeh = 0 To lbcVehCode.ListCount - 1 Step 1
        slNameCode = lbcVehCode.List(ilVeh)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        imVefCode = Val(slCode)
        tmVefSrchKey.iCode = imVefCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet = BTRV_ERR_NONE Then
            If bmFirstCallToVpfFind Then
                imVpfIndex = gVpfFind(LinkDlvy, imVefCode)
                bmFirstCallToVpfFind = False
            Else
                imVpfIndex = gVpfFindIndex(imVefCode)
            End If
            mReadLcf tmVef.sType
            ilDEvt = LBound(tmDEvt)
            llMatchTime = -1
            ilMatchEtf = -1
            ilMatchEnf = -1
            tmDlfSrchKey.iVefCode = imVefCode
            tmDlfSrchKey.sAirDay = slDay
            'tmDlfSrchKey.iStartDate(0) = 257  'Year 1/1/1900
            'tmDlfSrchKey.iStartDate(1) = 2100
            tmDlfSrchKey.iStartDate(0) = imDate0  'Year 1/1/1900
            tmDlfSrchKey.iStartDate(1) = imDate1
            tmDlfSrchKey.iAirTime(0) = 0
            tmDlfSrchKey.iAirTime(1) = 25 * 256 'Hour
            ilRet = btrGetLessOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If imDelIndex = 0 Then
                Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay) And (tlDlf.iMnfFeed <> imMnfFeed)
                    ilRet = btrGetPrevious(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Loop
            End If
            If (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay) And (((imDelIndex = 1) And (tlDlf.iMnfFeed > 0)) Or ((tlDlf.iMnfFeed = imMnfFeed) And (imDelIndex = 0))) Then
                'Start at earliest time and merge Lcf
                tmDlfSrchKey.iVefCode = imVefCode
                tmDlfSrchKey.sAirDay = slDay
                tmDlfSrchKey.iStartDate(0) = tlDlf.iStartDate(0)  'Year 1/1/1900
                tmDlfSrchKey.iStartDate(1) = tlDlf.iStartDate(1)
                tmDlfSrchKey.iAirTime(0) = 0
                tmDlfSrchKey.iAirTime(1) = 0
                ilRet = btrGetGreaterOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay) And ((imMnfFeed = -1) Or (tlDlf.iMnfFeed = imMnfFeed))
                    If ((tlDlf.iTermDate(0) = 0) And (tlDlf.iTermDate(1) = 0)) Or ((tlDlf.iTermDate(1) > imDate1) Or ((tlDlf.iTermDate(1) = imDate1) And (tlDlf.iTermDate(0) >= imDate0))) And ((imMnfFeed = -1) Or (tlDlf.iMnfFeed = imMnfFeed)) Then
                        'Test if time still exist or record should be deleted
                        ilStatus = 2    'Remove old or events that don't belong
                        ilExcluded = False
                        If imDelOrEngr = 0 Then 'all events- only version one
                            If (tlDlf.sCmmlSched = "N") And (tlDlf.iMnfSubFeed = 0) Then
                                ilExcluded = True
                            End If
                        Else    'Only avails
                            If (tlDlf.sFed = "N") And (tlDlf.iMnfSubFeed = 0) Then
                                ilExcluded = True
                            End If
                        End If
                        If (tlDlf.iStartDate(1) > imDate1) Or ((tlDlf.iStartDate(1) = imDate1) And (tlDlf.iStartDate(0) > imDate0)) Then
                            ilExcluded = True
                        End If
                        If Not ilExcluded Then
                            gUnpackTime tlDlf.iAirTime(0), tlDlf.iAirTime(1), "A", "1", slTime
                            llTime = CLng(gTimeToCurrency(slTime, True))
                            For ilDEvt = LBound(tmDEvt) To UBound(tmDEvt) - 1 Step 1
                                If (llTime = tmDEvt(ilDEvt).lTime) And (tlDlf.iEtfCode = tmDEvt(ilDEvt).iEtfCode) And (tlDlf.iEnfCode = tmDEvt(ilDEvt).iEnfCode) Then
                                    ilStatus = 1
                                    tmDEvt(ilDEvt).iStatus = 1  'Used
                                    'mResetDlfRec tmDEvt(ilDEvt), tlDlf
                                    Exit For
                                End If
                            Next ilDEvt
                        End If
                        'Do
                        '    If ilDEvt < UBound(tmDEvt) Then
                        '        gUnpackTime tlDlf.iAirTime(0), tlDlf.iAirTime(1), "A", "1", slTime
                        '        llTime = CLng(gTimeToCurrency(slTime, True))
                        '        If (llMatchTime = llTime) Then 'And (ilMatchEtf = tlDlf.iEtfCode) And (ilMatchEnf = tlDlf.iEnfCode) Then
                        '            'Scan forward and backward for all events with same time looking for match
                        '            ilTEvt = ilDEvt
                        '            Do While ilTEvt >= LBound(tmDEvt)
                        '                If tmDEvt(ilTEvt).lTime <> llTime Then
                        '                    Exit Do
                        '                End If
                        '                If (tlDlf.iEtfCode = tmDEvt(ilTEvt).iEtfCode) And (tlDlf.iEnfCode = tmDEvt(ilTEvt).iEnfCode) Then
                        '                    llMatchTime = llTime
                        '                    ilStatus = 1
                        '                    tmDEvt(ilTEvt).iStatus = 1  'Used
                        '                    Exit Do
                        '                End If
                        '                ilTEvt = ilTEvt - 1
                        '            Loop
                        '            If ilStatus = -1 Then
                        '                ilTEvt = ilDEvt + 1
                        '                Do While ilTEvt < UBound(tmDEvt)
                        '                    If tmDEvt(ilTEvt).lTime <> llTime Then
                        '                        Exit Do
                        '                    End If
                        '                    If (tlDlf.iEtfCode = tmDEvt(ilTEvt).iEtfCode) And (tlDlf.iEnfCode = tmDEvt(ilTEvt).iEnfCode) Then
                        '                        llMatchTime = llTime
                        '                        tmDEvt(ilTEvt).iStatus = 1  'Used
                        '                        ilStatus = 1
                        '                        Exit Do
                        '                    End If
                        '                    ilTEvt = ilTEvt + 1
                        '                Loop
                        '                If ilStatus = -1 Then
                        '                    ilStatus = 2
                        '                End If
                        '            End If
                        '        ElseIf tmDEvt(ilDEvt).lTime < llTime Then
                        '            If tmDEvt(ilDEvt).iStatus = 0 Then
                        '                mMakeDlfRec tmDEvt(ilDEvt)
                        '                tmDEvt(ilDEvt).iStatus = 1  'Used
                        '            End If
                        '            ilDEvt = ilDEvt + 1
                        '            llMatchTime = -1
                        '        ElseIf llTime < tmDEvt(ilDEvt).lTime Then
                        '            ilStatus = 2
                        '            llMatchTime = -1
                        '        Else    'Same time
                        '            llMatchTime = llTime    'Check it event same at top of loop
                        '        'ElseIf (tlDlf.iEtfCode = tmDEvt(ilDEvt).iEtfCode) And (tlDlf.iEnfCode = tmDEvt(ilDEvt).iEnfCode) Then
                        '        '    llMatchTime = llTime
                        '        '    ilMatchEtf = tlDlf.iEtfCode
                        '        '    ilMatchEnf = tlDlf.iEnfCode
                        '        '    ilStatus = 1
                        '        '    ilDEvt = ilDEvt + 1
                        '        'Else    'Same time but different events
                        '            'If ilDEvt < UBound(tmDEvt) Then
                        '            '    If (tmDEvt(ilDEvt + 1).lTime = llTime) And (tlDlf.iEtfCode = tmDEvt(ilDEvt + 1).iEtfCode) And (tlDlf.iEnfCode = tmDEvt(ilDEvt + 1).iEnfCode) Then
                        '            '        'User Added event prior to matching event
                        '            '        mMakeDlfRec tmDEvt(ilDEvt)
                        '            '        ilDEvt = ilDEvt + 1
                        '            '        llMatchTime = -1
                        '            '    Else
                        '            '        llMatchTime = -1
                        '            '        ilStatus = 2
                        '            '    End If
                        '            'Else
                        '            '    llMatchTime = -1
                        '            '    ilStatus = 2
                        '            'End If
                        '        End If
                        '    Else
                        '        'Test if match last event processed
                        '        'gUnpackTime tlDlf.iAirTime(0), tlDlf.iAirTime(1), "A", "1", slTime
                        '        'llTime = CLng(gTimeToCurrency(slTime, True))
                        '        'If (llMatchTime = llTime) And (ilMatchEtf = tlDlf.iEtfCode) And (ilMatchEnf = tlDlf.iEnfCode) Then
                        '        '    ilStatus = 1
                        '        'Else
                        '            ilStatus = 2
                        '            llMatchTime = -1
                        '            ilMatchEtf = 0
                        '            ilMatchEnf = 0
                        '        'End If
                        '    End If
                        'Loop While ilStatus = -1
                        ilUpperBound = UBound(tmDlf)
                        tmDlf(ilUpperBound).DlfRec = tlDlf
                        tmDlf(ilUpperBound).iStatus = ilStatus
                        'ilRet = btrGetPosition(hmDlf, tmDlf(ilUpperBound).lRecPos)
                        tmDlf(ilUpperBound).lDlfCode = tlDlf.lCode
                        mEvtString tmDlf(ilUpperBound)
                        ilUpperBound = ilUpperBound + 1
                        ReDim Preserve tmDlf(0 To ilUpperBound) As DLFLIST
                        If ilStatus = 2 Then    'Set change so update can be pressed without changing any other field
                            imChg = True
                        End If
                    End If
                    ilRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                'Add any not already defined by delivery links
                'If imDelOrEngr = 0 Then
                    For ilLoop = LBound(tmDEvt) To UBound(tmDEvt) - 1 Step 1
                        If tmDEvt(ilLoop).iStatus = 0 Then
                            mMakeDlfRec tmDEvt(ilLoop)
                            tmDEvt(ilLoop).iStatus = 1
                        End If
                    Next ilLoop
                'End If
            Else    'Build from Lcf
                'If imDelOrEngr = 0 Then
                    For ilLoop = LBound(tmDEvt) To UBound(tmDEvt) - 1 Step 1
                        mMakeDlfRec tmDEvt(ilLoop)
                        tmDEvt(ilLoop).iStatus = 1
                    Next ilLoop
                'End If
            End If
        End If
    Next ilVeh
    'Sort by time
    mSort
    vbcDelivery.Min = 1
    mCompMax
    If vbcDelivery.Value <> vbcDelivery.Min Then
        vbcDelivery.Value = vbcDelivery.Min
    Else
        vbcDelivery_Change
    End If
    mSetCommands
End Sub
Private Sub mCompMax()
    Dim ilMax As Integer
    Dim ilIndex As Integer
    ilMax = 0
    For ilIndex = LBONE To UBound(tmDlf) - 1 Step 1
        'If (tmDlf(ilIndex).iStatus = 0) Or (tmDlf(ilIndex).iStatus = 1) Then
        If (tmDlf(ilIndex).iStatus = 0) Or (tmDlf(ilIndex).iStatus = 1) Or (tmDlf(ilIndex).iStatus = 2) Then
            'If (tmDlf(ilIndex).DlfRec.sFed = "Y") Or (Not ckcFedEvtOnly.Value) Then
                ilMax = ilMax + 1
            'End If
        End If
    Next ilIndex
    If ilMax <= vbcDelivery.LargeChange + 1 Then
        vbcDelivery.Max = 1
    Else
        vbcDelivery.Max = ilMax - vbcDelivery.LargeChange '(ilMax - vbcLog1.Min) \ (vbcLog1.LargeChange + 1) + 1
    End If
    imMaxRowNo = ilMax
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mDeleteEvtNameCtrl             *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Reduce control array to 1       *
'*                     (index =0)                      *
'*                                                     *
'*******************************************************
Private Sub mDeleteEvtNameCtrl()
    Dim ilMaxNoCtrls As Integer
    Dim ilLoop As Integer
    ilMaxNoCtrls = 0
    On Error GoTo gDetMaxCtrlCntErr
    Do While (Err = 0)
        ilMaxNoCtrls = lbcEvtName(ilMaxNoCtrls).Index
        ilMaxNoCtrls = ilMaxNoCtrls + 1
    Loop
gDetMaxCtrlCntErr:
    On Error GoTo 0
    ilMaxNoCtrls = ilMaxNoCtrls - 1
    For ilLoop = ilMaxNoCtrls To 1 Step -1
        Unload lbcEvtName(ilLoop)
        Unload lbcEvtNameCode(ilLoop)
    Next ilLoop
    Exit Sub    'Require to avoid error as no resume was executed
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mDirection                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:process arrow keys              *
'*                                                     *
'*******************************************************
Private Sub mDirection(ilMoveDir As Integer)
'
'   mDirection ilMove
'   Where:
'       ilMove (I)- 0=Up; 1= down; 2= left; 3= right
'
    mSetShow imBoxNo, False
    Select Case ilMoveDir
        Case KEYUP  'Up
            If imRowNo > 1 Then
                imRowNo = imRowNo - 1
                If imRowNo < vbcDelivery.Value Then
                    imSettingValue = True
                    vbcDelivery.Value = vbcDelivery.Value - 1
                End If
            End If
        Case KEYDOWN  'Down
            If imRowNo < UBound(tmDlf) - 1 Then
                imRowNo = imRowNo + 1
                If imRowNo > vbcDelivery.Value + vbcDelivery.LargeChange Then
                    imSettingValue = True
                    vbcDelivery.Value = vbcDelivery.Value + 1
                End If
            Else
                imRowNo = 1
                imSettingValue = True
                vbcDelivery.Value = 1
            End If
        Case KEYLEFT  'Left
            If imBoxNo > VEHICLEINDEX Then
                imBoxNo = imBoxNo - 1
            Else
                imBoxNo = SCHINDEX
            End If
        Case KEYRIGHT  'Right
            If imBoxNo < SCHINDEX Then
                imBoxNo = imBoxNo + 1
            Else
                imBoxNo = VEHICLEINDEX
            End If
    End Select
    imSettingValue = False
    mEnableBox imBoxNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mEnableBox                     *
'*                                                     *
'*             Created:7/27/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Enable controls                 *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slChar As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If

    If (imRowNo < vbcDelivery.Value) Or (imRowNo >= vbcDelivery.Value + vbcDelivery.LargeChange + 1) Then
        mSetShow ilBoxNo, False
        Exit Sub
    End If

    If Not mFindRowIndex(imRowNo) Then
        pbcArrow.Visible = False
        lacFrame(imDelIndex).Visible = False
        Exit Sub
    End If
    lacFrame(imDelIndex).Move 0, tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) - 30
    lacFrame(imDelIndex).Visible = True
    pbcArrow.Move pbcArrow.Left, plcDelivery.Top + tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHICLEINDEX
            If imDelIndex = 0 Then
                lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 10)
                edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
                If tgSpf.iVehLen <= 40 Then
                    edcDropDown.MaxLength = tgSpf.iVehLen
                Else
                    edcDropDown.MaxLength = 20
                End If
                gMoveTableCtrl pbcDelivery(imDelIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                gFindMatch tmDlf(imDlfIndex).sVehicle, 0, lbcVehicle
                imChgMode = True
                If gLastFound(lbcVehicle) >= 0 Then
                    lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                    edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                Else
                    If imRowNo > 1 Then
                        If imPrevIndex <> -1 Then
                            gFindMatch tmDlf(imPrevIndex).sVehicle, 0, lbcVehicle
                            If gLastFound(lbcVehicle) >= 0 Then
                                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                                edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                            Else
                                If lbcVehicle.ListCount <= 0 Then
                                    lbcVehicle.ListIndex = -1
                                    edcDropDown.Text = ""
                                Else
                                    lbcVehicle.ListIndex = 0
                                    edcDropDown.Text = lbcVehicle.List(0)
                                End If
                            End If
                        Else
                            If lbcVehicle.ListCount <= 0 Then
                                lbcVehicle.ListIndex = -1
                                edcDropDown.Text = ""
                            Else
                                lbcVehicle.ListIndex = 0
                                edcDropDown.Text = lbcVehicle.List(0)
                            End If
                        End If
                    Else
                        If lbcVehicle.ListCount <= 0 Then
                            lbcVehicle.ListIndex = -1
                            edcDropDown.Text = ""
                        Else
                            lbcVehicle.ListIndex = 0
                            edcDropDown.Text = lbcVehicle.List(0)
                        End If
                    End If
                End If
                imComboBoxIndex = lbcVehicle.ListIndex
                imChgMode = False
                If edcDropDown.Top + edcDropDown.Height + lbcVehicle.Height < cmcDone.Top Then
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Else
                lbcFeed.Height = gListBoxHeight(lbcFeed.ListCount, 10)
                edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
                edcDropDown.MaxLength = 20
                gMoveTableCtrl pbcDelivery(imDelIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                gFindMatch tmDlf(imDlfIndex).sFeed, 0, lbcFeed
                imChgMode = True
                If gLastFound(lbcFeed) >= 0 Then
                    lbcFeed.ListIndex = gLastFound(lbcFeed)
                    edcDropDown.Text = lbcFeed.List(lbcFeed.ListIndex)
                Else
                    If imRowNo > 1 Then
                        If imPrevIndex <> -1 Then
                            gFindMatch tmDlf(imPrevIndex).sFeed, 0, lbcFeed
                            If gLastFound(lbcFeed) >= 0 Then
                                lbcFeed.ListIndex = gLastFound(lbcFeed)
                                edcDropDown.Text = lbcFeed.List(lbcFeed.ListIndex)
                            Else
                                If lbcFeed.ListCount <= 0 Then
                                    lbcFeed.ListIndex = -1
                                    edcDropDown.Text = ""
                                Else
                                    lbcFeed.ListIndex = 0
                                    edcDropDown.Text = lbcFeed.List(0)
                                End If
                            End If
                        Else
                            If lbcFeed.ListCount <= 0 Then
                                lbcFeed.ListIndex = -1
                                edcDropDown.Text = ""
                            Else
                                lbcFeed.ListIndex = 0
                                edcDropDown.Text = lbcFeed.List(0)
                            End If
                        End If
                    Else
                        If lbcFeed.ListCount <= 0 Then
                            lbcFeed.ListIndex = -1
                            edcDropDown.Text = ""
                        Else
                            lbcFeed.ListIndex = 0
                            edcDropDown.Text = lbcFeed.List(0)
                        End If
                    End If
                End If
                imComboBoxIndex = lbcFeed.ListIndex
                imChgMode = False
                If edcDropDown.Top + edcDropDown.Height + lbcFeed.Height < cmcDone.Top Then
                    lbcFeed.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                Else
                    lbcFeed.Move edcDropDown.Left, edcDropDown.Top - lbcFeed.Height
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            End If
        Case LOCALTIMEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + tmCtrls(ilBoxNo).fBoxW \ 3
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcDelivery(imDelIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            edcDropDown.Text = Trim$(tmDlf(imDlfIndex).sLocalTime)
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case FEEDTIMEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + tmCtrls(ilBoxNo).fBoxW \ 3
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcDelivery(imDelIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX - tmCtrls(ilBoxNo).fBoxW \ 4, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            edcDropDown.Text = Trim$(tmDlf(imDlfIndex).sFeedTime)
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TIMEZONEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcDelivery(imDelIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            edcDropDown.Text = Trim$(tmDlf(imDlfIndex).DlfRec.sZone)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case SUBFEEDINDEX 'Subfeed
            lbcSubfeed.Height = gListBoxHeight(lbcSubfeed.ListCount, 10)
            edcDropDown.Width = tmCtrls(SUBFEEDINDEX).fBoxW
            edcDropDown.MaxLength = 20
            gMoveTableCtrl pbcDelivery(imDelIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            imChgMode = True
            gFindMatch tmDlf(imDlfIndex).sSubfeed, 0, lbcSubfeed
            If gLastFound(lbcSubfeed) >= 0 Then
                lbcSubfeed.ListIndex = gLastFound(lbcSubfeed)
                edcDropDown.Text = lbcSubfeed.List(lbcSubfeed.ListIndex)
            Else
                If lbcSubfeed.ListCount > 0 Then
                    lbcSubfeed.ListIndex = 0
                    edcDropDown.Text = lbcSubfeed.List(lbcSubfeed.ListIndex)
                Else
                    lbcSubfeed.ListIndex = -1
                    edcDropDown.Text = ""
                End If
            End If
            imChgMode = False
            If edcDropDown.Top + edcDropDown.Height + lbcSubfeed.Height < cmcDone.Top Then
                lbcSubfeed.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcSubfeed.Move edcDropDown.Left, edcDropDown.Top - lbcSubfeed.Height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case EVTNAMEINDEX 'Event Name
            If tmDlf(imDlfIndex).DlfRec.iVefCode <= 0 Then
                imBoxNo = imBoxNo - 1
                pbcSTab.SetFocus
                Exit Sub
            End If
            imVefCode = tmDlf(imDlfIndex).DlfRec.iVefCode
            If bmFirstCallToVpfFind Then
                imVpfIndex = gVpfFind(LinkDlvy, imVefCode)
                bmFirstCallToVpfFind = False
            Else
                imVpfIndex = gVpfFindIndex(imVefCode)
            End If
            mInitEvtNamePop
'            mEvtNamePop imEvtNameIndex, lbcEvtName(imEvtNameIndex), lbcEvtNameCode(imEvtNameIndex)
            If imTerminate Then
                Exit Sub
            End If
            gFindMatch tmDlf(imDlfIndex).sEventType, 0, lbcEvtType
            If gLastFound(lbcEvtType) < 0 Then
                pbcSTab.SetFocus 'Go back to event type
                Exit Sub
            End If
            imEvtNameIndex = gLastFound(lbcEvtType) + 1
            lbcEvtName(imEvtNameIndex).Height = gListBoxHeight(lbcEvtName(imEvtNameIndex).ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 30
            gMoveTableCtrl pbcDelivery(imDelIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            imChgMode = True
            gFindMatch tmDlf(imDlfIndex).sEventName, 0, lbcEvtName(imEvtNameIndex)
            If gLastFound(lbcEvtName(imEvtNameIndex)) >= 0 Then
                lbcEvtName(imEvtNameIndex).ListIndex = gLastFound(lbcEvtName(imEvtNameIndex))
                edcDropDown.Text = lbcEvtName(imEvtNameIndex).List(lbcEvtName(imEvtNameIndex).ListIndex)
            Else
                If lbcEvtName(imEvtNameIndex).ListCount > 0 Then
                    lbcEvtName(imEvtNameIndex).ListIndex = 0
                    edcDropDown.Text = lbcEvtName(imEvtNameIndex).List(0)
                Else
                    lbcEvtName(imEvtNameIndex).ListIndex = -1
                    edcDropDown.Text = ""
                End If
            End If
            imChgMode = False
            If edcDropDown.Top + edcDropDown.Height + lbcEvtName(imEvtNameIndex).Height < cmcDone.Top Then
                lbcEvtName(imEvtNameIndex).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcEvtName(imEvtNameIndex).Move edcDropDown.Left, edcDropDown.Top - lbcEvtName(imEvtNameIndex).Height
            End If
            lbcEvtName(imEvtNameIndex).ZOrder vbBringToFront
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PROGCODEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 5
            gMoveTableCtrl pbcDelivery(imDelIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            edcDropDown.Text = Trim$(tmDlf(imDlfIndex).DlfRec.sProgCode)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case SHOWONINDEX
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveTableCtrl pbcDelivery(imDelIndex), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case FEDINDEX
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveTableCtrl pbcDelivery(imDelIndex), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case BUSINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 5
            gMoveTableCtrl pbcDelivery(imDelIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            If imDelOrEngr = 1 Then
                'Replace blanks with -
                slStr = ""
                For ilLoop = 1 To 5 Step 1
                    slChar = Mid$(tmDlf(imDlfIndex).DlfRec.sBus, ilLoop, 1)
                    If slChar = " " Then
                        slStr = slStr & "-"
                    Else
                        slStr = slStr & slChar
                    End If
                Next ilLoop
            Else
                slStr = Trim$(tmDlf(imDlfIndex).DlfRec.sBus)
            End If
            edcDropDown.Text = slStr    'tmDlf(imDlfIndex).DlfRec.sBus
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case SCHINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 2
            gMoveTableCtrl pbcDelivery(imDelIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            edcDropDown.Text = Trim$(tmDlf(imDlfIndex).DlfRec.sSchedule)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtNamePop                     *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: EvtNamePop the selection event *
'*                      name box                       *
'*                                                     *
'*******************************************************
Private Sub mEvtNamePop(ilEvtNameIndex As Integer, lbcName As Control, lbcNameCode As Control)
'
'   mEvtNamePop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slNameCode As String
    Dim ilLoop As Integer
    Dim slName As String
    Dim slCode As String
    ReDim ilfilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    If ilEvtNameIndex >= 0 Then 'Event index
        ilfilter(0) = INTEGERFILTER
        slFilter(0) = Trim$(Str$(imVefCode))
        ilOffSet(0) = gFieldOffset("Enf", "EnfVefCode") '2

        slNameCode = tmEvtTypeCode(ilEvtNameIndex).sKey    'lbcEvtTypeCode.List(ilEvtNameIndex)
        ilRet = gParseItem(slNameCode, 3, "\", slCode)
        On Error GoTo mEvtNamePopErr
        gCPErrorMsg ilRet, "mEvtNamePop (gParseItem field 3)", LinkDlvy
        On Error GoTo 0
        ilfilter(1) = INTEGERFILTER
        slFilter(1) = slCode
        ilOffSet(1) = gFieldOffset("Enf", "EnfEtfCode") '4
        ReDim tmNameCode(0 To 0) As SORTCODE
        smNameCodeTag = lbcNameCode.Tag
        For ilLoop = 0 To lbcNameCode.ListCount - 1 Step 1
            slName = lbcNameCode.List(ilLoop)
            gAddItemToSortCode slName, tmNameCode(), True
        Next ilLoop
        'ilRet = gIMoveListBox(LinkDlvy, lbcName, lbcNameCode, "Enf.Btr", gFieldOffset("Enf", "EnfName"), 30, ilFilter(), slFilter(), ilOffset())
        ilRet = gIMoveListBox(LinkDlvy, lbcName, tmNameCode(), smNameCodeTag, "Enf.Btr", gFieldOffset("Enf", "EnfName"), 30, ilfilter(), slFilter(), ilOffSet())
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mEvtNamePopErr
            gCPErrorMsg ilRet, "mEvtNamePop (gIMoveListBox)", LinkDlvy
            On Error GoTo 0
            lbcNameCode.Clear
            For ilLoop = 0 To UBound(tmNameCode) - 1 Step 1
                lbcNameCode.AddItem Trim$(tmNameCode(ilLoop).sKey), ilLoop
            Next ilLoop
            lbcNameCode.Tag = smNameCodeTag
        End If
    End If
    Exit Sub
mEvtNamePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mEvtString                     *
'*                                                     *
'*             Created:7/27/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:make strings within Dlf records *
'*                      images from events for day     *
'*                                                     *
'*******************************************************
Private Sub mEvtString(tlDlf As DLFLIST)
    Dim slTime As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    tlDlf.sVehicle = tmVef.sName
    gUnpackTime tlDlf.DlfRec.iAirTime(0), tlDlf.DlfRec.iAirTime(1), "A", "1", slTime
    tlDlf.sFeed = ""
    For ilLoop = 0 To lbcFeedCode.ListCount - 1 Step 1
        slNameCode = lbcFeedCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If tlDlf.DlfRec.iMnfFeed = Val(slCode) Then
            tlDlf.sFeed = Trim$(slName)
            Exit For
        End If
    Next ilLoop
    'Old value if invalid- replace with correct value
    If tlDlf.sFeed = "" Then
        If lbcFeedCode.ListCount = 1 Then
            slNameCode = lbcFeedCode.List(0)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tlDlf.sFeed = Trim$(slName)
            tlDlf.DlfRec.iMnfFeed = Val(slCode)
            imChg = True
        End If
    End If
    tlDlf.sAirTime = slTime
    gUnpackTime tlDlf.DlfRec.iLocalTime(0), tlDlf.DlfRec.iLocalTime(1), "A", "1", slTime
    tlDlf.sLocalTime = slTime
    gUnpackTime tlDlf.DlfRec.iFeedTime(0), tlDlf.DlfRec.iFeedTime(1), "A", "1", slTime
    tlDlf.sFeedTime = slTime
    tlDlf.sSubfeed = ""
    For ilLoop = 0 To UBound(tmSubFeedCode) - 1 Step 1 'lbcSubFeedCode.ListCount - 1 Step 1
        slNameCode = tmSubFeedCode(ilLoop).sKey 'lbcSubFeedCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If tlDlf.DlfRec.iMnfSubFeed = Val(slCode) Then
            tlDlf.sSubfeed = Trim$(slName)
            Exit For
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tmEvtTypeCode) - 1 Step 1  'lbcEvtTypeCode.ListCount - 1 Step 1
        slNameCode = tmEvtTypeCode(ilLoop).sKey    'lbcEvtTypeCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slName)
        ilRet = gParseItem(slNameCode, 3, "\", slCode)
        If tlDlf.DlfRec.iEtfCode = Val(slCode) Then
            tlDlf.sEventType = Trim$(slName)
            Exit For
        End If
    Next ilLoop
    If tlDlf.DlfRec.iEnfCode > 0 Then
        For ilLoop = 1 To imNoEnfSaved Step 1
            If imEnfCode(ilLoop) = tlDlf.DlfRec.iEnfCode Then
                tlDlf.sEventName = smEnfName(ilLoop)
                Exit Sub
            End If
        Next ilLoop
        tmEnfSrchKey.iCode = tlDlf.DlfRec.iEnfCode
        ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet = BTRV_ERR_NONE Then
            tlDlf.sEventName = Trim$(tmEnf.sName)
            For ilLoop = 19 To 1 Step -1
                imEnfCode(ilLoop + 1) = imEnfCode(ilLoop)
                smEnfName(ilLoop + 1) = smEnfName(ilLoop)
            Next ilLoop
            imEnfCode(1) = tlDlf.DlfRec.iEnfCode
            smEnfName(1) = tlDlf.sEventName
            If imNoEnfSaved < 20 Then
                imNoEnfSaved = imNoEnfSaved + 1
            End If
        Else
            tlDlf.sEventName = ""
        End If
    Else
        tlDlf.sEventName = ""
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mETypePop                       *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection event   *
'*                      type box                       *
'*                                                     *
'*******************************************************
Private Sub mEvtTypePop()
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcEvtType.ListIndex
    If ilIndex > 0 Then
        slName = lbcEvtType.List(ilIndex)
    End If
    'ilRet = gPopEvtNmByTypeBox(LinkDlvy, True, True, lbcEvtType, lbcEvtTypeCode)
    ilRet = gPopEvtNmByTypeBox(LinkDlvy, True, True, lbcEvtType, tmEvtTypeCode(), smEvtTypeCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mEvtTypePopErr
        gCPErrorMsg ilRet, "mEvtTypePop (gIMoveListBox: EvtType)", LinkDlvy
        On Error GoTo 0
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcEvtType
            If gLastFound(lbcEvtType) > 0 Then
                lbcEvtType.ListIndex = gLastFound(lbcEvtType)
            Else
                lbcEvtType.ListIndex = -1
            End If
        Else
            lbcEvtType.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mEvtTypePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Function mFindRowIndex(ilRowNo As Integer) As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim ilEvt As Integer
    Dim ilTestRowNo As Integer
    If ilRowNo = imSvRowNo Then
        mFindRowIndex = True
        Exit Function
    End If
    imSvRowNo = ilRowNo
    ilFound = False
    imPrevIndex = -1
    imDlfIndex = -1
    ilTestRowNo = 0
    For ilEvt = 0 To lbcSortIndex.ListCount - 1 Step 1
        'slNameCode = lbcSortIndex.List(ilEvt)
        'ilRet = gParseItem(slNameCode, 2, "\", slIndex)
        'ilIndex = Val(slIndex)
        ilIndex = lbcSortIndex.ItemData(ilEvt)
        'If (tmDlf(ilIndex).iStatus = 0) Or (tmDlf(ilIndex).iStatus = 1) Then
        If (tmDlf(ilIndex).iStatus = 0) Or (tmDlf(ilIndex).iStatus = 1) Or (tmDlf(ilIndex).iStatus = 2) Then
            'If (tmDlf(ilIndex).DlfRec.sFed = "Y") Or (Not ckcFedEvtOnly.Value) Then
                ilTestRowNo = ilTestRowNo + 1
                If ilTestRowNo = ilRowNo Then
                    imDlfIndex = ilIndex
                    imSortRowIndex = ilEvt
                    ilFound = True
                    Exit For
                Else
                    imPrevIndex = ilIndex
                End If
            'End If
        End If
    Next ilEvt
    mFindRowIndex = ilFound
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
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
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim ilZone As Integer
    Dim ilFound As Integer
    ReDim tmDlf(0 To 1) As DLFLIST                'DLF record image
    ''If ((Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName)) And (Len(Trim$(sgSpecialPassword)) = 4) Then
    'If ((Asc(tgSpf.sUsingFeatures8) And PREFEEDDEF) = PREFEEDDEF) Then      'yes, show comments on detail
    '    cmcPrefeed.Enabled = True
    'Else
    '    cmcPrefeed.Enabled = False
    'End If
    imFirstActivate = True
    bmFirstCallToVpfFind = True
    imLBCtrls = 1
    imTerminate = False
    imcKey.Picture = IconTraf!imcKey.Picture
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    imDelIndex = 1
    If Links!rbcLinks(1).Value Then
        imDelOrEngr = 0 'Delivery
        rbcSort(3).Enabled = False
        rbcSort(8).Enabled = False
        rbcSort(9).Enabled = False
        rbcSort(10).Enabled = False
    Else
        imDelOrEngr = 1
        rbcSort(2).Enabled = False
        rbcSort(8).Enabled = False
    End If
    mInitBox
    LinkDlvy.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone LinkDlvy
    'LinkDlvy.Show
    Screen.MousePointer = vbHourglass
    DoEvents
    ilRet = gVffRead()
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    imBypassFocus = False
    imSettingValue = False
    imBoxNo = -1 'Initialize current Box to N/A
    imRowNo = -1
    imChgMode = False
    imBSMode = False
    imChg = False
    imSvRowNo = -1
    imPrevIndex = -1
    imMaxRowNo = -1
    imDlfIndex = -1
    imDragType = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imChgMode = False
    imSettingValue = False
    imLbcArrowSetting = False
    imBSMode = False
    imFirstTimeFocus = True
    smEvtPrgName = "Programs"   'Used to test if event is a program
    smEvtAvName = "Contract Avails" 'Used to test if event is an avail
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    smDateFilter = Trim$(Links!edcStartDate.Text)   'Store Effective Date
    smEndDate = Trim$(Links!edcEndDate.Text)   'Store Effective Date
    If (Trim$(smEndDate) = "") Or (Trim$(smEndDate) = "TFN") Then
        imEndDate0 = 0
        imEndDate1 = 0
    Else
        gPackDate smEndDate, imEndDate0, imEndDate1
    End If
    If Links!rbcDay(0).Value Then     'M-F
        If gWeekDayStr(smDateFilter) <= 4 Then
            imTFNDay = gWeekDayStr(smDateFilter)
            smDateFilter = gObtainPrevMonday(smDateFilter)
        Else
            smDateFilter = gObtainNextMonday(smDateFilter)
            imTFNDay = gWeekDayStr(smDateFilter)
        End If
    ElseIf Links!rbcDay(1).Value Then 'Sa
        smDateFilter = gDecOneDay(gObtainNextSunday(smDateFilter))
        imTFNDay = gWeekDayStr(smDateFilter)
    Else                              'Su
        smDateFilter = gObtainNextSunday(smDateFilter)
        imTFNDay = gWeekDayStr(smDateFilter)
    End If
    gPackDate smDateFilter, imDate0, imDate1
    llDate = gDateValue(smDateFilter) - 1
    slDate = Format$(llDate, "m/d/yy")
    gPackDate slDate, imTermDate0, imTermDate1
    pbcDelivery(0).Visible = False
    pbcDelivery(1).Visible = True
    slNameCode = Trim$(tgVehCombo(Links!lbcACVeh.ListIndex).sKey)    'Links!lbcACVehCode.List(Links!lbcACVeh.ListIndex)
    ilRet = gParseItem(slNameCode, 1, "\", smVehName)
    ilRet = gParseItem(smVehName, 3, "|", smVehName)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    imVefCode = Val(slCode)
    imVefRecLen = Len(tmVef)  'Get and save MNF record length
    rbcSort(0).Caption = "Feed"
    slName = smVehName
    imMnfFeed = -1
    Screen.MousePointer = vbHourglass
    DoEvents
    If imDelOrEngr = 0 Then
        If Links!rbcDay(0).Value Then     'M-F
            imDateCode = 0
            smScreenCaption = "Delivery Link Definitions: Monday-Friday " & slName
        ElseIf Links!rbcDay(1).Value Then 'Sa
            imDateCode = 6
            smScreenCaption = "Delivery Link Definitions: Saturday " & slName
        Else                              'Su
            imDateCode = 7
            smScreenCaption = "Delivery Link Definitions: Sunday " & slName
        End If
    Else
        If Links!rbcDay(0).Value Then     'M-F
            imDateCode = 0
            smScreenCaption = "Engineering Link Definitions: Monday-Friday " & slName
        ElseIf Links!rbcDay(1).Value Then 'Sa
            imDateCode = 6
            smScreenCaption = "Engineering Link Definitions: Saturday " & slName
        Else                              'Su
            imDateCode = 7
            smScreenCaption = "Engineering Link Definitions: Sunday " & slName
        End If
    End If
    Screen.MousePointer = vbHourglass
    DoEvents
    imMnfRecLen = Len(tmMnf)  'Get and save MNF record length
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", LinkDlvy
    On Error GoTo 0
    imPffRecLen = Len(tmPFF)  'Get and save PFF record length
    hmPff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPff, "", sgDBPath & "Pff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pff.Btr)", LinkDlvy
    On Error GoTo 0
    If imDelIndex = 0 Then  'By Feed
        tmMnfSrchKey.iCode = imMnfFeed
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrGetEqual: Mnf.Btr)", LinkDlvy
        On Error GoTo 0
    End If
    If imDelOrEngr = 0 Then
        hmDlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hmDlf, "", sgDBPath & "Dlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Dlf.Btr)", LinkDlvy
        On Error GoTo 0
    Else
        hmDlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hmDlf, "", sgDBPath & "Egf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Egf.Btr)", LinkDlvy
        On Error GoTo 0
    End If
    imLcfRecLen = Len(tmLcf)  'Get and save LCF record length
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", LinkDlvy
    On Error GoTo 0
    imLvfRecLen = Len(tmLvf)  'Get and save LVF record length
    hmLvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lvf.Btr)", LinkDlvy
    On Error GoTo 0
    imLefRecLen = Len(tmLef)  'Get and save LEF record length
    hmLef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLef, "", sgDBPath & "Lef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lef.Btr)", LinkDlvy
    On Error GoTo 0
    imVefRecLen = Len(tmVef)  'Get and save Vlf record length
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", LinkDlvy
    On Error GoTo 0
    If imDelIndex = 1 Then
        tmVefSrchKey.iCode = imVefCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrGetEqual: Vef.Btr)", LinkDlvy
        On Error GoTo 0
    End If
    imEnfRecLen = Len(tmEnf)  'Get and save ENF record length
    hmEnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmEnf, "", sgDBPath & "Enf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Enf.Btr)", LinkDlvy
    On Error GoTo 0
    hmCef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cef.btr)", LinkDlvy
    On Error GoTo 0
    Screen.MousePointer = vbHourglass
    DoEvents
    lbcVehCode.Clear
    lbcVehicle.Clear
    lbcFeedCode.Clear
    lbcFeed.Clear
    If imDelIndex = 0 Then  'By feed
        For ilLoop = 0 To UBound(tgVehCombo) - 1 Step 1 'Links!lbcACVehCode.ListCount - 1 Step 1
            slNameCode = Trim$(tgVehCombo(ilLoop).sKey)    'Links!lbcACVehCode.List(ilLoop)
            lbcVehCode.AddItem slNameCode
        Next ilLoop
        For ilLoop = 0 To lbcVehCode.ListCount - 1 Step 1
            slNameCode = lbcVehCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            lbcVehicle.AddItem slName
        Next ilLoop
        lbcFeedCode.AddItem Trim$(tmMnf.sName) & "\" & Trim$(Str$(tmMnf.iCode))
        lbcFeed.AddItem Trim$(tmMnf.sName)
    Else    'by Vehicle
        slNameCode = Trim$(tgVehCombo(Links!lbcACVeh.ListIndex).sKey)    'Links!lbcACVehCode.List(Links!lbcACVeh.ListIndex)
        lbcVehCode.AddItem slNameCode
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        lbcVehicle.AddItem slName
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        imVefCode = Val(slCode)
        If bmFirstCallToVpfFind Then
            imVpfIndex = gVpfFind(LinkDlvy, imVefCode)
            bmFirstCallToVpfFind = False
        Else
            imVpfIndex = gVpfFindIndex(imVefCode)
        End If
        For ilZone = LBound(tgVpf(imVpfIndex).sGZone) To UBound(tgVpf(imVpfIndex).sGZone) Step 1
            If (Trim$(tgVpf(imVpfIndex).sGZone(ilZone)) <> "") And (tgVpf(imVpfIndex).iGMnfNCode(ilZone) > 0) Then
                ilFound = False
                For ilLoop = 0 To lbcFeedCode.ListCount - 1 Step 1
                    slNameCode = lbcFeedCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) = tgVpf(imVpfIndex).iGMnfNCode(ilZone) Then
                        ilFound = True
                        'This is Ok as long as the vehicle is only via one feed
                        If tmMnf.iCode <> Val(slCode) Then
                            tmMnfSrchKey.iCode = Val(slCode)
                            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                        End If
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    tmMnfSrchKey.iCode = tgVpf(imVpfIndex).iGMnfNCode(ilZone)
                    ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                    On Error GoTo mInitErr
                    gBtrvErrorMsg ilRet, "mInit (btrGetEqual: Mnf.Btr)", LinkDlvy
                    On Error GoTo 0
                    slName = Trim$(tmMnf.sName) & "\" & Trim$(Str$(tgVpf(imVpfIndex).iGMnfNCode(ilZone)))
                    lbcFeedCode.AddItem slName
                End If
            End If
        Next ilZone
        For ilLoop = 0 To lbcFeedCode.ListCount - 1 Step 1
            slNameCode = lbcFeedCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            lbcFeed.AddItem slName
        Next ilLoop
    End If
    Screen.MousePointer = vbHourglass
    DoEvents
    If (imDelIndex = 1) And ((Asc(tgSpf.sUsingFeatures8) And PREFEEDDEF) = PREFEEDDEF) Then
        'Look for the opposite Definition
        If imDelOrEngr = 0 Then
            tmPffSrchKey1.sType = "D"
        Else
            tmPffSrchKey1.sType = "E"
        End If
        tmPffSrchKey1.iVefCode = imVefCode
        tmPffSrchKey1.sAirDay = Trim$(Str$(imDateCode))
        gPackDate "12/31/2069", tmPffSrchKey1.iStartDate(0), tmPffSrchKey1.iStartDate(1)
        ilRet = btrGetLessOrEqual(hmPff, tmPFF, imPffRecLen, tmPffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If (ilRet = BTRV_ERR_NONE) And (tmPFF.sType = tmPffSrchKey1.sType) And (tmPFF.iVefCode = imVefCode) Then
            cmcPrefeed.Enabled = True
        Else
            If imDelOrEngr = 0 Then
                tmPffSrchKey1.sType = "E"   '"D"
            Else
                tmPffSrchKey1.sType = "D"   '"E"
            End If
            tmPffSrchKey1.iVefCode = imVefCode
            tmPffSrchKey1.sAirDay = Trim$(Str$(imDateCode))
            gPackDate "12/31/2069", tmPffSrchKey1.iStartDate(0), tmPffSrchKey1.iStartDate(1)
            ilRet = btrGetLessOrEqual(hmPff, tmPFF, imPffRecLen, tmPffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If (ilRet = BTRV_ERR_NONE) And (tmPFF.sType = tmPffSrchKey1.sType) And (tmPFF.iVefCode = imVefCode) Then
                cmcPrefeed.Enabled = False
            Else
                cmcPrefeed.Enabled = True
            End If
        End If
    Else
        cmcPrefeed.Enabled = False
    End If
    'lbcEvtTypeCode.Clear
    ReDim tmEvtTypeCode(0 To 0) As SORTCODE
    smEvtTypeCodeTag = ""
    mEvtTypePop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    lbcSubfeed.Clear
    If tmMnf.iGroupNo = 1 Then
        mSubfeedPop
        If imTerminate Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        ReDim tmSubFeedCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    End If
    Screen.MousePointer = vbHourglass
    For ilLoop = LBound(imSort) To UBound(imSort) Step 1
        imSort(ilLoop) = -1
    Next ilLoop
    If imDelIndex = 0 Then
        imSort(LBound(imSort)) = 0  'Vehicle
        imSort(LBound(imSort) + 1) = 3  'Feed
        imSort(LBound(imSort) + 2) = 1  '
    Else
        imSort(LBound(imSort)) = 4  'Zone
        imSort(LBound(imSort) + 1) = 3  'Feed
        rbcSort(1).Value = True
        'imSort(LBound(imSort) + 2) = 3  'Feed
    End If
    Screen.MousePointer = vbHourglass
    DoEvents
    imNoEnfSaved = 0
    mBuildDlfRec
    plcScreen_Paint
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                      *
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
    flTextHeight = pbcDelivery(0).TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcDelivery.Move 210, 225, pbcDelivery(imDelIndex).Width + vbcDelivery.Width + fgPanelAdj, pbcDelivery(imDelIndex).Height + fgPanelAdj
    pbcDelivery(imDelIndex).Move plcDelivery.Left + fgBevelX, plcDelivery.Top + fgBevelY
    vbcDelivery.Move pbcDelivery(imDelIndex).Left + pbcDelivery(imDelIndex).Width + 15, pbcDelivery(imDelIndex).Top, vbcDelivery.Width, pbcDelivery(imDelIndex).Height
    pbcArrow.Move plcDelivery.Left - pbcArrow.Width - 15    'set arrow    'Vehicle
    'Vehicle
    gSetCtrl tmCtrls(VEHICLEINDEX), 30, 375, 1185, fgBoxGridH
    'Air time
    gSetCtrl tmCtrls(AIRTIMEINDEX), 1230, tmCtrls(VEHICLEINDEX).fBoxY, 780, fgBoxGridH
    'Local time
    gSetCtrl tmCtrls(LOCALTIMEINDEX), 2025, tmCtrls(VEHICLEINDEX).fBoxY, 780, fgBoxGridH
    'Feed time
    gSetCtrl tmCtrls(FEEDTIMEINDEX), 2820, tmCtrls(VEHICLEINDEX).fBoxY, 780, fgBoxGridH
'    tmCtrls(AVAILINDEX).iReq = False
    'Time zone
    gSetCtrl tmCtrls(TIMEZONEINDEX), 3615, tmCtrls(VEHICLEINDEX).fBoxY, 390, fgBoxGridH
    'Event type
    gSetCtrl tmCtrls(SUBFEEDINDEX), 4020, tmCtrls(VEHICLEINDEX).fBoxY, 1185, fgBoxGridH
    'Event name
    gSetCtrl tmCtrls(EVTNAMEINDEX), 5220, tmCtrls(VEHICLEINDEX).fBoxY, 1185, fgBoxGridH
    'Program code
    gSetCtrl tmCtrls(PROGCODEINDEX), 6420, tmCtrls(VEHICLEINDEX).fBoxY, 645, fgBoxGridH
    'Show on cmml sch
    gSetCtrl tmCtrls(SHOWONINDEX), 7080, tmCtrls(VEHICLEINDEX).fBoxY, 450, fgBoxGridH
    'Fed
    gSetCtrl tmCtrls(FEDINDEX), 7545, tmCtrls(VEHICLEINDEX).fBoxY, 315, fgBoxGridH
    'Bus
    gSetCtrl tmCtrls(BUSINDEX), 7875, tmCtrls(VEHICLEINDEX).fBoxY, 615, fgBoxGridH
    'Schedule
    gSetCtrl tmCtrls(SCHINDEX), 8460, tmCtrls(VEHICLEINDEX).fBoxY, 315, fgBoxGridH
End Sub
Private Sub mInitEvtNamePop()
    Dim ilLoop As Integer
    'Test if different vehicle
    If Val(lbcEvtNameCode(0).Tag) = imVefCode Then
        Exit Sub
    End If
    lbcEvtNameCode(0).Tag = Trim$(Str$(imVefCode))
    mDeleteEvtNameCtrl
    For ilLoop = 0 To lbcEvtType.ListCount - 1 Step 1
        Load lbcEvtName(ilLoop + 1) 'Create list box
        Load lbcEvtNameCode(ilLoop + 1)
        lbcEvtName(ilLoop + 1).Clear
        mEvtNamePop ilLoop, lbcEvtName(ilLoop + 1), lbcEvtNameCode(ilLoop + 1)
        If imTerminate Then
            Exit Sub
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mMakeDlfRec                    *
'*                                                     *
'*             Created:7/27/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:make Dlf records images from    *
'*                     events for day                  *
'*                                                     *
'*******************************************************
Private Sub mMakeDlfRec(tlDEvt As DELEVT)
    Dim ilUpperBound As Integer
    Dim clTime As Currency
    Dim slTime As String
    Dim ilZone As Integer
    Dim ilTimeAdj As Integer
    Dim ilDispl As Integer
    Dim ilCreate As Integer
    Dim ilNumVar As Integer
    Dim ilVff As Integer

    If imDelOrEngr = 0 Then 'Delivery- only create one version for all events types
        ilNumVar = 1
    Else    'Engineering-  only create avails (all versions)
        If (tlDEvt.iEtfCode >= 2) And (tlDEvt.iEtfCode <= 9) Then  'Avail
            ilNumVar = 4    'Create all versions
        Else
            Exit Sub   'Ignore all event types except avails
        End If
    End If
    ilUpperBound = UBound(tmDlf)
    For ilZone = LBound(tgVpf(imVpfIndex).sGZone) To UBound(tgVpf(imVpfIndex).sGZone) Step 1
        If (Trim$(tgVpf(imVpfIndex).sGZone(ilZone)) <> "") And (((tgVpf(imVpfIndex).iGMnfNCode(ilZone) > 0) And (imDelIndex = 1)) Or ((tgVpf(imVpfIndex).iGMnfNCode(ilZone) = imMnfFeed) And (imDelIndex = 0))) Then
            For ilDispl = 1 To ilNumVar Step 1
                ilCreate = False
                Select Case ilDispl
                    Case 1  'Primary
                        ilTimeAdj = 60 * tgVpf(imVpfIndex).iGV1Z(ilZone)
                        ilCreate = True
                    Case 2
                        If tgVpf(imVpfIndex).iGV2Z(ilZone) <> 0 Then
                            ilTimeAdj = 60 * tgVpf(imVpfIndex).iGV2Z(ilZone)
                            ilCreate = True
                        End If
                    Case 3
                        If tgVpf(imVpfIndex).iGV3Z(ilZone) <> 0 Then
                            ilTimeAdj = 60 * tgVpf(imVpfIndex).iGV3Z(ilZone)
                            ilCreate = True
                        End If
                    Case 4
                        If tgVpf(imVpfIndex).iGV4Z(ilZone) <> 0 Then
                            ilTimeAdj = 60 * tgVpf(imVpfIndex).iGV4Z(ilZone)
                            ilCreate = True
                        End If
                End Select
                If ilCreate Then
                    tmDlf(ilUpperBound).DlfRec.iVefCode = imVefCode
                    tmDlf(ilUpperBound).DlfRec.sAirDay = Trim$(Str$(imDateCode))
                    clTime = tlDEvt.lTime
                    slTime = gCurrencyToTime(clTime)
                    gPackTime slTime, tmDlf(ilUpperBound).DlfRec.iAirTime(0), tmDlf(ilUpperBound).DlfRec.iAirTime(1)
                    clTime = tlDEvt.lTime + 3600 * tgVpf(imVpfIndex).iGLocalAdj(ilZone) + ilTimeAdj
                    slTime = gCurrencyToTime(clTime)
                    gPackTime slTime, tmDlf(ilUpperBound).DlfRec.iLocalTime(0), tmDlf(ilUpperBound).DlfRec.iLocalTime(1)
                    clTime = tlDEvt.lTime + 3600 * tgVpf(imVpfIndex).iGFeedAdj(ilZone) + ilTimeAdj
                    slTime = gCurrencyToTime(clTime)
                    gPackTime slTime, tmDlf(ilUpperBound).DlfRec.iFeedTime(0), tmDlf(ilUpperBound).DlfRec.iFeedTime(1)
                    tmDlf(ilUpperBound).DlfRec.sZone = tgVpf(imVpfIndex).sGZone(ilZone)
                    tmDlf(ilUpperBound).DlfRec.iEtfCode = tlDEvt.iEtfCode
                    tmDlf(ilUpperBound).DlfRec.iEnfCode = tlDEvt.iEnfCode
                    'Scan backwards for matching Vehicle, Local Time, and time zone- if found
                    'use its sProgCode
'                    For ilLoop = ilUpperBound - 1 To LBound(tmDlf) Step -1
'                        If (tmDlf(ilLoop).DlfRec.iVefCode = tmDlf(ilUpperBound).DlfRec.iVefCode) And (tmDlf(ilLoop).DlfRec.sZone = tmDlf(ilUpperBound).DlfRec.sZone) And (tmDlf(ilLoop).DlfRec.iLocalTime(0) = tmDlf(ilUpperBound).DlfRec.iLocalTime(0)) And (tmDlf(ilLoop).DlfRec.iLocalTime(1) = tmDlf(ilUpperBound).DlfRec.iLocalTime(1)) Then
'                            tlDEvt.sProgCode = tmDlf(ilLoop).DlfRec.sProgCode
'                            Exit For
'                        End If
'                    Next ilLoop
                    If imDelOrEngr = 0 Then
                        tmDlf(ilUpperBound).DlfRec.sProgCode = tlDEvt.sProgCode
                    Else
                        tmDlf(ilUpperBound).DlfRec.sProgCode = ""
                    End If
                    If imDelOrEngr = 0 Then
                        tmDlf(ilUpperBound).DlfRec.sCmmlSched = "Y"
                    Else
                        tmDlf(ilUpperBound).DlfRec.sCmmlSched = "N"
                    End If
                    'If tgVpf(imVpfIndex).sGCSVer(ilZone) = "A" Then
                    '    tmDlf(ilUpperBound).DlfRec.sCmmlSched = "Y"
                    'Else
                    '    If ilDispl = 1 Then
                    '        tmDlf(ilUpperBound).DlfRec.sCmmlSched = "Y"
                    '    Else
                    '        tmDlf(ilUpperBound).DlfRec.sCmmlSched = "N"
                    '    End If
                    'End If
                    tmDlf(ilUpperBound).DlfRec.iMnfFeed = tgVpf(imVpfIndex).iGMnfNCode(ilZone)
                    tmDlf(ilUpperBound).DlfRec.sBus = tgVpf(imVpfIndex).sGBus(ilZone)
                    tmDlf(ilUpperBound).DlfRec.sSchedule = tgVpf(imVpfIndex).sGSked(ilZone)
                    tmDlf(ilUpperBound).DlfRec.iStartDate(0) = imDate0
                    tmDlf(ilUpperBound).DlfRec.iStartDate(1) = imDate1
                    tmDlf(ilUpperBound).DlfRec.iTermDate(0) = 0
                    tmDlf(ilUpperBound).DlfRec.iTermDate(1) = 0
                    tmDlf(ilUpperBound).DlfRec.iMnfSubFeed = 0
                    If imDelOrEngr = 0 Then
                        'tmDlf(ilUpperBound).DlfRec.sFed = "N"
                        If (tlDEvt.iEtfCode = 1) Or (tlDEvt.iEtfCode > 13) Then  'Program event type are always set to No
                            tmDlf(ilUpperBound).DlfRec.sFed = "N"
                        Else
                            '5/11/12: Moved vpf.sGFed for delivery to vff
                            'tmDlf(ilUpperBound).DlfRec.sFed = tgVpf(imVpfIndex).sGFed(ilZone)
                            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                                If imVefCode = tgVff(ilVff).iVefCode Then
                                    tmDlf(ilUpperBound).DlfRec.sFed = tgVff(ilVff).sFedDelivery(ilZone)
                                    Exit For
                                End If
                            Next ilVff
                        End If
                    Else
                        tmDlf(ilUpperBound).DlfRec.sFed = "Y"
                    End If
                    'If (tlDEvt.iEtfCode = 1) Or (tlDEvt.iEtfCode > 13) Then  'Program event type are always set to No
                    '    tmDlf(ilUpperBound).DlfRec.sFed = "N"
                    'Else
                    '    tmDlf(ilUpperBound).DlfRec.sFed = tgVpf(imVpfIndex).sGFed(ilZone)
                    'End If
                    tmDlf(ilUpperBound).lDlfCode = 0
                    tmDlf(ilUpperBound).iStatus = 0
                    mEvtString tmDlf(ilUpperBound)
                    ilUpperBound = ilUpperBound + 1
                    ReDim Preserve tmDlf(0 To ilUpperBound) As DLFLIST
                    imChg = True
                End If
            Next ilDispl
        End If
    Next ilZone
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadCefRec                     *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified: 4/24/94      By:D. Hannifan    *
'*                                                     *
'*            Comments: Read in comment record         *
'*                                                     *
'*******************************************************
Private Function mReadCefRec(llCefCode As Long) As Integer
'
'   iRet = mReadCefRec()
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    tmCefSrchKey.lCode = llCefCode
    If llCefCode <> 0 Then
        tmCef.sComment = ""
        imCefRecLen = Len(tmCef)    '1009
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mReadCefRecErr
        gBtrvErrorMsg ilRet, "mReadCefRec (btrGetEqual:Comment)", LinkDlvy
        On Error GoTo 0
    Else
        tmCef.lCode = 0
        'tmCef.iStrLen = 0
        tmCef.sComment = ""
    End If
    mReadCefRec = True
    Exit Function
mReadCefRecErr:
    On Error GoTo 0
    tmCef.lCode = 0
    'tmCef.iStrLen = 0
    tmCef.sComment = ""
    mReadCefRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadLcf                        *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified: 4/24/94      By:D. Hannifan    *
'*                                                     *
'*            Comments: Read in all events for a date  *
'*                                                     *
'*******************************************************
Private Sub mReadLcf(slCAType As String)
'
'   slCAType(I)- Vehicle type "C"=Conventional; "A"=airing
'
'   tmDEvt (I/O)-contain the log calendar events
'   imVefCode (I)-Vehicle
'   smDateFilter contains the effective date
'


    Dim ilUpper As Integer          'Upperbound of tmDEvt array
    Dim ilSeqNo As Integer          'Sequence number
    Dim ilRet As Integer            'Return from call
    Dim ilIndex As Integer          'List index
    Dim slStartTime As String       'Effective start time
    Dim slStr As String             'Parse string
    Dim slTime As String
    Dim ilDate0 As Integer          'Byte 0 start date
    Dim ilDate1 As Integer          'Byte 1 start date
    ReDim tmDEvt(0 To 0) As DELEVT      'image
    Dim ilFound As Integer          'True=valid avail found
    Dim ilDay As Integer
    Dim slDate As String
    Dim ilType As Integer
    Dim slComment As String
    Dim slXMid As String
    On Error GoTo mReadLcfErr

    ilUpper = UBound(tmDEvt)
    ilType = 0
    ilSeqNo = 1
    gPackDate smDateFilter, ilDate0, ilDate1
    ilDay = gWeekDayStr(smDateFilter)
    If (slCAType = "A") Or (slCAType = "C") Then    'Determine effective date
        ilFound = False
        tmLcfSrchKey.iType = ilType    'On air
        tmLcfSrchKey.sStatus = "C"  'Current
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = ilDate0
        tmLcfSrchKey.iLogDate(1) = ilDate1
        tmLcfSrchKey.iSeqNo = ilSeqNo
        ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
        Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = imVefCode) And (tmLcf.sStatus = "C")
            gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slDate
            If ilDay <= 4 Then  'Test for Only partial week defined
                If (gWeekDayStr(slDate) >= 0) And (gWeekDayStr(slDate) <= 4) Then
                    ilDate0 = tmLcf.iLogDate(0)
                    ilDate1 = tmLcf.iLogDate(1)
                    ilFound = True
                    Exit Do
                End If
            Else    'Sat or Sun
                If ilDay = gWeekDayStr(slDate) Then
                    ilDate0 = tmLcf.iLogDate(0)
                    ilDate1 = tmLcf.iLogDate(1)
                    ilFound = True
                    Exit Do
                End If
            End If
            ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If Not ilFound Then
            'Use TFN
            tmLcfSrchKey.iType = ilType    'On air
            tmLcfSrchKey.sStatus = "C"  'Current
            tmLcfSrchKey.iVefCode = imVefCode
            tmLcfSrchKey.iLogDate(0) = imTFNDay + 1
            tmLcfSrchKey.iLogDate(1) = 0
            tmLcfSrchKey.iSeqNo = ilSeqNo
            ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
            If (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = imVefCode) And (tmLcf.sStatus = "C") Then
                If (tmLcf.iLogDate(0) <= 7) And (tmLcf.iLogDate(1) = 0) Then
                    If imTFNDay + 1 = tmLcf.iLogDate(0) Then
                        ilDate0 = tmLcf.iLogDate(0)
                        ilDate1 = tmLcf.iLogDate(1)
                        ilFound = True
                    End If
                End If
            End If
        End If
        If Not ilFound Then
            Exit Sub
        End If
    End If
    Do
        tmLcfSrchKey.iType = ilType
        tmLcfSrchKey.sStatus = "C"
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = ilDate0
        tmLcfSrchKey.iLogDate(1) = ilDate1
        tmLcfSrchKey.iSeqNo = ilSeqNo
        ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet = BTRV_ERR_NONE Then
            ilSeqNo = ilSeqNo + 1
            For ilIndex = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
                If tmLcf.lLvfCode(ilIndex) <> 0 Then
                    gUnpackTime tmLcf.iTime(0, ilIndex), tmLcf.iTime(1, ilIndex), "A", "1", slStartTime
                    'Read in Lnf to obtain name and length
                    tmLvfSrchKey.lCode = tmLcf.lLvfCode(ilIndex)
                    ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'Get current record
                    If ilRet = BTRV_ERR_NONE Then
                        'Read in all the event record (Lef)
                        tmLefSrchKey.lLvfCode = tmLcf.lLvfCode(ilIndex)
                        tmLefSrchKey.iStartTime(0) = 0
                        tmLefSrchKey.iStartTime(1) = 0
                        tmLefSrchKey.iSeqNo = 0
                        ilRet = btrGetGreaterOrEqual(hmLef, tmLef, imLefRecLen, tmLefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmLef.lLvfCode = tmLcf.lLvfCode(ilIndex))
                            gUnpackLength tmLef.iStartTime(0), tmLef.iStartTime(1), "3", False, slStr
                            gAddTimeLength slStartTime, slStr, "A", "1", slTime, slXMid
                            tmDEvt(ilUpper).lTime = CLng(gTimeToCurrency(slTime, True))
                            tmDEvt(ilUpper).iEtfCode = tmLef.iEtfCode
                            tmDEvt(ilUpper).iEnfCode = tmLef.iEnfCode
                            tmDEvt(ilUpper).iStatus = 0 'Unused
                            ilFound = False
                            Select Case tmLef.iEtfCode
                                Case 1  'Program
                                    If mReadCefRec(tmLef.lCefCode) Then
                                        'If tmCef.iStrLen > 0 Then
                                        '    slComment = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                                        'Else
                                        '    slComment = ""
                                        'End If
                                        slComment = gStripChr0(tmCef.sComment)
                                    Else
                                        slComment = ""
                                    End If
                                    If imDelOrEngr = 0 Then
                                        ilFound = True
                                    End If
                                Case 2  'Contract Avail
                                    'Use avail comment for progcode if defined
                                    If mReadCefRec(tmLef.lCefCode) Then
                                        'If tmCef.iStrLen > 0 Then
                                        '    slComment = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                                        'End If
                                        slComment = gStripChr0(tmCef.sComment)
                                    End If
                                    ilFound = True
                                Case 3
                                    ilFound = True
                                Case 4
                                    ilFound = True
                                Case 5
                                    ilFound = True
                                Case 6  'Cmml Promo
                                    ilFound = True
                                Case 7  'Feed avail
                                    ilFound = True
                                Case 8  'PSA/Promo (Avail)
                                    ilFound = True
                                Case 9
                                    ilFound = True
                                Case 10  'Page eject, Line space 1, 2 or 3
                                Case 11
                                Case 12
                                Case 13
                                Case Else   'Other
                                    If imDelOrEngr = 0 Then
                                        ilFound = True
                                    End If
                            End Select
                            If ilFound Then
                                tmDEvt(ilUpper).sProgCode = slComment   'Use program comment on all events as prog code
                                ilUpper = ilUpper + 1
                                ReDim Preserve tmDEvt(0 To ilUpper) As DELEVT
                            End If
                            ilRet = btrGetNext(hmLef, tmLef, imLefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
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
Exit Sub
mReadLcfErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mSaveRec                       *
'*                                                     *
'*             Created:7/27/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Terminate records and save new  *
'*                     ones                            *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilIndex As Integer   'For loop control
    Dim ilRowNo As Integer
    Dim ilEvt As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim tlDlf As DLF
    Dim llDlfCode As Long
    Dim slMsg As String
    Dim slDate As String
    Dim slNowDate As String
    Dim llNowDate As Long
    Dim slDay As String
    Dim ilVeh As Integer
    Dim ilIgnoreRec As Integer
    Screen.MousePointer = vbHourglass  'Wait
    mSetShow imBoxNo, False
    imBoxNo = -1
    imRowNo = -1
    ilRowNo = 0
    'Check that all required fields are answered
    For ilEvt = 0 To lbcSortIndex.ListCount - 1 Step 1
        'slNameCode = lbcSortIndex.List(ilEvt)
        'ilRet = gParseItem(slNameCode, 2, "\", slIndex)
        'ilIndex = Val(slIndex)
        ilIndex = lbcSortIndex.ItemData(ilEvt)
        If (tmDlf(ilIndex).iStatus = 0) Or (tmDlf(ilIndex).iStatus = 1) Then
            ilRowNo = ilRowNo + 1
            If mTestFields(ilIndex) = NO Then
                'ckcFedEvtOnly.Value = False
                'Position to row in error
                imSettingValue = True
                If ilRowNo < vbcDelivery.Max Then
                    vbcDelivery.Value = ilRowNo 'Make error row the top row
                Else
                    vbcDelivery.Value = vbcDelivery.Max
                End If
                imSettingValue = False
                imRowNo = ilRowNo
                mSaveRec = False
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
    Next ilEvt
    '
    'Remove records that match vehicle and day and start date is after
    'terminate date
    'Remove all teminated records in the past
    slNowDate = Format$(Now, "m/d/yy")
    llNowDate = gDateValue(slNowDate)
    imDlfRecLen = Len(tlDlf)  'Get and save DlF record length
    slDay = Trim$(Str$(imDateCode))
    ilRet = btrBeginTrans(hmDlf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Links")
        mSaveRec = False
        Exit Function
    End If
    For ilVeh = 0 To lbcVehCode.ListCount - 1 Step 1
        slNameCode = lbcVehCode.List(ilVeh)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        imVefCode = Val(slCode)
        tmDlfSrchKey.iVefCode = imVefCode
        tmDlfSrchKey.sAirDay = slDay
        tmDlfSrchKey.iStartDate(0) = 0
        tmDlfSrchKey.iStartDate(1) = 0
        tmDlfSrchKey.iAirTime(0) = 0
        tmDlfSrchKey.iAirTime(1) = 0
        ilRet = btrGetGreaterOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay)
            'ilRet = btrGetPosition(hmDlf, llDlfRecPos)
            llDlfCode = tlDlf.lCode
            ilIgnoreRec = False
            For ilIndex = LBONE To UBound(tmDlf) - 1 Step 1
                If (tmDlf(ilIndex).iStatus = 1) Or (tmDlf(ilIndex).iStatus = 2) Then  'Insert new record
                    If llDlfCode = tmDlf(ilIndex).lDlfCode Then
                        ilIgnoreRec = True  'Model record- process later
                        Exit For
                    End If
                End If
            Next ilIndex
            If Not ilIgnoreRec Then
                If (tlDlf.iTermDate(0) <> 0) Or (tlDlf.iTermDate(1) <> 0) Then
                    gUnpackDate tlDlf.iTermDate(0), tlDlf.iTermDate(1), slDate
                    If gDateValue(slDate) < llNowDate Then
                        ''The GetNext will still work even when record is deleted
                        'ilRet = btrGetPosition(hmDlf, llDlfRecPos)
                        'If ilRet <> BTRV_ERR_NONE Then
                        '    '6/6/16: Replaced GoSub
                        '    'GoSub mAbortSaveRec
                        '    mAbortSaveRec
                        '    mSaveRec = False
                        '    Exit Function
                        'End If
                        Do
                            'tmRec = tlDlf
                            'ilRet = gGetByKeyForUpdate("DLF", hmDlf, tmRec)
                            'tlDlf = tmRec
                            'If ilRet <> BTRV_ERR_NONE Then
                            '    GoSub mAbortSaveRec
                            '    Exit Function
                            'End If
                            ilRet = btrDelete(hmDlf)
                            If ilRet = BTRV_ERR_CONFLICT Then
                                tmDlfSrchKey.iVefCode = imVefCode
                                tmDlfSrchKey.sAirDay = slDay
                                tmDlfSrchKey.iStartDate(0) = 0
                                tmDlfSrchKey.iStartDate(1) = 0
                                tmDlfSrchKey.iAirTime(0) = 0
                                tmDlfSrchKey.iAirTime(1) = 0
                                ilCRet = btrGetGreaterOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay)
                                    If tlDlf.lCode = llDlfCode Then
                                        Exit Do
                                    End If
                                    ilCRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                Loop
                                If (ilCRet <> BTRV_ERR_NONE) Or (tlDlf.lCode <> llDlfCode) Then
                                    '6/6/16: Replaced GoSub
                                    'GoSub mAbortSaveRec
                                    mAbortSaveRec
                                    mSaveRec = False
                                    Exit Function
                                End If
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            '6/6/16: Replaced GoSub
                            'GoSub mAbortSaveRec
                            mAbortSaveRec
                            mSaveRec = False
                            Exit Function
                        End If
                    Else
                        If (imEndDate0 = 0) And (imEndDate1 = 0) Then
                            If (imTermDate1 < tlDlf.iStartDate(1)) Or ((imTermDate1 = tlDlf.iStartDate(1)) And (imTermDate0 < tlDlf.iStartDate(0))) Then
                                'ilRet = btrGetPosition(hmDlf, llDlfRecPos)
                                'If ilRet <> BTRV_ERR_NONE Then
                                '    '6/6/16: Replaced GoSub
                                '    'GoSub mAbortSaveRec
                                '    mAbortSaveRec
                                '    mSaveRec = False
                                '    Exit Function
                                'End If
                                Do
                                    'tmRec = tlDlf
                                    'ilRet = gGetByKeyForUpdate("DLF", hmDlf, tmRec)
                                    'tlDlf = tmRec
                                    'If ilRet <> BTRV_ERR_NONE Then
                                    '    GoSub mAbortSaveRec
                                    '    Exit Function
                                    'End If
                                    ilRet = btrDelete(hmDlf)
                                    If ilRet = BTRV_ERR_CONFLICT Then
                                        tmDlfSrchKey.iVefCode = imVefCode
                                        tmDlfSrchKey.sAirDay = slDay
                                        tmDlfSrchKey.iStartDate(0) = 0
                                        tmDlfSrchKey.iStartDate(1) = 0
                                        tmDlfSrchKey.iAirTime(0) = 0
                                        tmDlfSrchKey.iAirTime(1) = 0
                                        ilCRet = btrGetGreaterOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                        Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay)
                                            If tlDlf.lCode = llDlfCode Then
                                                Exit Do
                                            End If
                                            ilCRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        Loop
                                        If (ilCRet <> BTRV_ERR_NONE) Or (tlDlf.lCode <> llDlfCode) Then
                                            '6/6/16: Replaced GoSub
                                            'GoSub mAbortSaveRec
                                            mAbortSaveRec
                                            mSaveRec = False
                                            Exit Function
                                        End If
                                    End If
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                If ilRet <> BTRV_ERR_NONE Then
                                    '6/6/16: Replaced GoSub
                                    'GoSub mAbortSaveRec
                                    mAbortSaveRec
                                    mSaveRec = False
                                    Exit Function
                                End If
                            Else
                                If (imTermDate1 < tlDlf.iTermDate(1)) Or ((imTermDate1 = tlDlf.iTermDate(1)) And (imTermDate0 < tlDlf.iTermDate(0))) Then
                                    'ilRet = btrGetPosition(hmDlf, llDlfRecPos)
                                    'If ilRet <> BTRV_ERR_NONE Then
                                    '    '6/6/16: Replaced GoSub
                                    '    'GoSub mAbortSaveRec
                                    '    mAbortSaveRec
                                    '    mSaveRec = False
                                    '    Exit Function
                                    'End If
                                    Do
                                        'tmRec = tlDlf
                                        'ilRet = gGetByKeyForUpdate("DLF", hmDlf, tmRec)
                                        'tlDlf = tmRec
                                        'If ilRet <> BTRV_ERR_NONE Then
                                        '    GoSub mAbortSaveRec
                                        '    Exit Function
                                        'End If
                                        tlDlf.iTermDate(0) = imTermDate0
                                        tlDlf.iTermDate(1) = imTermDate1
                                        ilRet = btrUpdate(hmDlf, tlDlf, imDlfRecLen)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            tmDlfSrchKey.iVefCode = imVefCode
                                            tmDlfSrchKey.sAirDay = slDay
                                            tmDlfSrchKey.iStartDate(0) = 0
                                            tmDlfSrchKey.iStartDate(1) = 0
                                            tmDlfSrchKey.iAirTime(0) = 0
                                            tmDlfSrchKey.iAirTime(1) = 0
                                            ilCRet = btrGetGreaterOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                            Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay)
                                                If tlDlf.lCode = llDlfCode Then
                                                    Exit Do
                                                End If
                                                ilCRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            Loop
                                            If (ilCRet <> BTRV_ERR_NONE) Or (tlDlf.lCode <> llDlfCode) Then
                                                '6/6/16: Replaced GoSub
                                                'GoSub mAbortSaveRec
                                                mAbortSaveRec
                                                mSaveRec = False
                                                Exit Function
                                            End If
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    If ilRet <> BTRV_ERR_NONE Then
                                        '6/6/16: Replaced GoSub
                                        'GoSub mAbortSaveRec
                                        mAbortSaveRec
                                        mSaveRec = False
                                        Exit Function
                                    End If
                                End If
                            End If
                        Else
                            If (tlDlf.iStartDate(1) < imEndDate1) Or ((tlDlf.iStartDate(1) = imEndDate1) And (tlDlf.iStartDate(0) < imEndDate0)) Then
                                If (imTermDate1 < tlDlf.iStartDate(1)) Or ((imTermDate1 = tlDlf.iStartDate(1)) And (imTermDate0 < tlDlf.iStartDate(0))) Then
                                    'ilRet = btrGetPosition(hmDlf, llDlfRecPos)
                                    'If ilRet <> BTRV_ERR_NONE Then
                                    '    '6/6/16: Replaced GoSub
                                    '    'GoSub mAbortSaveRec
                                    '    mAbortSaveRec
                                    '    mSaveRec = False
                                    '    Exit Function
                                    'End If
                                    Do
                                        'tmRec = tlDlf
                                        'ilRet = gGetByKeyForUpdate("DLF", hmDlf, tmRec)
                                        'tlDlf = tmRec
                                        'If ilRet <> BTRV_ERR_NONE Then
                                        '    GoSub mAbortSaveRec
                                        '    Exit Function
                                        'End If
                                        ilRet = btrDelete(hmDlf)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            tmDlfSrchKey.iVefCode = imVefCode
                                            tmDlfSrchKey.sAirDay = slDay
                                            tmDlfSrchKey.iStartDate(0) = 0
                                            tmDlfSrchKey.iStartDate(1) = 0
                                            tmDlfSrchKey.iAirTime(0) = 0
                                            tmDlfSrchKey.iAirTime(1) = 0
                                            ilCRet = btrGetGreaterOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                            Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay)
                                                If tlDlf.lCode = llDlfCode Then
                                                    Exit Do
                                                End If
                                                ilCRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            Loop
                                            If (ilCRet <> BTRV_ERR_NONE) Or (tlDlf.lCode <> llDlfCode) Then
                                                '6/6/16: Replaced GoSub
                                                'GoSub mAbortSaveRec
                                                mAbortSaveRec
                                                mSaveRec = False
                                                Exit Function
                                            End If
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    If ilRet <> BTRV_ERR_NONE Then
                                        '6/6/16: Replaced GoSub
                                        'GoSub mAbortSaveRec
                                        mAbortSaveRec
                                        mSaveRec = False
                                        Exit Function
                                    End If
                                Else
                                    If (imTermDate1 < tlDlf.iTermDate(1)) Or ((imTermDate1 = tlDlf.iTermDate(1)) And (imTermDate0 < tlDlf.iTermDate(0))) Then
                                        'ilRet = btrGetPosition(hmDlf, llDlfRecPos)
                                        'If ilRet <> BTRV_ERR_NONE Then
                                        '    '6/6/16: Replaced GoSub
                                        '    'GoSub mAbortSaveRec
                                        '    mAbortSaveRec
                                        '    mSaveRec = False
                                        '    Exit Function
                                        'End If
                                        Do
                                            'tmRec = tlDlf
                                            'ilRet = gGetByKeyForUpdate("DLF", hmDlf, tmRec)
                                            'tlDlf = tmRec
                                            'If ilRet <> BTRV_ERR_NONE Then
                                            '    GoSub mAbortSaveRec
                                            '    Exit Function
                                            'End If
                                            tlDlf.iTermDate(0) = imTermDate0
                                            tlDlf.iTermDate(1) = imTermDate1
                                            ilRet = btrUpdate(hmDlf, tlDlf, imDlfRecLen)
                                            If ilRet = BTRV_ERR_CONFLICT Then
                                                tmDlfSrchKey.iVefCode = imVefCode
                                                tmDlfSrchKey.sAirDay = slDay
                                                tmDlfSrchKey.iStartDate(0) = 0
                                                tmDlfSrchKey.iStartDate(1) = 0
                                                tmDlfSrchKey.iAirTime(0) = 0
                                                tmDlfSrchKey.iAirTime(1) = 0
                                                ilCRet = btrGetGreaterOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                                Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay)
                                                    If tlDlf.lCode = llDlfCode Then
                                                        Exit Do
                                                    End If
                                                    ilCRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                Loop
                                                If (ilCRet <> BTRV_ERR_NONE) Or (tlDlf.lCode <> llDlfCode) Then
                                                    '6/6/16: Replaced GoSub
                                                    'GoSub mAbortSaveRec
                                                    mAbortSaveRec
                                                    mSaveRec = False
                                                    Exit Function
                                                End If
                                            End If
                                        Loop While ilRet = BTRV_ERR_CONFLICT
                                        If ilRet <> BTRV_ERR_NONE Then
                                            '6/6/16: Replaced GoSub
                                            'GoSub mAbortSaveRec
                                            mAbortSaveRec
                                            mSaveRec = False
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else    'Model from records prior to TFN records (the terminate date should always be prior to start date)
                    If (imTermDate1 < tlDlf.iStartDate(1)) Or ((imTermDate1 = tlDlf.iStartDate(1)) And (imTermDate0 < tlDlf.iStartDate(0))) Then
                        'ilRet = btrGetPosition(hmDlf, llDlfRecPos)
                        'If ilRet <> BTRV_ERR_NONE Then
                        '    '6/6/16: Replaced GoSub
                        '    'GoSub mAbortSaveRec
                        '    mAbortSaveRec
                        '    mSaveRec = False
                        '    Exit Function
                        'End If
                        Do
                            'tmRec = tlDlf
                            'ilRet = gGetByKeyForUpdate("DLF", hmDlf, tmRec)
                            'tlDlf = tmRec
                            'If ilRet <> BTRV_ERR_NONE Then
                            '    GoSub mAbortSaveRec
                            '    Exit Function
                            'End If
                            ilRet = btrDelete(hmDlf)
                            If ilRet = BTRV_ERR_CONFLICT Then
                                tmDlfSrchKey.iVefCode = imVefCode
                                tmDlfSrchKey.sAirDay = slDay
                                tmDlfSrchKey.iStartDate(0) = 0
                                tmDlfSrchKey.iStartDate(1) = 0
                                tmDlfSrchKey.iAirTime(0) = 0
                                tmDlfSrchKey.iAirTime(1) = 0
                                ilCRet = btrGetGreaterOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay)
                                    If tlDlf.lCode = llDlfCode Then
                                        Exit Do
                                    End If
                                    ilCRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                Loop
                                If (ilCRet <> BTRV_ERR_NONE) Or (tlDlf.lCode <> llDlfCode) Then
                                    '6/6/16: Replaced GoSub
                                    'GoSub mAbortSaveRec
                                    mAbortSaveRec
                                    mSaveRec = False
                                    Exit Function
                                End If
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            '6/6/16: Replaced GoSub
                            'GoSub mAbortSaveRec
                            mAbortSaveRec
                            mSaveRec = False
                            Exit Function
                        End If
                    Else
                        'ilRet = btrGetPosition(hmDlf, llDlfRecPos)
                        'If ilRet <> BTRV_ERR_NONE Then
                        '    '6/6/16: Replaced GoSub
                        '    'GoSub mAbortSaveRec
                        '    mAbortSaveRec
                        '    mSaveRec = False
                        '    Exit Function
                        'End If
                        Do
                            'tmRec = tlDlf
                            'ilRet = gGetByKeyForUpdate("DLF", hmDlf, tmRec)
                            'tlDlf = tmRec
                            'If ilRet <> BTRV_ERR_NONE Then
                            '    GoSub mAbortSaveRec
                            '    Exit Function
                            'End If
                            tlDlf.iTermDate(0) = imTermDate0
                            tlDlf.iTermDate(1) = imTermDate1
                            ilRet = btrUpdate(hmDlf, tlDlf, imDlfRecLen)
                            If ilRet = BTRV_ERR_CONFLICT Then
                                tmDlfSrchKey.iVefCode = imVefCode
                                tmDlfSrchKey.sAirDay = slDay
                                tmDlfSrchKey.iStartDate(0) = 0
                                tmDlfSrchKey.iStartDate(1) = 0
                                tmDlfSrchKey.iAirTime(0) = 0
                                tmDlfSrchKey.iAirTime(1) = 0
                                ilCRet = btrGetGreaterOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay)
                                    If tlDlf.lCode = llDlfCode Then
                                        Exit Do
                                    End If
                                    ilCRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                Loop
                                If (ilCRet <> BTRV_ERR_NONE) Or (tlDlf.lCode <> llDlfCode) Then
                                    '6/6/16: Replaced GoSub
                                    'GoSub mAbortSaveRec
                                    mAbortSaveRec
                                    mSaveRec = False
                                    Exit Function
                                End If
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            '6/6/16: Replaced GoSub
                            'GoSub mAbortSaveRec
                            mAbortSaveRec
                            mSaveRec = False
                            Exit Function
                        End If
                    End If
                End If
            End If
            ilRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Loop
    Next ilVeh
    'Terminate old records (records model from)
    imDlfRecLen = Len(tlDlf)  'Get and save DlF record length
    slMsg = "mSaveRec (btrUpdate: Delivery links)"
    For ilIndex = LBONE To UBound(tmDlf) - 1 Step 1
        If (tmDlf(ilIndex).iStatus = 1) Or (tmDlf(ilIndex).iStatus = 2) Then  'Insert new record
            Do
                'ilRet = btrGetDirect(hmDlf, tlDlf, imDlfRecLen, tmDlf(ilIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                tmDlfSrchKey1.lCode = tmDlf(ilIndex).lDlfCode
                ilRet = btrGetEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                If ilRet <> BTRV_ERR_NONE Then
                    '6/6/16: Replaced GoSub
                    'GoSub mAbortSaveRec
                    mAbortSaveRec
                    mSaveRec = False
                    Exit Function
                End If
                If (imTermDate1 < tlDlf.iStartDate(1)) Or ((imTermDate1 = tlDlf.iStartDate(1)) And (imTermDate0 < tlDlf.iStartDate(0))) Then
                    'ilRet = btrGetPosition(hmDlf, llDlfRecPos)
                    If ilRet <> BTRV_ERR_NONE Then
                        '6/6/16: Replaced GoSub
                        'GoSub mAbortSaveRec
                        mAbortSaveRec
                        mSaveRec = False
                        Exit Function
                    End If
                    Do
                        'tmRec = tlDlf
                        'ilRet = gGetByKeyForUpdate("DLF", hmDlf, tmRec)
                        'tlDlf = tmRec
                        'If ilRet <> BTRV_ERR_NONE Then
                        '    GoSub mAbortSaveRec
                        '    Exit Function
                        'End If
                        ilRet = btrDelete(hmDlf)
                        If ilRet = BTRV_ERR_CONFLICT Then
                            tmDlfSrchKey1.lCode = tmDlf(ilIndex).lDlfCode
                            ilCRet = btrGetEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                            If ilCRet <> BTRV_ERR_NONE Then
                                '6/6/16: Replaced GoSub
                                'GoSub mAbortSaveRec
                                mAbortSaveRec
                                mSaveRec = False
                                Exit Function
                            End If
                        End If
                    Loop While ilRet = BTRV_ERR_CONFLICT
                Else
                    'tmRec = tlDlf
                    'ilRet = gGetByKeyForUpdate("DLF", hmDlf, tmRec)
                    'tlDlf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    GoSub mAbortSaveRec
                    '    Exit Function
                    'End If
                    tlDlf.iTermDate(0) = imTermDate0
                    tlDlf.iTermDate(1) = imTermDate1
                    ilRet = btrUpdate(hmDlf, tlDlf, imDlfRecLen)
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                '6/6/16: Replaced GoSub
                'GoSub mAbortSaveRec
                mAbortSaveRec
                mSaveRec = False
                Exit Function
            End If
        End If
    Next ilIndex
    slMsg = "mSaveRec (btrInsert: Delivery links)"
    For ilIndex = LBONE To UBound(tmDlf) - 1 Step 1
        If (tmDlf(ilIndex).iStatus = 0) Or (tmDlf(ilIndex).iStatus = 1) Then  'Insert new record
            'Only required if iStatus =1 but Ok for iStatus = 0
            tmDlf(ilIndex).DlfRec.lCode = 0
            tmDlf(ilIndex).DlfRec.iStartDate(0) = imDate0
            tmDlf(ilIndex).DlfRec.iStartDate(1) = imDate1
            tmDlf(ilIndex).DlfRec.iTermDate(0) = imEndDate0
            tmDlf(ilIndex).DlfRec.iTermDate(1) = imEndDate1
            ilRet = btrInsert(hmDlf, tmDlf(ilIndex).DlfRec, imDlfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                '6/6/16: Replaced GoSub
                'GoSub mAbortSaveRec
                mAbortSaveRec
                mSaveRec = False
                Exit Function
            End If
            tmDlf(ilIndex).iStatus = 1
            'ilRet = btrGetPosition(hmDlf, tmDlf(ilIndex).lRecPos)
            'If ilRet <> BTRV_ERR_NONE Then
            '    '6/6/16: Replaced GoSub
            '    'GoSub mAbortSaveRec
            '    mAbortSaveRec
            '    mSaveRec = False
            '    Exit Function
            'End If
            tmDlf(ilIndex).lDlfCode = tmDlf(ilIndex).DlfRec.lCode
        End If
    Next ilIndex
    ilRet = btrEndTrans(hmDlf)
    'ilRet = btrGetFirst(hmDlf, tlDlf, imDlfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
    'Do While (ilRet = BTRV_ERR_NONE)
    '    If (tlDlf.iTermDate(0) <> 0) Or (tlDlf.iTermDate(1) <> 0) Then
    '        'remove any previously terminated before start
    '        gUnpackDate tlDlf.iTermDate(0), tlDlf.iTermDate(1), slDate
    '        If gDateValue(slDate) < llNowDate Then
    '            'The GetNext will still work even when record is deleted
    '            ilRet = btrDelete(hmDlf)
    '        End If
    '    End If
    '    ilRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'Loop
    ''Remove record that don't belong within Delivery
    'If imDelOrEngr = 0 Then
    '    ilRet = btrGetFirst(hmDlf, tlDlf, imDlfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
    '    Do While (ilRet = BTRV_ERR_NONE)
    '        If (tlDlf.sCmmlSched = "N") And (tlDlf.iMnfSubFeed = 0) Then
    '            'The GetNext will still work even when record is deleted
    '            ilRet = btrDelete(hmDlf)
    '        End If
    '        ilRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    '    Loop
    'End If
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function

    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    mSaveRec = False
    Exit Function
'mAbortSaveRec:
'    ilRet = btrAbortTrans(hmDlf)
'    Screen.MousePointer = vbDefault
'    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Links")
'    imTerminate = True
'    mSaveRec = False
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name: mSaveRecChg                    *
'*                                                     *
'*             Created:7/27/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine of record should be   *
'*                     saved                           *
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
    If imChg Then
        If ilAsk Then
            slMess = "Save changes"
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                mSaveRecChg = False
                Exit Function
            End If
            If ilRes = vbYes Then
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
            If ilRes = vbNo Then
                mSaveRecChg = True
                Exit Function
            End If
        Else
            ilRes = mSaveRec()
            mSaveRecChg = ilRes
            Exit Function
        End If
    End If
    mSaveRecChg = True
End Function
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

    If (imChg) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    If (lacFrame(imDelIndex).Visible) And (imUpdateAllowed) Then
        cmcDupl.Enabled = True
    Else
        cmcDupl.Enabled = False
    End If
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mSetFocus                      *
'*                                                     *
'*             Created:7/27/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set focus to controls           *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If

    If (imRowNo < vbcDelivery.Value) Or (imRowNo >= vbcDelivery.Value + vbcDelivery.LargeChange + 1) Then
        mSetShow ilBoxNo, False
        Exit Sub
    End If

    If Not mFindRowIndex(imRowNo) Then
        pbcArrow.Visible = False
        lacFrame(imDelIndex).Visible = False
        Exit Sub
    End If
    lacFrame(imDelIndex).Move 0, tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) - 30
    lacFrame(imDelIndex).Visible = True
    pbcArrow.Move pbcArrow.Left, plcDelivery.Top + tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHICLEINDEX
            If edcDropDown.Enabled Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case LOCALTIMEINDEX
            If edcDropDown.Enabled Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case FEEDTIMEINDEX
            If edcDropDown.Enabled Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case TIMEZONEINDEX
            If edcDropDown.Enabled Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case SUBFEEDINDEX 'Event Type
            If edcDropDown.Enabled Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case EVTNAMEINDEX 'Event Name
            If edcDropDown.Enabled Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case PROGCODEINDEX
            If edcDropDown.Enabled Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case SHOWONINDEX
            If pbcYN.Enabled Then
                pbcYN.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case FEDINDEX
            If pbcYN.Enabled Then
                pbcYN.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case BUSINDEX
            If edcDropDown.Enabled Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case SCHINDEX
            If edcDropDown.Enabled Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
    End Select
'    mSetCommands
End Sub
Private Sub mSetShow(ilBoxNo As Integer, ilArrow As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'       ilArrow (I)- True=Leave on; False=Make invisible
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilRet As Integer
    Dim ilTime0 As Integer
    Dim ilTime1 As Integer
    Dim slChar As String
    Dim slBus As String * 5
    If Not ilArrow Then
        pbcArrow.Visible = False
        lacFrame(imDelIndex).Visible = False
    End If
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHICLEINDEX
            If imDelIndex = 0 Then
                lbcVehicle.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                If lbcVehicle.ListIndex >= 0 Then
                    tmDlf(imDlfIndex).sVehicle = lbcVehicle.List(lbcVehicle.ListIndex)
                    slNameCode = lbcVehCode.List(lbcVehicle.ListIndex)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If tmDlf(imDlfIndex).DlfRec.iVefCode <> Val(slCode) Then
                        imChg = True
                        tmDlf(imDlfIndex).DlfRec.iEnfCode = 0
                        tmDlf(imDlfIndex).sEventName = ""
                    End If
                    tmDlf(imDlfIndex).DlfRec.iVefCode = Val(slCode)
                Else
                    tmDlf(imDlfIndex).sVehicle = ""
                    tmDlf(imDlfIndex).DlfRec.iVefCode = 0
                End If
            Else
                lbcFeed.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                If lbcFeed.ListIndex >= 0 Then
                    tmDlf(imDlfIndex).sFeed = lbcFeed.List(lbcFeed.ListIndex)
                    slNameCode = lbcFeedCode.List(lbcFeed.ListIndex)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If tmDlf(imDlfIndex).DlfRec.iMnfFeed <> Val(slCode) Then
                        imChg = True
                    End If
                    tmDlf(imDlfIndex).DlfRec.iMnfFeed = Val(slCode)
                Else
                    tmDlf(imDlfIndex).sFeed = ""
                    tmDlf(imDlfIndex).DlfRec.iMnfFeed = 0
                End If
            End If
        Case LOCALTIMEINDEX 'Local Time index
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                tmDlf(imDlfIndex).sLocalTime = slStr
                gPackTime slStr, ilTime0, ilTime1
                If (ilTime0 <> tmDlf(imDlfIndex).DlfRec.iLocalTime(0)) Or (ilTime1 <> tmDlf(imDlfIndex).DlfRec.iLocalTime(1)) Then
                    imChg = True
                End If
                tmDlf(imDlfIndex).DlfRec.iLocalTime(0) = ilTime0
                tmDlf(imDlfIndex).DlfRec.iLocalTime(1) = ilTime1
            Else
                Beep
                edcDropDown.Text = tmDlf(imDlfIndex).sLocalTime
            End If
        Case FEEDTIMEINDEX 'Feed Time index
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                tmDlf(imDlfIndex).sFeedTime = slStr
                gPackTime slStr, ilTime0, ilTime1
                If (ilTime0 <> tmDlf(imDlfIndex).DlfRec.iFeedTime(0)) Or (ilTime1 <> tmDlf(imDlfIndex).DlfRec.iFeedTime(1)) Then
                    imChg = True
                End If
                tmDlf(imDlfIndex).DlfRec.iFeedTime(0) = ilTime0
                tmDlf(imDlfIndex).DlfRec.iFeedTime(1) = ilTime1
            Else
                Beep
                edcDropDown.Text = tmDlf(imDlfIndex).sFeedTime
            End If
        Case TIMEZONEINDEX 'Time Zone
            edcDropDown.Visible = False  'Set visibility
            If Trim$(tmDlf(imDlfIndex).DlfRec.sZone) <> edcDropDown.Text Then
                imChg = True
            End If
            tmDlf(imDlfIndex).DlfRec.sZone = edcDropDown.Text
        Case SUBFEEDINDEX 'Event Type
            lbcSubfeed.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcSubfeed.ListIndex > 0 Then
                tmDlf(imDlfIndex).sSubfeed = lbcSubfeed.List(lbcSubfeed.ListIndex)
                slNameCode = tmSubFeedCode(lbcSubfeed.ListIndex - 1).sKey   'lbcSubFeedCode.List(lbcSubFeed.ListIndex - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If tmDlf(imDlfIndex).DlfRec.iMnfSubFeed <> Val(slCode) Then
                    imChg = True
                End If
                tmDlf(imDlfIndex).DlfRec.iMnfSubFeed = Val(slCode)
            Else
                tmDlf(imDlfIndex).sSubfeed = ""
                If tmDlf(imDlfIndex).DlfRec.iMnfSubFeed <> 0 Then
                    imChg = True
                End If
                tmDlf(imDlfIndex).DlfRec.iMnfSubFeed = 0
            End If
            If imDelOrEngr = 0 Then
                If tmDlf(imDlfIndex).DlfRec.iMnfSubFeed = 0 Then
                    tmDlf(imDlfIndex).DlfRec.sCmmlSched = "Y"
                Else
                    tmDlf(imDlfIndex).DlfRec.sCmmlSched = "N"
                End If
            Else
                tmDlf(imDlfIndex).DlfRec.sCmmlSched = "N"
            End If
        Case EVTNAMEINDEX 'Event Name
            lbcEvtName(imEvtNameIndex).Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcEvtName(imEvtNameIndex).ListIndex >= 0 Then
                tmDlf(imDlfIndex).sEventName = lbcEvtName(imEvtNameIndex).List(lbcEvtName(imEvtNameIndex).ListIndex)
                slNameCode = lbcEvtNameCode(imEvtNameIndex).List(lbcEvtName(imEvtNameIndex).ListIndex)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If tmDlf(imDlfIndex).DlfRec.iEnfCode <> Val(slCode) Then
                    imChg = True
                End If
                tmDlf(imDlfIndex).DlfRec.iEnfCode = Val(slCode)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                tmDlf(imDlfIndex).sEventName = slName
            Else
                If tmDlf(imDlfIndex).DlfRec.iEnfCode <> 0 Then
                    imChg = True
                End If
                tmDlf(imDlfIndex).DlfRec.iEnfCode = 0
                tmDlf(imDlfIndex).sEventName = ""
            End If
        Case PROGCODEINDEX 'Time Zone
            edcDropDown.Visible = False  'Set visibility
            If Trim$(tmDlf(imDlfIndex).DlfRec.sProgCode) <> Trim$(edcDropDown.Text) Then
                imChg = True
            End If
            tmDlf(imDlfIndex).DlfRec.sProgCode = Trim$(edcDropDown.Text)
        Case SHOWONINDEX
            pbcYN.Visible = False
        Case FEDINDEX
            pbcYN.Visible = False
        Case BUSINDEX 'Time Zone
            edcDropDown.Visible = False  'Set visibility
            slStr = ""
            slBus = edcDropDown.Text
            If imDelOrEngr = 1 Then
                For ilLoop = 1 To 5 Step 1
                    slChar = Mid$(slBus, ilLoop, 1)
                    If slChar = "-" Then
                        slStr = slStr & " "
                    Else
                        slStr = slStr & slChar
                    End If
                Next ilLoop
                slBus = slStr
            End If
            If tmDlf(imDlfIndex).DlfRec.sBus <> slBus Then
                imChg = True
            End If
            tmDlf(imDlfIndex).DlfRec.sBus = slBus
        Case SCHINDEX 'Time Zone
            edcDropDown.Visible = False  'Set visibility
            If Trim$(tmDlf(imDlfIndex).DlfRec.sSchedule) <> Trim$(edcDropDown.Text) Then
                imChg = True
            End If
            tmDlf(imDlfIndex).DlfRec.sSchedule = Trim$(edcDropDown.Text)
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSort                           *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified: 4/24/94      By:D. Hannifan    *
'*                                                     *
'*            Comments: Sort records                   *
'*                                                     *
'*******************************************************
Private Sub mSort()
    Dim ilLoop As Integer
    Dim ilSort As Integer
    Dim slATime As String
    Dim slTime As String
    Dim llTime As Long
    Dim slEvtType As String
    Dim slAddItem As String
    Dim slStr As String
    Dim ilEvtAddToTime As Integer
    lbcSortIndex.Clear
    'Sort by time
    For ilLoop = LBONE To UBound(tmDlf) - 1 Step 1
        'If (tmDlf(ilLoop).iStatus = 1) Or (tmDlf(ilLoop).iStatus = 0) Then
        If (tmDlf(ilLoop).iStatus = 1) Or (tmDlf(ilLoop).iStatus = 0) Or (tmDlf(ilLoop).iStatus = 2) Then
            slAddItem = ""
            ilEvtAddToTime = False
            For ilSort = UBound(imSort) To LBound(imSort) Step -1
                Select Case imSort(ilSort)
                    Case 0  'Vehicle
                        If imDelIndex = 0 Then
                            slStr = Left$(tmDlf(ilLoop).sVehicle, 6)
                            Do While Len(slStr) < 6
                                slStr = slStr & " "
                            Loop
                            slAddItem = slStr & slAddItem
                        Else
                            slStr = Left$(tmDlf(ilLoop).sFeed, 6)
                            Do While Len(slStr) < 6
                                slStr = slStr & " "
                            Loop
                            slAddItem = slStr & slAddItem
                        End If
                    Case 1  'Air time
                        gUnpackTime tmDlf(ilLoop).DlfRec.iAirTime(0), tmDlf(ilLoop).DlfRec.iAirTime(1), "A", "1", slTime
                        llTime = CLng(gTimeToCurrency(slTime, True))
                        slTime = Trim$(Str$(llTime))
                        Do While Len(slTime) < 5
                            slTime = "0" & slTime
                        Loop
                        If Not ilEvtAddToTime Then
                            slEvtType = Trim$(Str$(tmDlf(ilLoop).DlfRec.iEtfCode))
                            Do While Len(slEvtType) < 3
                                slEvtType = "0" & slEvtType
                            Loop
                            slAddItem = slTime & slEvtType & slAddItem
                            ilEvtAddToTime = True
                        Else
                            slAddItem = slTime & slAddItem
                        End If
                    Case 2  'Affiliate time
                        gUnpackTime tmDlf(ilLoop).DlfRec.iAirTime(0), tmDlf(ilLoop).DlfRec.iAirTime(1), "A", "1", slATime
                        gUnpackTime tmDlf(ilLoop).DlfRec.iLocalTime(0), tmDlf(ilLoop).DlfRec.iLocalTime(1), "A", "1", slTime
                        llTime = CLng(gTimeToCurrency(slTime, True))
                        If (InStr(slATime, "A") <> 0) And (InStr(slTime, "P") <> 0) Then
                            llTime = CLng(gTimeToCurrency(slTime, True))
                            llTime = llTime - 43200
                        ElseIf (InStr(slATime, "P") <> 0) And (InStr(slTime, "A") <> 0) Then
                            llTime = CLng(gTimeToCurrency(slTime, False))
                            llTime = llTime + 129600
                        Else
                            llTime = CLng(gTimeToCurrency(slTime, False))
                            llTime = llTime + 43200
                        End If
                        slTime = Trim$(Str$(llTime))
                        Do While Len(slTime) < 6    '5
                            slTime = "0" & slTime
                        Loop
                        If Not ilEvtAddToTime Then
                            slEvtType = Trim$(Str$(tmDlf(ilLoop).DlfRec.iEtfCode))
                            Do While Len(slEvtType) < 3
                                slEvtType = "0" & slEvtType
                            Loop
                            slAddItem = slTime & slEvtType & slAddItem
                            ilEvtAddToTime = True
                        Else
                            slAddItem = slTime & slAddItem
                        End If
                    Case 3  'Feed time
                        gUnpackTime tmDlf(ilLoop).DlfRec.iAirTime(0), tmDlf(ilLoop).DlfRec.iAirTime(1), "A", "1", slATime
                        gUnpackTime tmDlf(ilLoop).DlfRec.iFeedTime(0), tmDlf(ilLoop).DlfRec.iFeedTime(1), "A", "1", slTime
                        llTime = CLng(gTimeToCurrency(slTime, True))
                        If (InStr(slATime, "A") <> 0) And (InStr(slTime, "P") <> 0) Then
                            llTime = CLng(gTimeToCurrency(slTime, True))
                            llTime = llTime - 43200
                        ElseIf (InStr(slATime, "P") <> 0) And (InStr(slTime, "A") <> 0) Then
                            llTime = CLng(gTimeToCurrency(slTime, False))
                            llTime = llTime + 129600
                        Else
                            llTime = CLng(gTimeToCurrency(slTime, False))
                            llTime = llTime + 43200
                        End If
                        slTime = Trim$(Str$(llTime))
                        Do While Len(slTime) < 6    '5
                            slTime = "0" & slTime
                        Loop
                        If Not ilEvtAddToTime Then
                            slEvtType = Trim$(Str$(tmDlf(ilLoop).DlfRec.iEtfCode))
                            Do While Len(slEvtType) < 3
                                slEvtType = "0" & slEvtType
                            Loop
                            slAddItem = slTime & slEvtType & slAddItem
                            ilEvtAddToTime = True
                        Else
                            slAddItem = slTime & slAddItem
                        End If
                    Case 4  'Zone
                        Select Case UCase(Left$(tmDlf(ilLoop).DlfRec.sZone, 1))
                            Case "E"
                                slStr = "1"
                            Case "C"
                                slStr = "2"
                            Case "M"
                                slStr = "3"
                            Case "P"
                                slStr = "4"
                        End Select
                        slAddItem = slStr & slAddItem
                    Case 5  'Event Type
                        slEvtType = Trim$(Str$(tmDlf(ilLoop).DlfRec.iEtfCode))
                        Do While Len(slEvtType) < 3
                            slEvtType = "0" & slEvtType
                        Loop
                        slAddItem = slEvtType & slAddItem
                    Case 6  'Event name
                        slStr = Left$(tmDlf(ilLoop).sEventName, 3)
                        Do While Len(slStr) < 3
                            slStr = slStr & " "
                        Loop
                        slAddItem = slStr & slAddItem
                    Case 7  'Program code
                        slStr = Left$(tmDlf(ilLoop).DlfRec.sProgCode, 5)
                        Do While Len(slStr) < 5
                            slStr = slStr & " "
                        Loop
                        slAddItem = slStr & slAddItem
                    Case 8  'Cmml schd
                        slStr = Left$(tmDlf(ilLoop).DlfRec.sCmmlSched, 1)
                        slAddItem = slStr & slAddItem
                    Case 9  'Bus
                        slStr = tmDlf(ilLoop).DlfRec.sBus
                        slAddItem = slStr & slAddItem
                    Case 10 'Schedule
                        slStr = Left$(tmDlf(ilLoop).DlfRec.sSchedule, 1)
                        slAddItem = slStr & slAddItem
                End Select
            Next ilSort
            lbcSortIndex.AddItem slAddItem '& "\" & Trim$(Str$(ilLoop))
            lbcSortIndex.ItemData(lbcSortIndex.NewIndex) = ilLoop
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSubfeedPop                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Feed list             *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mSubfeedPop()
'
'   mSubfeedPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopMnfPlusFieldsBox(LinkDlvy, lbcSubFeed, lbcSubFeedCode, "NOS")
    ilRet = gPopMnfPlusFieldsBox(LinkDlvy, lbcSubfeed, tmSubFeedCode(), smSubFeedCodeTag, "NOS")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSubfeedPopErr
        gCPErrorMsg ilRet, "mSubfeedPop (gPopMnfPlusFieldsBox)", LinkDlvy
        On Error GoTo 0
        lbcSubfeed.AddItem "[None]", 0
    End If
    Exit Sub
mSubfeedPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:4/24/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: terminate LinksDef form        *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'

    'Unload form
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload LinkDlvy
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:4/24/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Test fields                    *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilDlfIndex As Integer) As Integer
'
'   iRet = mTestFields(imDlfIndex)
'   Where:
'       imDlfIndex (I)- Row index to be checked
'       iRet (O)- True if all mandatory fields answered correctly
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilRet As Integer

    If ilDlfIndex <= 0 Then
        mTestFields = YES
        Exit Function
    End If
    'Bypass test as field can't be altered
'    If imDelIndex = 0 Then
'       If tmDlf(ilDlfIndex).DlfRec.iVefCode <= 0 Then
'           ilRes = MsgBox("Vehicle must be specified", vbOkOnly + vbExclamation, "Incomplete")
'           imBoxNo = VEHICLEINDEX
'           mTestFields = No
'           Exit Function
'       End If
'    Else
'       If tmDlf(ilDlfIndex).DlfRec.iMnfFeed <= 0 Then
'           ilRes = MsgBox("Feed must be specified", vbOkOnly + vbExclamation, "Incomplete")
'           imBoxNo = VEHICLEINDEX
'           mTestFields = No
'           Exit Function
'       End If
'    End If
    'Bypass test as field can't be altered (if added, then air time must be checked that its
    'a valid time for the vehicle)
'    If Not gValidTime(tmDlf(ilDlfIndex).sAirTime) Then
'        ilRes = MsgBox("Air Time must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
'        imBoxNo = AIRTIMEINDEX
'        mTestFields = No
'        Exit Function
'    End If
    If imDelOrEngr = 0 Then
        If Not gValidTime(Trim$(tmDlf(ilDlfIndex).sLocalTime)) Then
            ilRes = MsgBox("Affiliate Time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = LOCALTIMEINDEX
            mTestFields = NO
            Exit Function
        End If
    End If
    If imDelOrEngr = 1 Then
        If Not gValidTime(Trim$(tmDlf(ilDlfIndex).sFeedTime)) Then
            ilRes = MsgBox("Feed Time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = FEEDTIMEINDEX
            mTestFields = NO
            Exit Function
        End If
    End If
    If Trim$(tmDlf(ilDlfIndex).DlfRec.sZone) = "" Then
        ilRes = MsgBox("Time Zone must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = TIMEZONEINDEX
        mTestFields = NO
        Exit Function
    End If
    'This field is required but don't test so partially completed screens can
    'be updated
'    If tmMnf.iGroupNo = 1 Then  'Subfeed
'        If tmDlf(ilDlfIndex).DlfRec.iMnfSubfeed <= 0 Then
'            ilRes = MsgBox("Subfeed must be specified", vbOkOnly + vbExclamation, "Incomplete")
'            imBoxNo = SUBFEEDINDEX
'            mTestFields = No
'            Exit Function
'        End If
'    End If
    'Bypass test as field can't be altered
'    If tmDlf(ilDlfIndex).DlfRec.iEtfCode <= 0 Then
'        ilRes = MsgBox("Event type must be specified", vbOkOnly + vbExclamation, "Incomplete")
'        imBoxNo = EVTTYPEINDEX
'        mTestFields = No
'        Exit Function
'    End If
    'Bypass test as field can't be altered
'    If tmDlf(ilDlfIndex).DlfRec.iEnfCode <= 0 Then
'        ilRes = MsgBox("Event name must be specified", vbOkOnly + vbExclamation, "Incomplete")
'        imBoxNo = EVTNAMEINDEX
'        mTestFields = No
'        Exit Function
'    End If
    If imDelOrEngr = 0 Then
        If tmDlf(ilDlfIndex).DlfRec.iVefCode <> tmVef.iCode Then
            tmVefSrchKey.iCode = tmDlf(ilDlfIndex).DlfRec.iVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        End If
        If tmVef.sType <> "C" Then
            If Trim$(tmDlf(ilDlfIndex).DlfRec.sProgCode) = "" Then
                ilRes = MsgBox("Program Code must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imBoxNo = PROGCODEINDEX
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If imDelOrEngr = 0 Then
        If Trim$(tmDlf(ilDlfIndex).DlfRec.sCmmlSched) = "" Then
            ilRes = MsgBox("Show on Commercial Schedule must be specified", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = SHOWONINDEX
            mTestFields = NO
            Exit Function
        End If
    End If
    'This field is required but don't test so partially completed screens can
    'be updated
    'If imDelOrEngr = 0 Then
        If Trim$(tmDlf(ilDlfIndex).DlfRec.sFed) = "" Then
            ilRes = MsgBox("Fed must be specified", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = FEDINDEX
            mTestFields = NO
            Exit Function
        End If
    'End If
    'Bus and schedule are required for dish only
'    If StrComp(Trim$(smFeedType), "Dish", 1) = 0 Then
'        If Trim$(tmDlf(ilDlfIndex).DlfRec.sBus) = "" Then
'            ilRes = MsgBox("Bus must be specified", vbOkOnly + vbExclamation, "Incomplete")
'            imBoxNo = BUSINDEX
'            mTestFields = No
'            Exit Function
'        End If
'        If Trim$(tmDlf(ilDlfIndex).DlfRec.sSchedule) = "" Then
'            ilRes = MsgBox("Schedule must be specified", vbOkOnly + vbExclamation, "Incomplete")
'            imBoxNo = SCHINDEX
'            mTestFields = No
'            Exit Function
'        End If
'    End If
    mTestFields = YES
End Function
Private Sub pbcArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo, False
    imBoxNo = -1
    imRowNo = -1
    lacFrame(imDelIndex).Visible = False
    pbcArrow.Visible = False
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
Private Sub pbcDelivery_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    lacFrame(imDelIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
End Sub
Private Sub pbcDelivery_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcDelivery_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim slStr As String
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    ilCompRow = vbcDelivery.LargeChange + 1
    If imMaxRowNo > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = imMaxRowNo
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    If (ilBox = VEHICLEINDEX) Or (ilBox = AIRTIMEINDEX) Or (ilBox = EVTNAMEINDEX) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If (ilBox = SUBFEEDINDEX) And (tmMnf.iGroupNo <> 1) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
'                    If (imBoxNo = VEHICLEINDEX) And (edcDropDown.Text = "") Then
'                        Beep
'                        edcDropDown.SetFocus
'                        Exit Sub
'                    End If
'                    If imDlfIndex > 0 Then
'                        If imBoxNo > VEHICLEINDEX And tmDlf(imDlfIndex).DlfRec.iVefCode = 0 Then
'                            Beep
'                            ilBox = VEHICLEINDEX
'                            mSetShow imBoxNo
'                            imRowNo = ilRowNo
'                            imBoxNo = ilBox
'                            mEnableBox ilBox
'                            Exit Sub
'                        End If
'                    End If
                    If imBoxNo = LOCALTIMEINDEX Then
                        slStr = edcDropDown.Text
                        If Not gValidTime(slStr) Then
                            Beep
                            If (edcDropDown.Enabled) And (edcDropDown.Visible) Then
                                edcDropDown.SetFocus
                            Else
                                pbcClickFocus.SetFocus
                            End If
                            Exit Sub
                        End If
                    End If
                    If imBoxNo = FEEDTIMEINDEX Then
                        slStr = edcDropDown.Text
                        If Not gValidTime(slStr) Then
                            Beep
                            If (edcDropDown.Enabled) And (edcDropDown.Visible) Then
                                edcDropDown.SetFocus
                            Else
                                pbcClickFocus.SetFocus
                            End If
                            Exit Sub
                        End If
                    End If
                    If (ilBox = LOCALTIMEINDEX) And (imDelOrEngr = 1) Then
                        Beep
                        If (edcDropDown.Enabled) And (edcDropDown.Visible) Then
                            edcDropDown.SetFocus
                        Else
                            pbcClickFocus.SetFocus
                        End If
                        Exit Sub
                    End If
                    If (ilBox = FEEDTIMEINDEX) And (imDelOrEngr = 0) Then
                        Beep
                        If (edcDropDown.Enabled) And (edcDropDown.Visible) Then
                            edcDropDown.SetFocus
                        Else
                            pbcClickFocus.SetFocus
                        End If
                        Exit Sub
                    End If
                    If (ilBox = PROGCODEINDEX) And (imDelOrEngr = 1) Then
                        Beep
                        If (edcDropDown.Enabled) And (edcDropDown.Visible) Then
                            edcDropDown.SetFocus
                        Else
                            pbcClickFocus.SetFocus
                        End If
                        Exit Sub
                    End If
                    If (ilBox = SHOWONINDEX) And (imDelOrEngr = 1) Then
                        Beep
                        If (edcDropDown.Enabled) And (edcDropDown.Visible) Then
                            edcDropDown.SetFocus
                        Else
                            pbcClickFocus.SetFocus
                        End If
                        Exit Sub
                    End If
                    If (ilBox = FEDINDEX) And (imDelOrEngr = 1) Then
                        Beep
                        If (edcDropDown.Enabled) And (edcDropDown.Visible) Then
                            edcDropDown.SetFocus
                        Else
                            pbcClickFocus.SetFocus
                        End If
                        Exit Sub
                    End If
                    'If (ilBox = BUSINDEX) And (imDelOrEngr = 0) Then
                    '    Beep
                    '    If (edcDropDown.Enabled) And (edcDropDown.Visible) Then
                    '        edcDropDown.SetFocus
                    '    Else
                    '        pbcClickFocus.SetFocus
                    '    End If
                    '    Exit Sub
                    'End If
                    If (ilBox = SCHINDEX) And (imDelOrEngr = 0) Then
                        Beep
                        If (edcDropDown.Enabled) And (edcDropDown.Visible) Then
                            edcDropDown.SetFocus
                        Else
                            pbcClickFocus.SetFocus
                        End If
                        Exit Sub
                    End If
                    ilRowNo = ilRow + vbcDelivery.Value - 1
                    imTabDirection = 0  'Set-Left to right
                    mSetShow imBoxNo, False
                    imRowNo = ilRowNo
                    imBoxNo = ilBox
                    If Not mFindRowIndex(imRowNo) Then
                        Beep
                        imBoxNo = -1
                        pbcClickFocus.SetFocus
                        Exit Sub
                    End If
                    'If (imBoxNo = FEEDTIMEINDEX) And (Trim$(tmDlf(imDlfIndex).DlfRec.sFed) = "N") Then
                    '    Beep
                    '    imBoxNo = -1
                    '    pbcClickFocus.SetFocus
                    '    Exit Sub
                    'End If
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSetFocus imBoxNo
End Sub
Private Sub pbcDelivery_Paint(Index As Integer)
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim ilEvt As Integer
    Dim slStr As String
    Dim ilPaintRow As Integer
    Dim ilIndex As Integer
    Dim ilBox As Integer
    Dim slChar As String
    Dim ilLoop As Integer
    Dim ilRowCount As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    ilStartRow = vbcDelivery.Value
    If ilStartRow = 0 Then
        Exit Sub
    End If
    If imDelIndex = 1 Then
        llColor = pbcDelivery(imDelIndex).ForeColor
        slFontName = pbcDelivery(imDelIndex).FontName
        flFontSize = pbcDelivery(imDelIndex).FontSize
        pbcDelivery(imDelIndex).ForeColor = BLUE
        pbcDelivery(imDelIndex).FontBold = False
        pbcDelivery(imDelIndex).FontSize = 7
        pbcDelivery(imDelIndex).FontName = "Arial"
        pbcDelivery(imDelIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcDelivery(imDelIndex).CurrentX = tmCtrls(BUSINDEX).fBoxX + 15 'fgBoxInsetX
        pbcDelivery(imDelIndex).CurrentY = tmCtrls(BUSINDEX).fBoxY - 2 * pbcDelivery(1).TextHeight("1") - 30 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If imDelOrEngr = 0 Then
            pbcDelivery(imDelIndex).Print "Dupe"
            pbcDelivery(imDelIndex).CurrentX = tmCtrls(BUSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcDelivery(imDelIndex).CurrentY = tmCtrls(BUSINDEX).fBoxY - pbcDelivery(1).TextHeight("1") - 30 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcDelivery(imDelIndex).Print "Avail ID"
        Else
            pbcDelivery(imDelIndex).Print "Bus"
        End If
        pbcDelivery(imDelIndex).FontSize = flFontSize
        pbcDelivery(imDelIndex).FontName = slFontName
        pbcDelivery(imDelIndex).FontSize = flFontSize
        pbcDelivery(imDelIndex).ForeColor = llColor
        pbcDelivery(imDelIndex).FontBold = True
    End If
    ilPaintRow = 1
    ilEndRow = vbcDelivery.LargeChange + ilPaintRow
    If ilEndRow > imMaxRowNo Then
        ilEndRow = imMaxRowNo
    End If
    ilEvt = 1
    ilRowCount = 0
    Do While (ilPaintRow <= ilEndRow) And (ilEvt <= lbcSortIndex.ListCount)
        'slNameCode = lbcSortIndex.List(ilEvt - 1)
        'ilRet = gParseItem(slNameCode, 2, "\", slIndex)
        'ilIndex = Val(slIndex)
        ilIndex = lbcSortIndex.ItemData(ilEvt - 1)
        'If (tmDlf(ilIndex).iStatus = 0) Or (tmDlf(ilIndex).iStatus = 1) Then
        If (tmDlf(ilIndex).iStatus = 0) Or (tmDlf(ilIndex).iStatus = 1) Or (tmDlf(ilIndex).iStatus = 2) Then
            'If (tmDlf(ilIndex).DlfRec.sFed = "Y") Or (Not ckcFedEvtOnly.Value) Then
                ilRowCount = ilRowCount + 1
                If ilRowCount >= ilStartRow Then
                    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
                        pbcDelivery(imDelIndex).CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
                        pbcDelivery(imDelIndex).CurrentY = tmCtrls(ilBox).fBoxY + (ilPaintRow - 1) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                        Select Case ilBox
                            Case VEHICLEINDEX
                                If imDelIndex = 0 Then
                                    slStr = Trim$(tmDlf(ilIndex).sVehicle)
                                Else
                                    slStr = Trim$(tmDlf(ilIndex).sFeed)
                                End If
                            Case AIRTIMEINDEX
                                slStr = Trim$(tmDlf(ilIndex).sAirTime)
                                If slStr <> "" Then
                                    slStr = gFormatTime(slStr, "A", "1")
                                End If
                            Case LOCALTIMEINDEX
                                If imDelOrEngr = 0 Then
                                    slStr = Trim$(tmDlf(ilIndex).sLocalTime)
                                    If slStr <> "" Then
                                        slStr = gFormatTime(slStr, "A", "1")
                                    End If
                                Else
                                    slStr = ""
                                End If
                            Case FEEDTIMEINDEX
                                If imDelOrEngr = 1 Then
                                    'If Trim$(tmDlf(ilIndex).DlfRec.sFed) = "Y" Then
                                        slStr = Trim$(tmDlf(ilIndex).sFeedTime)
                                        If slStr <> "" Then
                                            slStr = gFormatTime(slStr, "A", "1")
                                        End If
                                    'Else
                                    '    slStr = ""
                                    'End If
                                Else
                                    slStr = ""
                                End If
                            Case TIMEZONEINDEX
                                slStr = Trim$(tmDlf(ilIndex).DlfRec.sZone)
                            Case SUBFEEDINDEX
                                slStr = Trim$(tmDlf(ilIndex).sSubfeed)
                            Case EVTNAMEINDEX
                                slStr = Trim$(tmDlf(ilIndex).sEventName)
                            Case PROGCODEINDEX
                                slStr = Trim$(tmDlf(ilIndex).DlfRec.sProgCode)
                            Case SHOWONINDEX
                                gPaintArea pbcDelivery(imDelIndex), tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilPaintRow - 1) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
                                pbcDelivery(imDelIndex).CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
                                pbcDelivery(imDelIndex).CurrentY = tmCtrls(ilBox).fBoxY + (ilPaintRow - 1) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                                If imDelOrEngr = 0 Then
                                    If Trim$(tmDlf(ilIndex).DlfRec.sCmmlSched) = "Y" Then
                                        slStr = "Yes"
                                    ElseIf Trim$(tmDlf(ilIndex).DlfRec.sCmmlSched) = "N" Then
                                        slStr = "No"
                                    Else
                                        slStr = ""
                                    End If
                                Else
                                    slStr = ""
                                End If
                            Case FEDINDEX
                                If imDelOrEngr = 1 Then
                                    'If Trim$(tmDlf(ilIndex).DlfRec.sFed) = "Y" Then
                                        slStr = "Y"
                                    'Else
                                    '    slStr = "N"
                                    'End If
                                Else
                                    If Trim$(tmDlf(ilIndex).DlfRec.sFed) = "Y" Then
                                        slStr = "Y"
                                    Else
                                        slStr = "N"
                                    End If
                                End If
                            Case BUSINDEX
                                If imDelOrEngr = 1 Then
                                    'slStr = tmDlf(ilIndex).DlfRec.sBus
                                    slStr = ""
                                    For ilLoop = 1 To 5 Step 1
                                        slChar = Mid$(tmDlf(ilIndex).DlfRec.sBus, ilLoop, 1)
                                        If slChar = " " Then
                                            slStr = slStr & "-"
                                        Else
                                            slStr = slStr & slChar
                                        End If
                                    Next ilLoop
                                Else
                                    slStr = Trim$(tmDlf(ilIndex).DlfRec.sBus)
                                End If
                            Case SCHINDEX
                                If imDelOrEngr = 1 Then
                                    slStr = Trim$(tmDlf(ilIndex).DlfRec.sSchedule)
                                Else
                                    slStr = ""
                                End If
                        End Select
                        If tmDlf(ilIndex).iStatus = 0 Then  'Show new as MAGENTA
                            pbcDelivery(imDelIndex).ForeColor = MAGENTA
                        ElseIf tmDlf(ilIndex).iStatus = 2 Then  'Show Deleted as Cyan
                            pbcDelivery(imDelIndex).ForeColor = CYAN
                        Else
                            If (tmDlf(ilIndex).DlfRec.iEtfCode = 1) Then
                                pbcDelivery(imDelIndex).ForeColor = DARKGREEN
                            Else
                                pbcDelivery(imDelIndex).ForeColor = BLACK
                            End If
                        End If
                        gSetShow pbcDelivery(imDelIndex), slStr, tmCtrls(ilBox)
                        pbcDelivery(imDelIndex).Print tmCtrls(ilBox).sShow
                    Next ilBox
                    ilPaintRow = ilPaintRow + 1
                End If
            'End If
        End If
        ilEvt = ilEvt + 1
    Loop
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    Dim ilFound As Integer
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = -1  'Set-Right to left
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                If UBound(tmDlf) = 1 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imRowNo = 1
                If Not mFindRowIndex(imRowNo) Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imSettingValue = True
                vbcDelivery.Value = vbcDelivery.Min
                imSettingValue = False
                If imDelOrEngr = 0 Then
                    ilBox = LOCALTIMEINDEX
                Else
                    ilBox = FEEDTIMEINDEX
                End If
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            Case LOCALTIMEINDEX 'Name (first control within header)
                If imDelOrEngr = 0 Then
                    slStr = edcDropDown.Text
                    If Not gValidTime(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                End If
                mSetShow imBoxNo, False
                If imDelOrEngr = 0 Then
                    ilBox = FEDINDEX
                Else
                    ilBox = SCHINDEX
                End If
                If imRowNo <= 1 Then
                    imBoxNo = -1
                    imRowNo = -1
                    If cmcUpdate.Enabled Then
                        cmcUpdate.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                End If
                imRowNo = imRowNo - 1
                If Not mFindRowIndex(imRowNo) Then
                    imBoxNo = -1
                    imRowNo = -1
                    If cmcUpdate.Enabled Then
                        cmcUpdate.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                End If
                If imRowNo < vbcDelivery.Value Then
                    imSettingValue = True
                    vbcDelivery.Value = vbcDelivery.Value - 1
                    imSettingValue = False
                End If
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            Case FEEDTIMEINDEX
                If imDelOrEngr = 1 Then
                    slStr = edcDropDown.Text
                    If Not gValidTime(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                    ilFound = False
                End If
                ilBox = LOCALTIMEINDEX
            Case TIMEZONEINDEX
                If imDelOrEngr = 0 Then
                    ilFound = False
                End If
                ilBox = FEEDTIMEINDEX
                'If (Trim$(tmDlf(imDlfIndex).DlfRec.sFed) = "N") Then
                '    ilBox = LOCALTIMEINDEX
                'Else
                '    ilBox = FEEDTIMEINDEX
                'End If
            Case PROGCODEINDEX
                If tmMnf.iGroupNo = 1 Then
                    ilBox = SUBFEEDINDEX
                Else
                    ilBox = TIMEZONEINDEX
                End If
            Case SHOWONINDEX
                If imDelOrEngr = 1 Then
                    ilFound = False
                End If
                ilBox = PROGCODEINDEX
            Case FEDINDEX
                If imDelOrEngr = 1 Then
                    ilFound = False
                End If
                ilBox = SHOWONINDEX
            Case BUSINDEX
                If imDelOrEngr = 1 Then
                    ilFound = False
                    ilBox = FEDINDEX
                Else
                    If (tmDlf(imDlfIndex).DlfRec.iEtfCode = 1) Or (tmDlf(imDlfIndex).DlfRec.iEtfCode > 13) Then  'Program event type
                        ilBox = SHOWONINDEX
                    Else
                        ilBox = FEDINDEX
                    End If
                End If
            Case SCHINDEX
                'If imDelOrEngr = 0 Then
                '    ilFound = False
                'End If
                ilBox = BUSINDEX
            Case Else
                ilBox = ilBox - 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo, False
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    Dim slStr As String

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    If imDirProcess >= 0 Then
        mDirection imDirProcess
        imDirProcess = -1
        Exit Sub
    End If
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTabDirection = -1  'Set-Right to left
                imRowNo = imMaxRowNo 'UBound(tmDlf) - 1
                imSettingValue = True
                If imRowNo <= vbcDelivery.LargeChange + 1 Then
                    vbcDelivery.Value = 1
                Else
                    vbcDelivery.Value = imRowNo - vbcDelivery.LargeChange
                End If
                imSettingValue = False
                If imDelOrEngr = 0 Then
                    ilBox = BUSINDEX   'FEDINDEX    'PROGCODEINDEX
                Else
                    ilBox = SCHINDEX
                End If
            Case 0  'Arrow
                If imDelOrEngr = 0 Then
                    ilBox = LOCALTIMEINDEX
                Else
                    ilBox = FEEDTIMEINDEX
                End If
            Case LOCALTIMEINDEX
                If imDelOrEngr = 0 Then
                    slStr = edcDropDown.Text
                    If Not gValidTime(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                    ilFound = False
                End If
                'If (Trim$(tmDlf(imDlfIndex).DlfRec.sFed) = "N") Then
                    'ilBox = TIMEZONEINDEX
                'Else
                '    ilBox = FEEDTIMEINDEX
                'End If
                ilBox = FEEDTIMEINDEX
            Case FEEDTIMEINDEX
                If imDelOrEngr = 1 Then
                    slStr = edcDropDown.Text
                    If Not gValidTime(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                End If
                ilBox = TIMEZONEINDEX
            Case TIMEZONEINDEX
                If tmMnf.iGroupNo = 1 Then
                    ilBox = SUBFEEDINDEX
                Else
                    If imDelOrEngr = 1 Then
                        ilFound = False
                    End If
                    ilBox = PROGCODEINDEX
                End If
            Case SUBFEEDINDEX
                If imDelOrEngr = 1 Then
                    ilFound = False
                End If
                ilBox = PROGCODEINDEX
            Case PROGCODEINDEX
                If imDelOrEngr = 1 Then
                    ilFound = False
                End If
                ilBox = SHOWONINDEX
            Case SHOWONINDEX
                If imDelOrEngr = 1 Then
                    ilFound = False
                End If
                ilBox = FEDINDEX
            Case FEDINDEX
                'If imDelOrEngr = 0 Then
                '    ilFound = False
                'End If
                ilBox = BUSINDEX
            Case BUSINDEX
                If imDelOrEngr = 0 Then
                    ilFound = False
                End If
                ilBox = SCHINDEX
            Case SCHINDEX 'Last control
                mSetShow imBoxNo, False
                If mTestFields(imDlfIndex) = NO Then
                    mEnableBox imBoxNo
                    Exit Sub
                End If
                If imRowNo >= UBound(tmDlf) - 1 Then
                    imBoxNo = -1
                    imRowNo = -1
                    If cmcUpdate.Enabled Then
                        cmcUpdate.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                End If
                imRowNo = imRowNo + 1
                If Not mFindRowIndex(imRowNo) Then
                    imBoxNo = -1
                    imRowNo = -1
                    If cmcUpdate.Enabled Then
                        cmcUpdate.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                End If
                If imRowNo > vbcDelivery.Value + vbcDelivery.LargeChange Then
                    imSettingValue = True
                    vbcDelivery.Value = vbcDelivery.Value + 1
                    imSettingValue = False
                End If
                If imDelOrEngr = 0 Then
                    ilBox = LOCALTIMEINDEX
                Else
                    ilBox = FEEDTIMEINDEX
                End If
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            Case Else
                ilBox = ilBox + 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo, False
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    Select Case imBoxNo
                        Case LOCALTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                        Case FEEDTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                    End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If imBoxNo = SHOWONINDEX Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            tmDlf(imDlfIndex).DlfRec.sCmmlSched = "Y"
            imChg = True
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            tmDlf(imDlfIndex).DlfRec.sCmmlSched = "N"
            imChg = True
            pbcYN_Paint
        End If
    ElseIf imBoxNo = FEDINDEX Then
        If (tmDlf(imDlfIndex).DlfRec.iEtfCode = 1) Or (tmDlf(imDlfIndex).DlfRec.iEtfCode > 13) Then   'Program event type
            Beep
            Exit Sub
        End If
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            tmDlf(imDlfIndex).DlfRec.sFed = "Y"
            imChg = True
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            tmDlf(imDlfIndex).DlfRec.sFed = "N"
            imChg = True
            pbcYN_Paint
        End If
    End If
End Sub
Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = SHOWONINDEX Then
        If tmDlf(imDlfIndex).DlfRec.sCmmlSched = "Y" Then
            tmDlf(imDlfIndex).DlfRec.sCmmlSched = "N"
            imChg = True
        Else
            tmDlf(imDlfIndex).DlfRec.sCmmlSched = "Y"
            imChg = True
        End If
        pbcYN_Paint
    ElseIf imBoxNo = FEDINDEX Then
        If (tmDlf(imDlfIndex).DlfRec.iEtfCode = 1) Or (tmDlf(imDlfIndex).DlfRec.iEtfCode > 13) Then   'Program event type
            Beep
            Exit Sub
        End If
        If tmDlf(imDlfIndex).DlfRec.sFed = "Y" Then
            tmDlf(imDlfIndex).DlfRec.sFed = "N"
            imChg = True
        Else
            tmDlf(imDlfIndex).DlfRec.sFed = "Y"
            imChg = True
        End If
        pbcYN_Paint
    End If
End Sub
Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If imBoxNo = SHOWONINDEX Then
        If tmDlf(imDlfIndex).DlfRec.sCmmlSched = "Y" Then
            pbcYN.Print "Yes"
        Else
            pbcYN.Print "No"
            If tmDlf(imDlfIndex).DlfRec.sCmmlSched <> "N" Then  'Correct for invalid design- A or O wasn't set to N or Y
                tmDlf(imDlfIndex).DlfRec.sCmmlSched = "N"
                imChg = True
            End If
        End If
    ElseIf imBoxNo = FEDINDEX Then
        If tmDlf(imDlfIndex).DlfRec.sFed = "Y" Then
            pbcYN.Print "Yes"
        Else
            pbcYN.Print "No"
        End If
    End If
End Sub
Private Sub plcDelivery_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub plcSort_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub rbcSort_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSort(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    If Value Then
        Screen.MousePointer = vbHourglass
        For ilLoop = UBound(imSort) - 1 To LBound(imSort) Step -1
            imSort(ilLoop + 1) = imSort(ilLoop)
        Next ilLoop
        imSort(LBound(imSort)) = Index
        mSort
        pbcDelivery(imDelIndex).Cls
'        pbcDelivery_Paint
        If vbcDelivery.Value <> vbcDelivery.Min Then
            vbcDelivery.Value = vbcDelivery.Min
        Else
            vbcDelivery_Change
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub rbcSort_GotFocus(Index As Integer)
    mSetShow imBoxNo, False
    imRowNo = -1
    imBoxNo = -1
    lacFrame(imDelIndex).Visible = False
    pbcArrow.Visible = False
End Sub
Private Sub rbcSort_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            ilCompRow = vbcDelivery.LargeChange + 1
            If imMaxRowNo > ilCompRow Then
                ilMaxRow = ilCompRow
            Else
                ilMaxRow = imMaxRowNo
            End If
            For ilRow = 1 To ilMaxRow Step 1
                If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(VEHICLEINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(VEHICLEINDEX).fBoxY + tmCtrls(VEHICLEINDEX).fBoxH)) Then
                    mSetShow imBoxNo, False
                    imBoxNo = -1
                    imRowNo = -1
                    imRowNo = ilRow + vbcDelivery.Value - 1
                    If Not mFindRowIndex(imRowNo) Then
                        pbcArrow.Visible = False
                        lacFrame(imDelIndex).Visible = False
                        Exit Sub
                    End If
                    lacFrame(imDelIndex).Move 0, tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) - 30
                    lacFrame(imDelIndex).Visible = True
                    pbcArrow.Move pbcArrow.Left, plcDelivery.Top + tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcDelivery.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    lacFrame(imDelIndex).Drag vbBeginDrag
                    Exit Sub
                End If
            Next ilRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub vbcDelivery_Change()

    Screen.MousePointer = vbHourglass
    If imSettingValue Then
        pbcDelivery(imDelIndex).Cls
        pbcDelivery_Paint imDelIndex
        imSettingValue = False
    Else
        mSetShow imBoxNo, False
        pbcDelivery(imDelIndex).Cls
        pbcDelivery_Paint imDelIndex
        mEnableBox imBoxNo
    End If
    vbcDelivery_Scroll  'set Scroll box value
    Screen.MousePointer = vbDefault
End Sub
Private Sub vbcDelivery_GotFocus()
    mSetShow imBoxNo, False
    imRowNo = -1
    imBoxNo = -1
    lacFrame(imDelIndex).Visible = False
    pbcArrow.Visible = False
End Sub
Private Sub vbcDelivery_Scroll()
    Dim ilIndex As Integer
    Dim ilSort As Integer
    Dim slStr As String
    Dim slShow As String
    Dim ilRow As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slChar As String
    ilRow = 0
    ilFound = False
    For ilLoop = 0 To lbcSortIndex.ListCount - 1 Step 1
        ilIndex = lbcSortIndex.ItemData(ilLoop)
        'If (tmDlf(ilIndex).iStatus = 0) Or (tmDlf(ilIndex).iStatus = 1) Then
        If (tmDlf(ilIndex).iStatus = 0) Or (tmDlf(ilIndex).iStatus = 1) Or (tmDlf(ilIndex).iStatus = 2) Then
            'If (tmDlf(ilIndex).DlfRec.sFed = "Y") Or (Not ckcFedEvtOnly.Value) Then
                ilRow = ilRow + 1
                If ilRow = vbcDelivery.Value Then
                    ilFound = True
                    Exit For
                End If
            'End If
        End If
    Next ilLoop
    If Not ilFound Then
        plcScroll.Caption = ""
        smScrollCaption = ""
        'plcScroll.Cls
        Exit Sub
    End If
    slShow = ""
    For ilSort = UBound(imSort) To LBound(imSort) Step -1
        Select Case imSort(ilSort)
            Case 0  'Vehicle
                If imDelIndex = 0 Then
                    slStr = Left$(tmDlf(ilIndex).sVehicle, 6)
                Else
                    slStr = Left$(tmDlf(ilIndex).sFeed, 6)
                End If
            Case 1  'Air time
                slStr = Trim$(tmDlf(ilIndex).sAirTime)
                If slStr <> "" Then
                    slStr = gFormatTime(slStr, "A", "1")
                End If
            Case 2  'Local time
                slStr = Trim$(tmDlf(ilIndex).sLocalTime)
                If slStr <> "" Then
                    slStr = gFormatTime(slStr, "A", "1")
                End If
            Case 3  'Feed time
                slStr = Trim$(tmDlf(ilIndex).sFeedTime)
                If slStr <> "" Then
                    slStr = gFormatTime(slStr, "A", "1")
                End If
            Case 4  'Zone
                slStr = Trim$(tmDlf(ilIndex).DlfRec.sZone)
            Case 5  'Event Type
                slStr = Trim$(tmDlf(ilIndex).sEventType)
            Case 6  'Event name
                slStr = Trim$(tmDlf(ilIndex).sEventName)
            Case 7  'Program code
                slStr = Trim$(tmDlf(ilIndex).DlfRec.sProgCode)
            Case 8  'Cmml schd
                If Trim$(tmDlf(ilIndex).DlfRec.sCmmlSched) = "Y" Then
                    slStr = "Yes"
                Else
                    slStr = "No"
                End If
            Case 9  'Bus
                'slStr = tmDlf(ilIndex).DlfRec.sBus
                slStr = ""
                For ilLoop = 1 To 5 Step 1
                    slChar = Mid$(tmDlf(ilIndex).DlfRec.sBus, ilLoop, 1)
                    If slChar = " " Then
                        slStr = slStr & "-"
                    Else
                        slStr = slStr & slChar
                    End If
                Next ilLoop
            Case 10 'Schedule
                slStr = Trim$(tmDlf(ilIndex).DlfRec.sSchedule)
        End Select
        If slShow = "" Then
            slShow = slStr
        Else
            slShow = slStr & "/" & slShow
        End If
    Next ilSort
    smScrollCaption = slShow
    plcScroll.Caption = " " & smScrollCaption
End Sub
Private Sub plcSort_Paint()
    plcSort.CurrentX = 0
    plcSort.CurrentY = 0
    plcSort.Print "Sort by"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub

Private Sub mAbortSaveRec()
    Dim ilRet As Integer
    
    ilRet = btrAbortTrans(hmDlf)
    Screen.MousePointer = vbDefault
    ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Links")
    imTerminate = True

End Sub
