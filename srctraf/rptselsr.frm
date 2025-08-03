VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelSR 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5535
   ClientLeft      =   195
   ClientTop       =   1545
   ClientWidth     =   9270
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
   ScaleHeight     =   5535
   ScaleWidth      =   9270
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6600
      TabIndex        =   18
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8145
      TabIndex        =   9
      Top             =   -15
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmcDone 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6675
      Top             =   -180
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
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
      Height          =   285
      Left            =   7215
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   -15
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
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
      Height          =   285
      Left            =   7575
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   -60
      Visible         =   0   'False
      Width           =   525
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
      ScaleWidth      =   30
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4140
      Top             =   4875
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".Txt"
      Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
      FilterIndex     =   1
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Frame frcCopies 
      Caption         =   "Printing"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2070
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.TextBox edcCopies 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1065
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frcFile 
      Caption         =   "Save to File"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2070
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox cbcFileType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   780
         TabIndex        =   12
         Top             =   270
         Width           =   2925
      End
      Begin VB.TextBox edcFileName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   780
         TabIndex        =   14
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Copy Regions"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3990
      Left            =   45
      TabIndex        =   16
      Top             =   1515
      Width           =   9090
      Begin VB.PictureBox pbcSelC 
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
         Height          =   3360
         Left            =   90
         ScaleHeight     =   3360
         ScaleWidth      =   4620
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   4620
         Begin VB.PictureBox plcDetail 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2220
            ScaleHeight     =   240
            ScaleWidth      =   1965
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1290
            Visible         =   0   'False
            Width           =   1965
            Begin VB.CheckBox ckcInclStations 
               Caption         =   "Include Stations"
               Height          =   255
               Left            =   0
               TabIndex        =   57
               Top             =   15
               Width           =   1695
            End
         End
         Begin VB.PictureBox plcCodeSort 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   15
            ScaleHeight     =   480
            ScaleWidth      =   4275
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   2850
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcCodeSort 
               Caption         =   "Advertiser"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   52
               Top             =   0
               Width           =   1305
            End
            Begin VB.OptionButton rbcCodeSort 
               Caption         =   "Region Name"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2010
               TabIndex        =   51
               Top             =   0
               Value           =   -1  'True
               Width           =   1425
            End
            Begin VB.OptionButton rbcCodeSort 
               Caption         =   "Region Code"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   705
               TabIndex        =   50
               Top             =   225
               Width           =   1830
            End
         End
         Begin VB.TextBox lacCodeTo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   0
            TabIndex        =   48
            Text            =   "End Region Code"
            Top             =   2610
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox lacCodeFrom 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   30
            TabIndex        =   47
            Text            =   "Start Region Code"
            Top             =   2370
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox lacDateTo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   15
            TabIndex        =   46
            Text            =   "Creation End Date"
            Top             =   435
            Width           =   1680
         End
         Begin VB.TextBox lacDateFrom 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   0
            TabIndex        =   45
            Text            =   "Creation Start Date"
            Top             =   75
            Width           =   1695
         End
         Begin VB.TextBox edcCodeTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2805
            MaxLength       =   10
            TabIndex        =   44
            Top             =   2520
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox edcCodeFrom 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1635
            MaxLength       =   10
            TabIndex        =   43
            Top             =   2565
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox edcDateTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1950
            MaxLength       =   10
            TabIndex        =   42
            Top             =   375
            Width           =   1080
         End
         Begin VB.TextBox edcDateFrom 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1950
            MaxLength       =   10
            TabIndex        =   41
            Top             =   15
            Width           =   1080
         End
         Begin VB.CheckBox ckcDormant 
            Caption         =   "Include Dormant"
            Height          =   255
            Left            =   15
            TabIndex        =   38
            Top             =   1305
            Width           =   1875
         End
         Begin VB.PictureBox plcCategory 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   0
            ScaleHeight     =   600
            ScaleWidth      =   4380
            TabIndex        =   31
            Top             =   1860
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcCat 
               Caption         =   "Format"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   840
               TabIndex        =   32
               Top             =   0
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.CheckBox ckcCat 
               Caption         =   "Time Zone"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   3090
               TabIndex        =   37
               Top             =   240
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox ckcCat 
               Caption         =   "Station"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   1920
               TabIndex        =   36
               Top             =   240
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.CheckBox ckcCat 
               Caption         =   "State"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   840
               TabIndex        =   35
               Top             =   240
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.CheckBox ckcCat 
               Caption         =   "MSA Mkt"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   3090
               TabIndex        =   34
               Top             =   0
               Value           =   1  'Checked
               Width           =   1320
            End
            Begin VB.CheckBox ckcCat 
               Caption         =   "DMA Mkt"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1935
               TabIndex        =   33
               Top             =   0
               Value           =   1  'Checked
               Width           =   1140
            End
         End
         Begin VB.PictureBox plcWhichSplit 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            ScaleHeight     =   240
            ScaleWidth      =   4275
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   765
            Width           =   4275
            Begin VB.OptionButton rbcWhichSplit 
               Caption         =   "Both"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3120
               TabIndex        =   30
               Top             =   0
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.OptionButton rbcWhichSplit 
               Caption         =   "Copy"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2280
               TabIndex        =   29
               Top             =   0
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.OptionButton rbcWhichSplit 
               Caption         =   "Network"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   1200
               TabIndex        =   28
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox plcSortBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            ScaleHeight     =   240
            ScaleWidth      =   4275
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1020
            Width           =   4275
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Station"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   2640
               TabIndex        =   25
               Top             =   0
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Region Definitions"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   720
               TabIndex        =   26
               Top             =   0
               Width           =   1860
            End
         End
      End
      Begin VB.PictureBox pbcOption 
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
         Height          =   3795
         Left            =   4605
         ScaleHeight     =   3795
         ScaleWidth      =   4455
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4455
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   3
            ItemData        =   "rptselsr.frx":0000
            Left            =   225
            List            =   "rptselsr.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   59
            Top             =   525
            Visible         =   0   'False
            Width           =   4005
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1080
            Index           =   2
            ItemData        =   "rptselsr.frx":0004
            Left            =   270
            List            =   "rptselsr.frx":0006
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   58
            Top             =   2475
            Width           =   4005
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   1
            ItemData        =   "rptselsr.frx":0008
            Left            =   255
            List            =   "rptselsr.frx":000A
            MultiSelect     =   2  'Extended
            TabIndex        =   54
            Top             =   2250
            Width           =   4005
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Advertisers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2565
            TabIndex        =   39
            Top             =   45
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   0
            ItemData        =   "rptselsr.frx":000C
            Left            =   240
            List            =   "rptselsr.frx":0013
            TabIndex        =   40
            Top             =   360
            Width           =   4005
         End
         Begin VB.Label lacRegions 
            Caption         =   "Regions"
            Height          =   270
            Left            =   255
            TabIndex        =   56
            Top             =   1965
            Width           =   1095
         End
         Begin VB.Label lacAdvt 
            Caption         =   "Advertisers"
            Height          =   270
            Left            =   255
            TabIndex        =   55
            Top             =   90
            Width           =   1095
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   19
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   6
      Top             =   105
      Width           =   2805
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   1890
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Save to File"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   1395
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Print"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   555
         Width           =   750
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8850
      Top             =   1035
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RptSelSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptselsr.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  smLogUserCode                                                                         *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelSR.Frm  - Split Regions
'
' Release: 5.5
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
Dim imSetAll As Integer
Dim imAllClicked As Integer


Dim imTerminate As Integer

'Dim tmSRec As LPOPREC
'Rate Card
Private Sub cbcFileType_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcFileType.Text <> "" Then
            gManLookAhead cbcFileType, imBSMode, imComboBoxIndex
        End If
        imFTSelectedIndex = cbcFileType.ListIndex
        imChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcFileType_Click()
    imComboBoxIndex = cbcFileType.ListIndex
    imFTSelectedIndex = cbcFileType.ListIndex
    mSetCommands
End Sub
Private Sub cbcFileType_GotFocus()
    If cbcFileType.Text = "" Then
        cbcFileType.ListIndex = 0
    End If
    imComboBoxIndex = cbcFileType.ListIndex
    gCtrlGotFocus cbcFileType
End Sub
Private Sub cbcFileType_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcFileType_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcFileType.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub


Private Sub ckcAll_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLbcIndex                                                                            *
'******************************************************************************************

 'Code added because Value removed as parameter
    Dim Value As Integer

    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    Dim ilIndex As Integer
    ilIndex = lbcRptType.ListIndex          'report index

    ilValue = Value
    If imSetAll Then
        If (Asc(tgSpf.sUsingFeatures2) And REGIONALCOPY) = REGIONALCOPY Then        'using regional copy (vs splits)
            imAllClicked = True
            llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(3).hwnd, LB_SELITEMRANGE, ilValue, llRg)
            imAllClicked = False
        End If
    Else
        imAllClicked = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If

    mSetCommands
End Sub

Private Sub ckcDormant_Click()
Dim ilIncludeDormant As Integer
Dim slNameCode As String
Dim ilRet As Integer
Dim ilAdfCode As Integer
Dim slCode As String

    ilIncludeDormant = False
    If ckcDormant.Value = vbChecked Then
        ilIncludeDormant = True
    End If
    
    If rbcWhichSplit(0).Value = True Then           'network option
        sgRegionCodeTag = ""            'make sure to repopulate with/without dormant regions
        ilRet = gPopRegionBox(RptSelSR, 0, "N", ilIncludeDormant, lbcSelection(2), tgRegionCode(), sgRegionCodeTag)
    Else
        If lbcSelection(0).SelCount > 0 Then
            'If lbcSelection(0).ListIndex >= 1 Then
                slNameCode = tgAdvertiser(lbcSelection(0).ListIndex).sKey   'Traffic!lbcAdvertiser.List(lbcAdvt.ListIndex - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilAdfCode = Val(Trim$(slCode))
    
                'ALL option no longer allowed, single advt only
                'once an advt is selectec, populate the regions for that adv
                sgRegionCodeTag = ""            'make sure to repopulate with/without dormant regions
                ilRet = gPopRegionBox(RptSelSR, ilAdfCode, "C", ilIncludeDormant, lbcSelection(1), tgRegionCode(), sgRegionCodeTag)
        '        If Not imAllClicked Then
        '            imSetAll = False
        '            ckcAll.Value = vbUnchecked  '12-11-01 False
        '            imSetAll = True
        '        Else
                    'imSetAll = False
                    'ckcAll.Value = False
                    'imSetAll = True
            'End If
        End If
    End If
    mSetCommands
End Sub

Private Sub cmcBrowse_Click()
    cdcSetup.flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
    cdcSetup.fileName = edcFileName.Text
    cdcSetup.InitDir = Left$(sgRptSavePath, Len(sgRptSavePath) - 1)
    cdcSetup.Action = 2    'DLG_FILE_SAVE
    edcFileName.Text = cdcSetup.fileName
    mSetCommands
    gChDrDir        '3-25-03
    'ChDrive Left$(sgCurDir, 2)  'Set the default drive
    'ChDir sgCurDir              'set the default path
End Sub
Private Sub cmcBrowse_GotFocus()
    gCtrlGotFocus cmcBrowse
End Sub
Private Sub cmcCancel_Click()
    If igGenRpt Then
        Exit Sub
    End If
    mTerminate False
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcGen_Click()
    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim ilListIndex As Integer
    If igGenRpt Then
        Exit Sub
    End If
    igGenRpt = True
    igOutput = frcOutput.Enabled
    igCopies = frcCopies.Enabled
    'igWhen = frcWhen.Enabled
    igFile = frcFile.Enabled
    igOption = frcOption.Enabled
    'igReportType = frcRptType.Enabled
    frcOutput.Enabled = False
    frcCopies.Enabled = False
    'frcWhen.Enabled = False
    frcFile.Enabled = False
    frcOption.Enabled = False
    'frcRptType.Enabled = False
    ilListIndex = lbcRptType.ListIndex

    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs

        If Not gGenReportSR(ilListIndex) Then       'open report
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If
        ilRet = gCmcGenSR(ilListIndex)              ' formulas
        '-1 is a Crystal failure of gSetSelection or gSEtFormula
        If ilRet = -1 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            'mTerminate
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
        ElseIf ilRet = 0 Then   '0 = invalid input data, stay in
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        ElseIf ilRet = 2 Then           'successful return from bridge reports
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
        End If
       '1 falls thru - successful crystal report

        If ilListIndex = CR_SPLITREGION Or ilListIndex = CR_COPYREGION Then             'split region rept
            Screen.MousePointer = vbHourglass
            gCreateSplitRegions
            Screen.MousePointer = vbDefault
        End If

        If rbcOutput(0).Value Then
            DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
            igDestination = 0
            Report.Show vbModal
        ElseIf rbcOutput(1).Value Then
            ilCopies = Val(edcCopies.Text)
            ilRet = gOutputToPrinter(ilCopies)
        Else
            slFileName = edcFileName.Text
           ' ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
        End If
    Next ilJobs
    imGenShiftKey = 0

    Screen.MousePointer = vbHourglass
    gCRGrfClear
    Screen.MousePointer = vbDefault

    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    'frcWhen.Enabled = igWhen
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    Exit Sub
End Sub
Private Sub cmcGen_GotFocus()
    gCtrlGotFocus cmcGen
End Sub
Private Sub cmcGen_KeyDown(KeyCode As Integer, Shift As Integer)
    imGenShiftKey = Shift
End Sub
Private Sub cmcList_Click()
    If igGenRpt Then
        Exit Sub
    End If
    mTerminate True
End Sub
Private Sub cmcSetup_Click()
    'cdcSetup.Flags = cdlPDPrintSetup
    'cdcSetup.Action = 5    'DLG_PRINT
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub
Private Sub edcCopies_Change()
    mSetCommands
End Sub
Private Sub edcCopies_GotFocus()
    gCtrlGotFocus edcCopies
End Sub
Private Sub edcCopies_KeyPress(KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcDateFrom_GotFocus()
    gCtrlGotFocus edcDateFrom
End Sub

Private Sub edcDateTo_GotFocus()
    gCtrlGotFocus edcDateTo
End Sub

Private Sub edcFileName_Change()
    mSetCommands
End Sub
Private Sub edcFileName_GotFocus()
    gCtrlGotFocus edcFileName
End Sub
Private Sub edcFileName_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer

    ilPos = InStr(edcFileName.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcFileName.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    If ((KeyAscii <> KEYBACKSPACE) And (KeyAscii <= 32)) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KEYDOWN) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    RptSelSR.Refresh
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
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
    imTerminate = False
    igGenRpt = False
    mParseCmmdLine
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    mInit
    If imTerminate Then 'Used for print only
        mTerminate True
        Exit Sub
    End If
    'RptSelSR.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    PECloseEngine
    
    Set RptSelSR = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/16/93       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*            Place focus before populating all lists  *                                                   *
'*******************************************************
Private Sub mInit()
Dim ilRet As Integer
Dim ilLoop As Integer
Dim slStr As String
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    'Start Crystal report engine
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        Exit Sub
    End If
    'Set options for report generate
    'hdJob = rpcRpt.hJob
    'ilMultiTable = True
    'ilDummy = LlSetOption(hdJob, LL_OPTION_SORTVARIABLES, True)
    'ilDummy = LlSetOption(hdJob, LL_OPTION_ONLYONETABLE, Not ilMultiTable)

    RptSelSR.Caption = smSelectedRptName & " Report"
    'frcOption.Caption = smSelectedRptName & " Selection"
    slStr = Trim$(smSelectedRptName)
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    gCenterStdAlone RptSelSR
    
    sgAdvertiserTag = ""
    ilRet = gRptAdvtPop(RptSelSR, lbcSelection(0))      'single select advt list box for split copy and/or split networks
    sgAdvertiserTag = ""
    ilRet = gRptAdvtPop(RptSelSR, lbcSelection(3))      'multi select advt list box for Regional Copy
    If ilRet = True Then
        imTerminate = True
        Exit Sub
    End If
    imAllClicked = False
    imSetAll = True


End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitReport                     *
'*                                                     *
'*             Created:8/16/00       By:D. Smith       *
'*             Modified:             By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*                                                     *
'*******************************************************
Private Sub mInitReport()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilDay                         llDate                        slNowDate                 *
'*                                                                                        *
'******************************************************************************************




    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    gPopExportTypes cbcFileType         '10-20-01


    '7-8-10  Copy Regions report will vary based on Site Options.
    'if Site Option:  Using Regional Copy is checked (CR_COPYREGION), it then generates the RAF.rpt report
    'If Site Option:  Using Split Networks or Split Copy is checked (CR_SPLITREGION),
    '                 it generates copysplit.rpt or copysplitstn.rpt
    '
    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60

    pbcSelC.Visible = True
    pbcOption.Visible = True
    lbcRptType.AddItem "Copy Regions", CR_COPYREGION  '2-12-09 changed from Split Regions
    'lbcRptType.AddItem "Copy Split Region", CR_COPYREGION  '7-7-10 changed from Split Regions
    lbcRptType.AddItem "Split Network Avails", CR_SPLITREGION

    If lbcRptType.ListCount > 0 Then
        If (Asc(tgSpf.sUsingFeatures2) And REGIONALCOPY) = REGIONALCOPY Then        'using regional copy
            lbcRptType.Selected(CR_COPYREGION) = True
            smSelectedRptName = "Copy Regions"
            'smSelectedRptName = "Copy Split Region"
        Else
            lbcRptType.Selected(CR_SPLITREGION) = True
            smSelectedRptName = "Split Network Avails"
        End If
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
    End If
    mSetCommands
    Screen.MousePointer = vbDefault
    'gCenterModalForm RptSel





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
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slRptListCmmd As String

    slCommand = sgCommandStr    'Command$
    ilRet = gParseItem(slCommand, 1, "||", smCommand)
    If (ilRet <> CP_MSG_NONE) Then
        smCommand = slCommand
    End If
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False  'True 'False
    '    imShowHelpmsg = False
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
    '    imShowHelpmsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpmsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptSelSR, slStr, ilTestSystem
    'ilRet = gParseItem(slCommand, 3, "\", slStr)
    'igRptCallType = Val(slStr)
    If igRptCallType <> GENERICBUTTON Then
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        igRptType = Val(slStr)
    Else
        igRptType = 0
    End If
    'If igStdAloneMode Then
    '    smSelectedRptName = "Sales Pricing Analysis"
    '    igRptCallType = -1  'unused in standalone exe
    '    igRptType = -1   'unused in standalone exe
    'Else
        ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
        If (ilRet = CP_MSG_NONE) Then
            ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
            ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
            igRptCallType = Val(slStr)
        End If
    'End If

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/21/93       By:D. LeVine      *
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
    Dim ilListIndex As Integer
    Dim ilEnable As Integer
    
    ilListIndex = lbcRptType.ListIndex
    ilEnable = False

    If ilListIndex = CR_SPLITREGION Then
        If rbcWhichSplit(1).Value = True Then           'copy splits
            If lbcSelection(0).SelCount > 0 And lbcSelection(1).SelCount > 0 Then
                ilEnable = True
            End If
        Else                                            'network splits
            If lbcSelection(2).SelCount > 0 Then
                ilEnable = True
            End If
        End If
    Else
        ilEnable = True
    End If
    cmcGen.Enabled = ilEnable
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate(ilFromCancel As Integer)
'
'   mTerminate
'   Where:
'

    If ilFromCancel Then
        igRptReturn = True
    Else
        igRptReturn = False
    End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RptSelSR
    igManUnload = NO
End Sub

Private Sub lbcRptType_Click()
Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If ilListIndex = CR_COPYREGION Then         'using regional copy (not splits)
        plcWhichSplit.Visible = False
        plcSortBy.Visible = False
        plcCategory.Visible = False
        ckcDormant.Visible = False
        plcDetail.Visible = False
        edcCodeFrom.Move edcDateTo.Left, edcDateTo.Top + edcDateTo.Height + 60
        edcCodeTo.Move edcCodeFrom.Left, edcCodeFrom.Top + edcCodeFrom.Height + 60
        lacCodeFrom.Move 60, edcCodeFrom.Top + 60
        lacCodeTo.Move 60, edcCodeTo.Top + 60
        plcCodeSort.Move 60, edcCodeTo.Top + edcCodeTo.Height + 60
        ckcDormant.Move 60, plcCodeSort.Top + plcCodeSort.Height
        ckcDormant.Visible = True
        edcCodeFrom.Visible = True
        edcCodeTo.Visible = True
        lacCodeFrom.Visible = True
        lacCodeTo.Visible = True
        plcCodeSort.Visible = True
        lacDateFrom.Visible = True
        lacDateTo.Visible = True
        edcDateFrom.Visible = True
        edcDateTo.Visible = True
        lbcSelection(3).Move 270, 360, 4005, 3390       'multi-select advertisers
        lbcSelection(3).Visible = True
        lbcSelection(0).Visible = False     'single select advt
        lbcSelection(1).Visible = False     'hide region list box for the split regions
        lbcSelection(2).Visible = False     'hide region list box for split networks
        lacRegions.Visible = False
        lacAdvt.Visible = False
        ckcAll.Move 270, 45
        ckcAll.Visible = True
    ElseIf ilListIndex = CR_SPLITREGION Then     'using Copy Split regions and/or Split Networks
        plcWhichSplit.Move 60, 0
        plcSortBy.Move 60, plcWhichSplit.Top + plcWhichSplit.Height
        'plcDetail.Move 60, plcSortBy.Top + plcSortBy.Height         'hidden for now
        ckcDormant.Move 60, plcSortBy.Top + plcSortBy.Height
        lacDateFrom.Visible = False
        lacDateTo.Visible = False
        edcDateFrom.Visible = False
        edcDateTo.Visible = False
        lbcSelection(2).Visible = False     'network region list box
    End If

End Sub

Private Sub lbcSelection_Click(Index As Integer)
Dim ilIncludeDormant As Integer
Dim slNameCode As String
Dim ilRet As Integer
Dim slCode As String
Dim ilAdfCode As Integer
Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex

    ilIncludeDormant = False
    If ckcDormant.Value = vbChecked Then
        ilIncludeDormant = True
    End If

    If (Asc(tgSpf.sUsingFeatures2) And REGIONALCOPY) = REGIONALCOPY Then        'using regional copy (vs splits)
        If Not imAllClicked Then
            imSetAll = False
            ckcAll.Value = vbUnchecked  '12-11-01 False
            imSetAll = True
        Else
            imSetAll = False
            ckcAll.Value = False
            imSetAll = True
        End If
    Else                                    'split copy & split networks
        If ilListIndex = 0 Then             'network version
            ilListIndex = ilListIndex
        Else                                'copy version
            If Index = 0 Then               'selecting an advertiser, need to populate the regions for that advt
                'If lbcSelection(0).ListIndex >= 1 Then
                    slNameCode = tgAdvertiser(lbcSelection(0).ListIndex).sKey   'Traffic!lbcAdvertiser.List(lbcAdvt.ListIndex - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilAdfCode = Val(Trim$(slCode))
        
                    'ALL option no longer allowed, single advt only
                    'once an advt is selectec, populate the regions for that adv
                    sgRegionCodeTag = ""
                    ilRet = gPopRegionBox(RptSelSR, ilAdfCode, "C", ilIncludeDormant, lbcSelection(1), tgRegionCode(), sgRegionCodeTag)
            '        If Not imAllClicked Then
            '            imSetAll = False
            '            ckcAll.Value = vbUnchecked  '12-11-01 False
            '            imSetAll = True
            '        Else
                        'imSetAll = False
                        'ckcAll.Value = False
                        'imSetAll = True
                'End If
        
            End If
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub


Private Sub plcCategory_Paint()
    plcCategory.Cls
    plcCategory.CurrentX = 0
    plcCategory.CurrentY = 0
    plcCategory.Print "Include"
End Sub

Private Sub plcCodeSort_Paint()
    plcCodeSort.Cls
    plcCodeSort.CurrentX = 0
    plcCodeSort.CurrentY = 0
    plcCodeSort.Print "Sort by"
End Sub

Private Sub plcDetail_Paint()
    plcDetail.Cls
    plcDetail.CurrentX = 0
    plcDetail.CurrentY = 0
    plcDetail.Print "Show"
End Sub

Private Sub plcSortBy_Paint()
    plcSortBy.Cls
    plcSortBy.CurrentX = 0
    plcSortBy.CurrentY = 0
    plcSortBy.Print "Sort by"
End Sub

Private Sub plcWhichSplit_Paint()
    plcWhichSplit.Cls
    plcWhichSplit.CurrentX = 0
    plcWhichSplit.CurrentY = 0
    plcWhichSplit.Print "Include Split"
End Sub

Private Sub rbcOutput_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcOutput(Index).Value
    'End of coded added
    If rbcOutput(Index).Value Then
        Select Case Index
            Case 0  'Display
                frcFile.Enabled = False
                frcCopies.Visible = False   'Print Box
                frcFile.Visible = False     'Save to File Box
                frcCopies.Enabled = False
                'frcWhen.Enabled = False
                'pbcWhen.Visible = False
            Case 1  'Print
                frcFile.Visible = False
                frcFile.Enabled = False
                frcCopies.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcCopies.Visible = True
            Case 2  'File
                frcCopies.Visible = False
                frcCopies.Enabled = False
                frcFile.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcFile.Visible = True
        End Select
    End If
    mSetCommands
End Sub
Private Sub rbcOutput_GotFocus(Index As Integer)
    If imFirstTime Then
        mInitReport
        If imTerminate Then 'Used for print only
            mTerminate True
            Exit Sub
        End If
        imFirstTime = False
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub

Private Sub rbcWhichSplit_Click(Index As Integer)
Dim ilRet As Integer
Dim ilIncludeDormant As Integer
Dim ilTemp As Integer
    If Index = 0 Then           'split network, no advt association
        ilIncludeDormant = False
        If ckcDormant.Value = vbChecked Then
            ilIncludeDormant = True
        End If
        ckcAll.Visible = False
        lbcSelection(0).Visible = False 'hide advertiser list box
        lbcSelection(1).Visible = False 'hide copy region list box
        lbcSelection(2).Visible = True  'Network region list box
        lacAdvt.Visible = False
        ckcAll.Value = vbChecked
        'rbcSortBy(1).Value = True
        'rbcSortBy(0).Enabled = False
        sgRegionCodeTag = ""
        ilRet = gPopRegionBox(RptSelSR, 0, "N", ilIncludeDormant, lbcSelection(2), tgRegionCode(), sgRegionCodeTag)
        lacRegions.Move 255, 90
        lbcSelection(2).Move 270, 360, 4005, 3390
    Else
        ckcAll.Value = vbUnchecked
        'ckcAll.Visible = True              'all advt disabled
        If lbcSelection(0).SelCount > 0 Then            'something already selected, repopulate regions with the same advt regions (changed from option Network to Copy
            For ilTemp = 0 To lbcSelection(0).ListCount - 1
                If lbcSelection(0).Selected(ilTemp) = True Then
                    lbcSelection_Click 0
                    Exit For
                End If
            Next ilTemp
        End If
        lbcSelection(0).Visible = True
        lbcSelection(1).Visible = True
        lacAdvt.Visible = True
        lacRegions.Visible = True
        rbcSortBy(0).Enabled = True
        lbcSelection(2).Visible = False     'hide the Network region list box
        lacRegions.Move 255, 1965
    End If
    mSetCommands
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
