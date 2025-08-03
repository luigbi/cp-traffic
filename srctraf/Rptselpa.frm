VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelPA 
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
      TabIndex        =   38
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8145
      TabIndex        =   29
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
      TabIndex        =   45
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
      TabIndex        =   44
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
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   855
      Top             =   4770
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
         TabIndex        =   27
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   28
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
      TabIndex        =   30
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   35
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
         TabIndex        =   32
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
         TabIndex        =   34
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Sales Pricing Analysis Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4050
      Left            =   45
      TabIndex        =   36
      Top             =   1455
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
         Height          =   3675
         Left            =   90
         ScaleHeight     =   3675
         ScaleWidth      =   4890
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4890
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   30
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Text            =   "9/11/19"
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BorderStyle     =   1
            CSI_ShowDropDownOnFocus=   0   'False
            CSI_InputBoxBoxAlignment=   0
            CSI_CalBackColor=   16777130
            CSI_CalDateFormat=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CSI_CurDayBackColor=   16777215
            CSI_CurDayForeColor=   51200
            CSI_ForceMondaySelectionOnly=   0   'False
            CSI_AllowBlankDate=   0   'False
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom1 
            Height          =   315
            Left            =   1320
            TabIndex        =   8
            Top             =   420
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Text            =   "9/11/19"
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BorderStyle     =   1
            CSI_ShowDropDownOnFocus=   0   'False
            CSI_InputBoxBoxAlignment=   0
            CSI_CalBackColor=   16777130
            CSI_CalDateFormat=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CSI_CurDayBackColor=   16777215
            CSI_CurDayForeColor=   51200
            CSI_ForceMondaySelectionOnly=   0   'False
            CSI_AllowBlankDate=   -1  'True
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalTo 
            Height          =   315
            Left            =   3000
            TabIndex        =   7
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Text            =   "9/11/19"
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BorderStyle     =   1
            CSI_ShowDropDownOnFocus=   0   'False
            CSI_InputBoxBoxAlignment=   0
            CSI_CalBackColor=   16777130
            CSI_CalDateFormat=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CSI_CurDayBackColor=   16777215
            CSI_CurDayForeColor=   51200
            CSI_ForceMondaySelectionOnly=   0   'False
            CSI_AllowBlankDate=   0   'False
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalTo1 
            Height          =   315
            Left            =   3000
            TabIndex        =   9
            Top             =   420
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Text            =   "9/11/19"
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BorderStyle     =   1
            CSI_ShowDropDownOnFocus=   0   'False
            CSI_InputBoxBoxAlignment=   0
            CSI_CalBackColor=   16777130
            CSI_CalDateFormat=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CSI_CurDayBackColor=   16777215
            CSI_CurDayForeColor=   51200
            CSI_ForceMondaySelectionOnly=   0   'False
            CSI_AllowBlankDate=   -1  'True
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin VB.PictureBox plcSelC5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   960
            Left            =   120
            ScaleHeight     =   960
            ScaleWidth      =   4380
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   840
            Width           =   4380
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Per Inquiry"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   1320
               TabIndex        =   16
               Top             =   480
               Value           =   1  'Checked
               Width           =   1260
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   720
               TabIndex        =   15
               Top             =   480
               Value           =   1  'Checked
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Remnant"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   3120
               TabIndex        =   14
               Top             =   240
               Value           =   1  'Checked
               Width           =   1140
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   1920
               TabIndex        =   13
               Top             =   240
               Value           =   1  'Checked
               Width           =   1200
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   720
               TabIndex        =   12
               Top             =   240
               Value           =   1  'Checked
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1680
               TabIndex        =   11
               Top             =   0
               Value           =   1  'Checked
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   10
               Top             =   0
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   2640
               TabIndex        =   17
               Top             =   480
               Width           =   720
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   3360
               TabIndex        =   18
               Top             =   480
               Width           =   960
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Trade"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   720
               TabIndex        =   19
               Top             =   720
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.PictureBox plcSelC2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   3075
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3075
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "Detail"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   600
               TabIndex        =   20
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "Summary"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1560
               TabIndex        =   21
               Top             =   0
               Width           =   1200
            End
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Active From"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   25
            Top             =   60
            Width           =   1005
         End
         Begin VB.Label lacSelCTo1 
            Appearance      =   0  'Flat
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2640
            TabIndex        =   43
            Top             =   450
            Width           =   255
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2640
            TabIndex        =   26
            Top             =   60
            Width           =   285
         End
         Begin VB.Label lacSelCTo 
            Appearance      =   0  'Flat
            Caption         =   "Entered From"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   42
            Top             =   450
            Width           =   1185
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
         Height          =   3780
         Left            =   5010
         ScaleHeight     =   3780
         ScaleWidth      =   4050
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4050
         Begin VB.CheckBox ckcAll 
            Caption         =   "All "
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   60
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3180
            Index           =   0
            ItemData        =   "Rptselpa.frx":0000
            Left            =   240
            List            =   "Rptselpa.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   3660
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   39
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   24
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
Attribute VB_Name = "RptSelPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselpa.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RPTSelPA.Frm  - CPP/CPM by Demo
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
'Vehicle link file- used to obtain start date
'Delivery file- used to obtain start date
'Vehicle conflict file- used to obtain start date
'Spot projection- used to obtain date status
'Library calendar file- used to obtain post log date status
'User- used to obtain discrepancy contract that was currently being processed
'      this is used if the system gos down
'Log
Dim imCodes() As Integer
Dim smLogUserCode As String
'Import contract report
'Spot week Dump
Dim imVefCode As Integer
Dim smVehName As String
Dim lmNoRecCreated As Long
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
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
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
    Dim slRptName As String
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
    'illistindex = lbcRptType.ListIndex
    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If rbcSelCInclude(0).Value Then
            slRptName = "PRICEDET.RPT"
        Else
            slRptName = "PRICESUM.RPT"
        End If
        If Not gOpenPrtJob(slRptName) Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gCmcGenPA(imGenShiftKey, smLogUserCode)
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
        Screen.MousePointer = vbHourglass
        gSalesPricingAnalysisGen
        Screen.MousePointer = vbDefault

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
    gCrCbfClear
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

Private Sub CSI_CalFrom_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalFrom_Change()
    mSetCommands
End Sub

Private Sub CSI_CalFrom_GotFocus()
    gCtrlGotFocus CSI_CalFrom
End Sub

Private Sub CSI_CalFrom1_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalFrom1_Change()
    mSetCommands
End Sub

Private Sub CSI_CalFrom1_GotFocus()
    gCtrlGotFocus CSI_CalFrom1
End Sub

Private Sub CSI_CalTo_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalTo_Change()
    mSetCommands
End Sub

Private Sub CSI_CalTo_GotFocus()
    gCtrlGotFocus CSI_CalTo
End Sub

Private Sub CSI_CalTo1_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalTo1_Change()
    mSetCommands
End Sub

Private Sub CSI_CalTo1_GotFocus()
    gCtrlGotFocus CSI_CalTo1
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
'Private Sub edcSelCFrom_Change()
'mSetCommands
'End Sub
'Private Sub edcSelCFrom_GotFocus()
'    gCtrlGotFocus edcSelCFrom
'End Sub
'Private Sub edcSelCFrom1_Change()
'    mSetCommands
'End Sub
'Private Sub edcSelCFrom1_GotFocus()
'    gCtrlGotFocus edcSelCFrom1
'End Sub
'Private Sub edcSelCTo_Change()
'mSetCommands
'End Sub
'Private Sub edcSelCTo_GotFocus()
'    gCtrlGotFocus edcSelCTo
'End Sub
'Private Sub edcSelCTo1_Change()
'    mSetCommands
'End Sub
'Private Sub edcSelCTo1_GotFocus()
'    gCtrlGotFocus edcSelCTo1
'End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    RptSelPA.Refresh
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
    'RptSelPA.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgClfPA
    Erase tgCffPA
    Erase imCodes
    PECloseEngine
    
    Set RptSelPA = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcSelection_Click(Index As Integer)
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  'False
        imSetAll = True
    End If


    mSetCommands
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    gCtrlGotFocus lbcSelection(Index)
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

    RptSelPA.Caption = smSelectedRptName & " Report"
    'frcOption.Caption = smSelectedRptName & " Selection"
    slStr = Trim$(smSelectedRptName)
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False

'    pbcSelC.Move 90, 255, 4515, 3360
    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    If tgSpf.sRUseCorpCal = "Y" Then       'if Using corp cal,retain in memory
        ilRet = gObtainCorpCal()
    End If
'    lbcSelection(0).Move 15, ckcAll.Height + 30, 4380, 3000

    plcSelC2.Enabled = True
    imSetAll = False
    ckcAll.Value = vbUnchecked  'False
    imSetAll = True

'    lacSelCFrom.Move 120, 75, 1380
'    edcSelCFrom.Move 1500, 30, 1350
'    lacSelCTo.Move 120, 390, 1380
'    lacSelCTo1.Move 2400, 390
'    edcSelCTo.Move 1500, 345, 1350
'    edcSelCTo1.Move 2715, 345

    'plcSelC1.Move 120, 675, 240
    'plcSelC2.Move 120, 885, 240

    gCenterStdAlone RptSelPA
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
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    gPopExportTypes cbcFileType         '10-20-01

    ckcAll.Enabled = True
    mSPersonPop lbcSelection(0)
    lbcSelection(0).Visible = True         'see demo list box
    ckcAll.Caption = "All Salespeople"
    ckcAll.Visible = True
    frcOption.Enabled = True
'    pbcSelC.Height = pbcSelC.Height - 60

    pbcSelC.Visible = True
    pbcOption.Visible = True

    mSetCommands
    Screen.MousePointer = vbDefault
    'gCenterModalForm RptSel
'    lacSelCFrom.Left = 120
'    lacSelCFrom1.Move 2325, 75
'    edcSelCFrom.Move 1290, edcSelCFrom.Top, 945
'    edcSelCFrom1.Move 2700, edcSelCFrom.Top, 945
'    lacSelCTo.Left = 120
'    edcSelCTo.Move 1290, edcSelCTo.Top, 945
'    lacSelCTo1.Left = 2325
'
'    'Date Selections
'    edcSelCTo1.Move 2700, edcSelCTo.Top, 945
'    edcSelCTo.MaxLength = 10
'    edcSelCTo1.MaxLength = 10
'    edcSelCFrom.MaxLength = 10
'    edcSelCFrom1.MaxLength = 10
'    lacSelCFrom.Caption = "Active: From"
'    lacSelCFrom1.Caption = "To"
'    lacSelCFrom.Visible = True
'    lacSelCFrom1.Visible = True
'    lacSelCTo.Caption = "Entered: From"
'    lacSelCTo1.Caption = "To"
'    lacSelCTo.Visible = True
'    lacSelCTo1.Visible = True
'    edcSelCFrom.Visible = True
'    edcSelCFrom1.Visible = True
'    edcSelCTo.Visible = True
'    edcSelCTo1.Visible = True
'    'Contract Type selection
'    plcSelC5.Height = 1040
'    plcSelC5.Move 120, 740
'    'plcSelC5.Caption = "Include"
'    ckcSelC5(0).Caption = "Holds"
'    ckcSelC5(0).Move 720, -30, 825
'    ckcSelC5(0).Visible = True
'    ckcSelC5(0).Value = vbChecked
'    ckcSelC5(1).Caption = "Orders"
'    ckcSelC5(1).Move 1545, -30, 900
'    ckcSelC5(1).Visible = True
'    ckcSelC5(1).Value = vbChecked
'    plcSelC5.Visible = True                            'hold/order boxes
'    ckcSelC5(2).Move 720, 220, 1060
'    ckcSelC5(2).Caption = "Standard"
'    ckcSelC5(2).Value = vbChecked
'    ckcSelC5(2).Visible = True
'    If ckcSelC5(2).Value = vbChecked Then
''           ckcSelC6_click 0, True
'    Else
''           ckcSelC5(2).Value = True
'    End If
'    ckcSelC5(3).Move 1800, 220, 1200
'    ckcSelC5(3).Caption = "Reserved"
'    ckcSelC5(3).Value = vbChecked
'    ckcSelC5(3).Visible = True
'    If ckcSelC5(3).Value = vbChecked Then
'        'ckcSelC6_click 1, True
'    Else
'        ckcSelC5(3).Value = vbChecked
'    End If
'    ckcSelC5(4).Move 2950, 240, 1200
'    ckcSelC5(4).Caption = "Remnant"
'    ckcSelC5(4).Value = vbChecked
'    ckcSelC5(4).Visible = True
'    If ckcSelC5(4).Value = vbChecked Then
''         ckcSelC6_click 2, True
'    Else
''         ckcSelC6(2).Value = True
'    End If
'    ckcSelC5(5).Move 720, 470, 600
'    ckcSelC5(5).Caption = "DR"
'    ckcSelC5(5).Value = vbChecked
'    ckcSelC5(5).Visible = True
'    If ckcSelC5(5).Value = vbChecked Then
''           ckcSelC6_click 3, True
'    Else
''           ckcSelC6(3).Value = True
'    End If
'    ckcSelC5(6).Move 1260, 470, 1260
'    ckcSelC5(6).Caption = "Per Inquiry"
'    ckcSelC5(6).Value = vbChecked
'    ckcSelC5(6).Visible = True
'    If ckcSelC5(6).Value = vbChecked Then
''           ckcSelC6_click 4, True
'    Else
''           ckcSelC6(4).Value = True
'    End If
'    ckcSelC5(7).Move 2500, 470, 700
'    ckcSelC5(7).Caption = "PSA"
'    ckcSelC5(7).Visible = True
'    If ckcSelC5(7).Value = vbChecked Then
''           ckcSelC6_click 4, True
'    Else
''           ckcSelC6(4).Value = True
'    End If
'
'    ckcSelC5(8).Move 3200, 470, 920
'    ckcSelC5(8).Caption = "Promo"
'    ckcSelC5(8).Visible = True
'    If ckcSelC5(8).Value = vbChecked Then
''           ckcSelC6_click 4, True
'    Else
''           ckcSelC6(4).Value = True
'    End If
'    ckcSelC5(9).Move 720, 720, 920
'    ckcSelC5(9).Caption = "Trade"
'    ckcSelC5(9).Value = vbChecked
'    ckcSelC5(9).Visible = True
'    If ckcSelC5(9).Value = vbChecked Then
''           ckcSelC6_click 4, True
'    Else
''           ckcSelC6(4).Value = True
'    End If
'    plcSelC2.Visible = True
'    plcSelC2.Width = 3500
'    plcSelC2.Move 120, 1800
'    'plcSelC2.Caption = "Show"
'    rbcSelCInclude(0).Width = 825
'    rbcSelCInclude(0).Caption = "Detail"
'    rbcSelCInclude(1).Move 1425
'    rbcSelCInclude(1).Caption = "Summary"
'    rbcSelCInclude(0).Value = True


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
    'gInitStdAlone RptSelPA, slStr, ilTestSystem
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
    If (igRptCallType = CONTRACTSJOB) And (igRptType = 3) Then
        igStdAloneMode = True 'Switch from/to stand alone mode-No DDE
        ilRet = gParseItem(slCommand, 5, "\", smLogUserCode)
        ilRet = gParseItem(slCommand, 6, "\", slStr)
        imVefCode = Val(slStr)
        ilRet = gParseItem(slCommand, 7, "\", smVehName)
        ilRet = gParseItem(slCommand, 8, "\", slStr)
        lmNoRecCreated = Val(slStr)
    End If
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
    Dim ilEnable As Integer
    Dim ilLoop As Integer
    ilEnable = False
    If ckcAll.Value = vbUnchecked Then
        ilEnable = True
    End If
    For ilLoop = 0 To lbcSelection(0).ListCount - 1
        If lbcSelection(0).Selected(ilLoop) Then
            ilEnable = True
        End If
    Next ilLoop

    cmcGen.Enabled = ilEnable
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSPersonPop                     *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Sales office list box *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mSPersonPop(lbcSelection As Control)
'
'   mSPersonPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopSalespersonBox(RptSelCt, 0, True, True, lbcSelection, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(RptSelPA, 0, True, True, lbcSelection, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gPopSalespersonBox)", RptSelPA
        On Error GoTo 0
    End If
    Exit Sub
mSPersonPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Unload RptSelPA
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
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
Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcSelC5_Paint()
    plcSelC5.CurrentX = 0
    plcSelC5.CurrentY = 0
    plcSelC5.Print "Include"
End Sub
Private Sub plcSelC2_Paint()
    plcSelC2.CurrentX = 0
    plcSelC2.CurrentY = 0
    plcSelC2.Print "Show"
End Sub
