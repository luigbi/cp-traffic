VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelAp 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5535
   ClientLeft      =   255
   ClientTop       =   2040
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
      Left            =   6615
      TabIndex        =   18
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8145
      TabIndex        =   8
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
      TabIndex        =   53
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
      TabIndex        =   40
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
      Left            =   360
      Top             =   4920
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
         TabIndex        =   6
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   7
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
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   14
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
         TabIndex        =   11
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
         TabIndex        =   13
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Vehicle Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3810
      Left            =   45
      TabIndex        =   15
      Top             =   1665
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
         ScaleWidth      =   4530
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4530
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   1440
            TabIndex        =   58
            Top             =   30
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Text            =   "9/10/2019"
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
         Begin VB.TextBox edcSelC10 
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
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   52
            Top             =   2580
            Width           =   1185
         End
         Begin VB.PictureBox plcSelC9 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   3930
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   2280
            Width           =   3930
            Begin VB.CheckBox ckcSelC9 
               Caption         =   "Hard Cost"
               Height          =   240
               Index           =   1
               Left            =   1560
               TabIndex        =   50
               Top             =   0
               Width           =   1125
            End
            Begin VB.CheckBox ckcSelC9 
               Caption         =   "NTR"
               Height          =   240
               Index           =   0
               Left            =   720
               TabIndex        =   49
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.PictureBox plcSelC8 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   57
            Top             =   990
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Pessimistic"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2880
               TabIndex        =   32
               Top             =   0
               Width           =   1395
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Optimistic"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1440
               TabIndex        =   31
               Top             =   -30
               Width           =   1305
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Most Likely"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   0
               TabIndex        =   30
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1260
            End
         End
         Begin VB.PictureBox plcSelC7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   3300
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   720
            Width           =   3300
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1800
               TabIndex        =   29
               Top             =   0
               Width           =   1065
            End
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Corporate"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   480
               TabIndex        =   28
               Top             =   0
               Width           =   1140
            End
         End
         Begin VB.PictureBox plcSelC6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   120
            ScaleHeight     =   480
            ScaleWidth      =   4380
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1770
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "PI"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   1920
               TabIndex        =   47
               Top             =   240
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   720
               TabIndex        =   46
               Top             =   240
               Value           =   1  'Checked
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Remnants"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3120
               TabIndex        =   45
               Top             =   0
               Value           =   1  'Checked
               Width           =   1200
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1920
               TabIndex        =   44
               Top             =   0
               Value           =   1  'Checked
               Width           =   1185
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   43
               Top             =   0
               Value           =   1  'Checked
               Width           =   1185
            End
         End
         Begin VB.PictureBox plcSelC5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1500
            Width           =   4380
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Cash"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   39
               Top             =   0
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Trade"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1800
               TabIndex        =   41
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox plcSelC3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1230
            Width           =   4380
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1680
               TabIndex        =   35
               Top             =   0
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   34
               Top             =   0
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.TextBox edcSelCTo1 
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
            Left            =   2640
            MaxLength       =   1
            TabIndex        =   27
            Top             =   360
            Width           =   345
         End
         Begin VB.TextBox edcSelCTo 
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
            Left            =   720
            MaxLength       =   4
            TabIndex        =   24
            Top             =   360
            Width           =   690
         End
         Begin VB.Label lacContract 
            Caption         =   "Contract #"
            Height          =   255
            Left            =   0
            TabIndex        =   51
            Top             =   2630
            Width           =   975
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   22
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label lacSelCTo1 
            Appearance      =   0  'Flat
            Caption         =   "Quarter"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1800
            TabIndex        =   25
            Top             =   390
            Width           =   765
         End
         Begin VB.Label lacSelCTo 
            Appearance      =   0  'Flat
            Caption         =   "Year"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   390
            Width           =   540
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
         Height          =   3420
         Left            =   4845
         ScaleHeight     =   3420
         ScaleWidth      =   4215
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4215
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   0
            ItemData        =   "Rptselap.frx":0000
            Left            =   120
            List            =   "Rptselap.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   55
            Top             =   360
            Width           =   3900
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Advertisers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   0
            Width           =   3945
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   1
            ItemData        =   "Rptselap.frx":0004
            Left            =   120
            List            =   "Rptselap.frx":0006
            TabIndex        =   56
            Top             =   360
            Visible         =   0   'False
            Width           =   3900
         End
         Begin VB.Label laclbcName 
            Appearance      =   0  'Flat
            Caption         =   "Compare To"
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
            Index           =   1
            Left            =   2300
            TabIndex        =   36
            Top             =   3060
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label laclbcName 
            Appearance      =   0  'Flat
            Caption         =   "Budget Names"
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
            Index           =   0
            Left            =   2055
            TabIndex        =   38
            Top             =   3120
            Visible         =   0   'False
            Width           =   1005
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
      TabIndex        =   17
      Top             =   150
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
         Width           =   1335
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
Attribute VB_Name = "RptSelAp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselap.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelAp.Frm       Actual Projection Comparison
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
'Dim hmVlf As Integer            'Vehicle link file handle
'Dim tmVlf As VLF                'VLF record image
'Dim imVlfRecLen As Integer      'VLF record length
'Dim tmVlfSrchKey0 As VLFKEY0            'VLF record image
'Dim tmVlfSrchKey1 As VLFKEY1            'VLF record image
'Delivery file- used to obtain start date
'Dim hmDlf As Integer            'Delivery Vehicle link file handle
'Dim tmDlf As DLF                'DLF record image
'Dim tmDlfSrchKey As DLFKEY0     'DLF record image
'Dim imDlfRecLen As Integer      'DLF record length
'Dim smDlfStartDate As String
'Vehicle conflict file- used to obtain start date
'Dim hmVcf As Integer            'Vehicle conflict file handle
'Dim tmVcf As VCF                'VCF record image
'Dim tmVcfSrchKey As VCFKEY0     'VCF record image
'Dim imVcfRecLen As Integer      'VCF record length
'Spot projection- used to obtain date status
'Dim hmJsr As Integer            'Spot Projection file handle
'Dim tmJsr As JSR                'JSR record image
'Dim tmJsrSrchKey As JSRKEY0     'JSR record image
'Dim imJsrRecLen As Integer      'JSR record length
'Library calendar file- used to obtain post log date status
'Dim hmLcf As Integer            'Library calendar file handle
'Dim tmLcf As LCF                'LCF record image
'Dim tmLcfSrchKey As LCFKEY0            'LCF record image
'Dim imLcfRecLen As Integer        'LCF record length
'User- used to obtain discrepancy contract that was currently being processed
'      this is used if the system gos down
'Dim hmSpf As Integer            'Site file handle
'Dim tmSpf As SPF                'SPF record image
'Dim tmSpfSrchKey As INTKEY0            'SPF record image
'Dim imSpfRecLen As Integer        'SPF record length
'
'Dim hmChf As Integer            'Contract header file handle
'Dim imChfRecLen As Integer
'Dim tmChfAdvtExt() As CHFADVTEXT
'Log
Dim smLogUserCode As String
'Import contract report
'Dim smChfConvName As String
'Dim smChfConvDate As String
'Dim smChfConvTime As String
'
'Spot week Dump
'Dim imVefCode As Integer
'Dim smVehName As String
'Dim lmNoRecCreated As Long
Dim imTerminate As Integer
'Dim ilAASCodes()  As Integer
'Dim tmsRec As LPOPREC
'Rate Card
'Dim tmRifRec() As RIFREC
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
    'ilIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).HWnd, LB_SELITEMRANGE, ilValue, llRg)
    Else
        imAllClicked = False
    End If

    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub
'Private Sub ckcAllAAS_Click()
'    'Code added because Value removed as parameter
'    Dim Value As Integer
'    Value = False
'    If ckcAllAAS.Value = vbChecked Then
'        Value = True
'    End If
'    'End of coded added
'    Dim ilIndex As Integer
'    Dim ilValue As Integer
'    ilIndex = lbcRptType.ListIndex
'    ilValue = Value
'    mSetCommands
'End Sub
Private Sub ckcSelC3_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSelC3(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
    End Select
    mSetCommands
End Sub
Private Sub ckcSelC5_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSelC5(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
'    ilListIndex = lbcRptType.ListIndex
'    Select Case igRptCallType
'        Case CONTRACTSJOB
'            If (igRptType = 0) And (ilListIndex > 1) Then
'                ilListIndex = ilListIndex + 1
'            End If
'            If ilListIndex = CNT_BOB_BYCNT Or ilListIndex = CNT_BOB_BYSPOT Or ilListIndex = CNT_BOB_BYSPOT_REPRINT Then
'                lbcSelection_click 5
'            End If
'    End Select
 '
    mSetCommands
End Sub
Private Sub ckcSelC6_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSelC6(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
'    ilListIndex = lbcRptType.ListIndex
'    Select Case igRptCallType
'        Case CONTRACTSJOB
'            If (igRptType = 0) And (ilListIndex > 1) Then
'                ilListIndex = ilListIndex + 1
'            End If
 '           If ilListIndex = CNT_BR Then
 '               If index = 1 And Value Then             'incl research clicked (if on, force rates to be included too)
 '                   ckcSelC6(0).Value = True
 '                   ckcSelC6(0).Enabled = False
 '               Else
 '                   ckcSelC6(0).Enabled = True
 '               End If
 '           End If
'    End Select
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
    Dim ilListIndex As Integer
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim llEarliestDate As Long
    Dim llLatestDate As Long
    Dim slStr As String
    Dim llCompareDate As Long
    Dim llCompareDate2 As Long
    Dim ilTemp As Integer
    Dim ilMonth As Integer
    Dim ilQuarter As Integer
    If igGenRpt Then
        Exit Sub
    End If
    igGenRpt = True
    igOutput = frcOutput.Enabled
    igCopies = frcCopies.Enabled
    igFile = frcFile.Enabled
    igOption = frcOption.Enabled
    frcOutput.Enabled = False
    frcCopies.Enabled = False
    frcFile.Enabled = False
    frcOption.Enabled = False

    ilListIndex = lbcRptType.ListIndex
    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenReportAp() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If
        ilRet = gCmcGenAp(ilListIndex, imGenShiftKey, smLogUserCode)

        If ilRet = -1 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
        ElseIf ilRet = 0 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If

        'Determine contracts to process based on their entered and modified dates
        Screen.MousePointer = vbHourglass
        slStr = RptSelAp!edcSelCTo.Text                'year
        ilTemp = Val(slStr)
        ilQuarter = RptSelAp!edcSelCTo1.Text              'month in text form (jan..dec)
        'If Mid$(slStr, 3, 1) = "/" Then
        '    ilMonth = Val(Left$(slStr, 2))
        'Else
        '    ilMonth = Val(Left$(slStr, 1))
        'End If
        'gGetMonthNoFromString slMonth, ilSaveMonth    'getmonth #
        'If ilSaveMonth = 0 Then                       'input isn't text month name, try month #
        '    ilSaveMonth = Val(slStr)
        'End If

        'If ilMonth = 0 Then                            'input isn't text month name, try month #
        '    ilMonth = Val(slStr)
        'End If
        ilMonth = (ilQuarter - 1) * 3 + 1                 'obtain starting month from the starting quarter
        ilRet = mObtainStartEndDates(ilTemp, ilMonth, 3, llEarliestDate, llLatestDate)

        'obtain previous years dates
        ilTemp = ilTemp - 1                     'previous year
        'month (ilSaveMonth) is the same
        ilRet = mObtainStartEndDates(ilTemp, ilMonth, 3, llCompareDate, llCompareDate2)

        gGetActProj llEarliestDate, llLatestDate, llCompareDate, llCompareDate2
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
            'ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
            '10-20-01
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)

        End If
        Screen.MousePointer = vbHourglass
        gCRGrfClear                  'Clear grf.btr file for reuse
        Screen.MousePointer = vbDefault
    Next ilJobs
    imGenShiftKey = 0
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
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

Private Sub CSI_CalFrom_GotFocus()
    gCtrlGotFocus CSI_CalFrom
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

'
'
'Private Sub edcSelCFrom_Change()
'    Dim ilListIndex As Integer
'    ilListIndex = lbcRptType.ListIndex
'    'Select Case igRptCallType
'    '    Case CONTRACTSJOB
'     '       If (igRptType = 0) And (ilListIndex > 1) Then
'    ' '           ilListIndex = ilListIndex + 1
'    '        End If
'    '        If ilListIndex = CNT_MAKEPLAN Then
'    '            ilLen = Len(edcSelCFrom)
'    '            If ilLen = 4 Then
'    '                slDate = "1/15/" & Trim$(edcSelCFrom)
'    '                slDate = gObtainStartStd(slDate)
'    ''                llDate = gDateValue(slDate)
'     '               mBudgetPop
'     '               'populate Rate Cards and bring in Rcf, Rif, and Rdf
'    '                ilRet = gPopRateCardBox(RptSelAp, llDate, RptSelAp!lbcSelection(12), tgRateCardCode(), smRateCardTag, -1)
'    '            End If
'    '        End If
'    ''End Select
'    mSetCommands
'End Sub
'Private Sub edcSelCFrom_GotFocus()
'    gCtrlGotFocus edcSelCFrom
'End Sub

Private Sub edcSelCTo_Change()
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    'Select Case igRptCallType
    '    Case CONTRACTSJOB
     '       If (igRptType = 0) And (ilListIndex > 1) Then
     ''           ilListIndex = ilListIndex + 1
     '       End If
     '       If ilListIndex = CNT_SALESANALYSIS Then
      '          ilLen = Len(edcSelCTo)
      '          If ilLen = 4 Then
      '              lbcSelection(4).Clear
      '              slDate = "1/15/" & Trim$(edcSelCTo)
      '              slDate = gObtainStartStd(slDate)
       '             llDate = gDateValue(slDate)
      '              mBudgetPop
     '               lbcSelection(4).Move 15, ckcall.Top + ckcall.Height + 30, 4380, 3000
     '               lbcSelection(4).Visible = True
     '               laclbcName(0).Visible = True
     '               laclbcName(0).Caption = "Budget Names"
      '              laclbcName(0).Move ckcall.Left, ckcall.Top + 30, 1725
      '              laclbcName(1).Visible = False
      '          End If
      '      End If
    'End Select
    mSetCommands
End Sub
Private Sub edcSelCTo_GotFocus()
    gCtrlGotFocus edcSelCTo
End Sub
Private Sub edcSelCTo_KeyPress(KeyAscii As Integer)
    Exit Sub
End Sub
Private Sub edcSelCTo1_Change()
    mSetCommands
End Sub
Private Sub edcSelCTo1_GotFocus()
    gCtrlGotFocus edcSelCTo1
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
    RptSelAp.Refresh
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
    'RptSelAp.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgRptSelBudgetCodeAP
    Erase tgMultiCntrCodeAP
    Erase lgPrintedCnts
    Erase tgClfAP
    Erase tgCffAP
    'Erase imCodes
    PECloseEngine
    
    Set RptSelAp = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcRptType_Click()

    pbcOption.Enabled = True
    pbcOption.Visible = True
    pbcSelC.Visible = True

    plcSelC3.Visible = True
    plcSelC5.Visible = True
    plcSelC6.Visible = True
'    plcSelC7.Visible = True
    plcSelC8.Visible = True
'    plcSelC1.Visible = False
'    plcSelC2.Visible = False
'    plcSelC4.Visible = False
'    mAskEffDate
    'plcSelC7.Caption = ""
    'plcSelC7.Move 0, 1000
'    plcSelC7.Move 120, edcSelCTo.Top + edcSelCTo.Height
    'rbcSelC7(0).Left = 60
'    rbcSelC7(0).Left = 0
'    rbcSelC7(0).Width = 1300
'    rbcSelC7(0).Caption = "Corporate"
'    rbcSelC7(0).Visible = True
'    rbcSelC7(1).Left = 1600
'    rbcSelC7(1).Width = 1200
'    rbcSelC7(1).Caption = "Standard"
'    rbcSelC7(1).Visible = True
'    rbcSelC7(2).Visible = False

    If tgSpf.sRUseCorpCal <> "N" Then       'if Using corp cal, dfault it; otherwise disable it
        rbcSelC7(1).Enabled = True
        rbcSelC7(0).Value = True
    Else
        rbcSelC7(0).Enabled = False
        rbcSelC7(1).Value = True
    End If

    'plcSelC8.Caption = ""
    'plcSelC8.Move 0, 1400
'    plcSelC8.Move 120, plcSelC7.Top + plcSelC7.Height
'    'ckcSelC8(0).Left = 60
'    ckcSelC8(0).Left = 0
'    ckcSelC8(0).Width = 1280
'    ckcSelC8(0).Value = vbChecked
'    ckcSelC8(0).Caption = "Most Likely"
'    ckcSelC8(0).Visible = True
'    ckcSelC8(1).Left = 1400
'    ckcSelC8(1).Width = 1180
'    ckcSelC8(1).Caption = "Optimistic"
'    ckcSelC8(1).Value = vbUnchecked
'    ckcSelC8(1).Visible = True
'    ckcSelC8(2).Left = 2600
'    ckcSelC8(2).Width = 1280
'    ckcSelC8(2).Value = vbUnchecked
'    ckcSelC8(2).Caption = "Pessimistic"
'    ckcSelC8(2).Visible = True
'
'    'plcSelC3.Caption = "Include"
'    'plcSelC3.Move 100, 1800
'    plcSelC3.Move 120, plcSelC8.Top + plcSelC8.Height
'    ckcSelC3(0).Left = 700
'    ckcSelC3(0).Width = 780
'    ckcSelC3(0).Value = vbChecked
'    ckcSelC3(0).Caption = "Holds"
'    ckcSelC3(0).Visible = True
'    ckcSelC3(1).Left = 1600
'    ckcSelC3(1).Width = 980
'    ckcSelC3(1).Value = vbChecked
'    ckcSelC3(1).Caption = "Orders"
'    ckcSelC3(1).Visible = True
    'plcSelC5.Caption = ""
    'plcSelC5.Move 100, 2030
'    plcSelC5.Move 120, plcSelC3.Top + plcSelC3.Height
'    ckcSelC5(0).Left = 700
'    ckcSelC5(0).Width = 780
'    ckcSelC5(0).Value = vbChecked
'    ckcSelC5(0).Caption = "Cash"
'    ckcSelC5(0).Visible = True
'    ckcSelC5(1).Left = 1600
'    ckcSelC5(1).Width = 1100
'    ckcSelC5(1).Value = vbUnchecked
'    ckcSelC5(1).Caption = "Trade"
'    ckcSelC5(1).Visible = True

    'plcSelC6.Caption = ""
    'plcSelC6.Move 100, 2260
'    plcSelC6.Move 120, plcSelC5.Top + plcSelC5.Height
'    plcSelC6.Height = 480
'    ckcSelC6(0).Left = 700
'    ckcSelC6(0).Width = 1200
'    ckcSelC6(0).Value = vbChecked
'    ckcSelC6(0).Caption = "Standard"
'    ckcSelC6(0).Visible = True
'    ckcSelC6(1).Left = 1800
'    ckcSelC6(1).Width = 1200
'    ckcSelC6(1).Value = vbChecked
'    ckcSelC6(1).Caption = "Reserved"
'    ckcSelC6(1).Visible = True
'    ckcSelC6(2).Left = 3000
'    ckcSelC6(2).Width = 1200
'    ckcSelC6(2).Value = vbChecked
'    ckcSelC6(2).Caption = "Remnants"
'    ckcSelC6(2).Visible = True
'    ckcSelC6(3).Move 700, 180
'    ckcSelC6(3).Width = 800
'    ckcSelC6(3).Value = vbChecked
'    ckcSelC6(3).Caption = "DR"
'    ckcSelC6(3).Visible = True
'    ckcSelC6(4).Move 1800, 180
'    ckcSelC6(4).Width = 800
'    ckcSelC6(4).Value = vbChecked
'    ckcSelC6(4).Caption = "PI"
'    ckcSelC6(4).Visible = True
mSetCommands
End Sub
Private Sub lbcSelection_Click(Index As Integer)
    Dim ilListIndex As Integer
    If Not lbcSelection(0).Index = 0 Then
        mSetCommands
    Else
        If Not imAllClicked Then
            ilListIndex = lbcRptType.ListIndex
            ckcAll.Enabled = True
            ckcAll.Visible = True
            ckcAll.Value = vbUnchecked
            lbcSelection(0).Visible = True
        Else
            imSetAll = False
            ckcAll.Value = vbUnchecked
            imSetAll = True
        End If
    End If
mSetCommands
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    gCtrlGotFocus lbcSelection(Index)
End Sub

'
'
'                   mAskEffDate - Ask Effective Date, Start Year
'                                 and Quarter
'
'                   6/7/97
'
'
'Private Sub mAskEffDate()
'    lacSelCFrom.Left = 120
'    edcSelCFrom.Move 1350, edcSelCFrom.Top, 945
'    edcSelCFrom.MaxLength = 10  '8  5/27/99 chged for short form m/d/yyyy
'    lacSelCFrom.Caption = "Effective Date"
'    lacSelCFrom.Top = 75
'    lacSelCFrom.Visible = True
'    edcSelCFrom.Visible = True
'    lacSelCTo.Caption = "Year"
'    lacSelCTo.Visible = True
'    lacSelCTo.Left = 120
'    lacSelCTo.Top = edcSelCFrom.Top + edcSelCFrom.Height + 75
'    lacSelCTo1.Left = 1580
'    lacSelCTo1.Caption = "Quarter"
'    lacSelCTo1.Width = 810
'    lacSelCTo1.Top = edcSelCFrom.Top + edcSelCFrom.Height + 75
'    lacSelCTo1.Visible = True
'    edcSelCTo.Move 600, edcSelCFrom.Top + edcSelCFrom.Height + 30, 600
'    edcSelCTo1.Move 2340, edcSelCFrom.Top + edcSelCFrom.Height + 30, 300
'    edcSelCTo.MaxLength = 4
'    edcSelCTo1.MaxLength = 1
'    edcSelCTo.Visible = True
'    edcSelCTo1.Visible = True
'End Sub

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
    imFirstActivate = True
    gCenterStdAlone RptSelAp
    Screen.MousePointer = vbHourglass
    'Start Crystal report engine
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        Exit Sub
    End If
    gPopExportTypes cbcFileType '10-20-01
    imAllClicked = False
    imSetAll = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitReport                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*                                                     *
'*******************************************************
Private Sub mInitReport()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Screen.MousePointer = vbHourglass
    lbcRptType.AddItem smSelectedRptName    'Do not change
    frcOption.Enabled = True
    pbcOption.Enabled = True
    frcOption.Visible = True
    pbcOption.Visible = True

'    lacSelCFrom.Move 120, 75
'    lacSelCTo.Move 120, 390
''    lacSelCFrom1.Move 2400, 75
'    lacSelCTo1.Move 2400, 390
'    edcSelCFrom.Move 1500, 30
''    edcSelCFrom1.Move 3240, 30
'    edcSelCTo.Move 1500, 345
'    edcSelCTo1.Move 2715, 345
'    lbcSelection(0).Visible = True          'Draw advertiser list box
'    'lbcSelection(0).Move 10, 400
'    lbcSelection(0).Move 15, ckcAll.Height + 30, 4380, 3000
'    'lbcSelection(0).Height = 3200
'
'    'Used for calculating rollover date; salesperson not selected
''    lbcSelection(2).Visible = False
'    lbcSelection(1).Visible = False         '9-10-19 all extra list boxes removed, change index
''    ckcAllAAS.Visible = False
'    'ckcAll.Move 0, 50
'    ckcAll.Move lbcSelection(0).Left, 0
'    ckcAll.Caption = "All Advertisers"
    'Populate advertiser list box
    ilRet = gPopAdvtBox(RptSelAp, lbcSelection(0), tgAdvertiser(), sgAdvertiserTag)

    'Populate lbcSelection(2) with salespeople
'    ilRet = gPopSalespersonBox(RptSelAp, 0, True, True, lbcSelection(2), tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(RptSelAp, 0, True, True, lbcSelection(1), tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)

    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
        RptSelAp.Caption = smSelectedRptName & " Report"
        frcOption.Caption = smSelectedRptName & " Selection"
        slStr = Trim$(smSelectedRptName)
        ilLoop = InStr(slStr, "&")
        If ilLoop > 0 Then
            slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
        End If
        frcOption.Caption = slStr & " Selection"
    End If
        ' Dan M. 7-01-08 added ntr/hard cost option
'    If UCase(RptSelAp.Caption) = "ACTUAL/PROJECTION COMPARISON REPORT" Then
'        mAddNTRHardCost
'        lacSelC10.Move 120, 100
'        edcSelC10.Left = lacSelC10.Left + lacSelC10.Width
'        edcSelC10.Visible = True
'        lacContract.Move 120, plcSelC9.Top + plcSelC9.Height
'        edcSelC10.Move lacContract.Left + lacContract.Width, lacContract.Top - 30
'        plcSelc10.Move 0, plcSelC9.Top + plcSelC9.Height
'        plcSelc10.Visible = True
'    End If
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub

'
'
'                   mObtainStartEndDates - obtain Standard Start and
'                   end date of given quarter
'                   <input>  ilYear = year to process
'                            ilmonth = starting month #
'                            ilNoMonths = total months to calc end date
'                   <return> llStartDate - STd start date of period
'                            llEnd Date - Std end date of period
'                            ilRet = true = ok
Private Function mObtainStartEndDates(ilYear As Integer, ilMonth As Integer, ilNoMonths As Integer, llStdStart As Long, llStdEnd As Long) As Integer
Dim slTemp As String
Dim slDate As String
        mObtainStartEndDates = False
        slDate = Trim$(str$(ilMonth)) & "/15/" & Trim$(str$(ilYear))
        slDate = gObtainStartStd(slDate)
        Do While ilNoMonths <> 0
            slTemp = gObtainEndStd(slDate)
            slDate = gObtainStartStd(slTemp)
            llStdEnd = gDateValue(slTemp)
            llStdEnd = llStdEnd + 1
            slDate = Format$(llStdEnd, "m/d/yy")
            ilNoMonths = ilNoMonths - 1
        Loop
        llStdStart = gDateValue(slDate)
        llStdEnd = llStdEnd - 1
        mObtainStartEndDates = True
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
    'gInitStdAlone RptSelAp, slStr, ilTestSystem
    igRptType = -1

    'If igStdAloneMode Then
    '    smSelectedRptName = "Actual/Projection Detail"
    '    igRptCallType = -1 'SETDEFSJOB 'PROPOSALPROJECTION 'NYFEED  'COLLECTIONSJOB 'SLSPCOMMSJOB   'LOGSJOB 'COPYJOB 'COLLECTIONSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB 'PROPOSALPROJECTION
    '    slCommand = ""
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
    Dim ilEnable As Integer
    Dim ilLoop As Integer
    Dim ilListIndex As Integer
    Dim Advt As Integer

    ilListIndex = lbcRptType.ListIndex

    If Not (ckcAll.Value = vbChecked) Then     'Check advertiser list box for selection
        For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
            If lbcSelection(0).Selected(ilLoop) Then
                Advt = True
                Exit For
            End If
        Next ilLoop
    Else
        Advt = True
    End If
    If Advt = True Then
'        If edcSelCFrom.Text <> "" And edcSelCTo.Text <> "" And edcSelCTo1.Text <> "" Then
        If CSI_CalFrom.Text <> "" And edcSelCTo.Text <> "" And edcSelCTo1.Text <> "" Then
            ilEnable = True       'Besides selecting advertiser, must also enter Eff.Date and # Qtrs.
        Else
            ilEnable = False
        End If
    End If

    If ilEnable Then
        If rbcOutput(0).Value Then  'Display
            ilEnable = True
        ElseIf rbcOutput(1).Value Then  'Print
            If edcCopies.Text <> "" Then
                ilEnable = True
            Else
                ilEnable = False
           End If
        Else    'Save As
           If (imFTSelectedIndex >= 0) And (edcFileName.Text <> "") Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        End If
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
    Unload RptSelAp
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcSelC9_Click()
    plcSelC9.CurrentX = 0
    plcSelC9.CurrentY = 0
    plcSelC9.Print "Include"
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
'Private Sub rbcSelCSelect_click(Index As Integer)
'    'Code added because Value removed as parameter
'    Dim Value As Integer
'    Value = rbcSelCSelect(Index).Value
'    'End of coded added
'End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcSelC8_Paint()
    plcSelC8.CurrentX = 0
    plcSelC8.CurrentY = 0
    plcSelC8.Print ""
End Sub
Private Sub plcSelC7_Paint()
    plcSelC7.CurrentX = 0
    plcSelC7.CurrentY = 0
    plcSelC7.Print "By"
End Sub
Private Sub plcSelC6_Paint()
    plcSelC6.CurrentX = 0
    plcSelC6.CurrentY = 0
    plcSelC6.Print ""
End Sub
Private Sub plcSelC5_Paint()
    plcSelC5.CurrentX = 0
    plcSelC5.CurrentY = 0
    plcSelC5.Print "Include"
End Sub
Private Sub plcSelC3_Paint()
    plcSelC3.CurrentX = 0
    plcSelC3.CurrentY = 0
    plcSelC3.Print "Include"
End Sub
'Private Sub plcSelC2_Paint()
'    plcSelC2.CurrentX = 0
'    plcSelC2.CurrentY = 0
'    plcSelC2.Print "Include"
'End Sub
'Private Sub plcSelC1_Paint()
'    plcSelC1.CurrentX = 0
'    plcSelC1.CurrentY = 0
'    plcSelC1.Print "Use"
'End Sub
'Private Sub plcSelC4_Paint()
'    plcSelC4.CurrentX = 0
'    plcSelC4.CurrentY = 0
'    plcSelC4.Print "Option"
'End Sub
    Private Sub mAddNTRHardCost()
'    lacSelC9.Caption = "Include "
'    lacSelC9.Move 0, 0, 670
'    ckcSelC9(0).Caption = "NTR"
'    ckcSelC9(0).Move lacSelC9.Width + lacSelC9.Left, 0, 750
'    Load ckcSelC9(1)
'    ckcSelC9(1).Caption = "Hard Cost"
'    ckcSelC9(1).Move ckcSelC9(0).Left + ckcSelC9(0).Width, 0, 2000
    plcSelC9.Move 120, plcSelC6.Top + plcSelC6.Height + 30, 5000
'    lacSelC9.Visible = True
'    ckcSelC9(0).Visible = True
'    ckcSelC9(1).Visible = True
'    plcSelC9.Visible = True
    End Sub
