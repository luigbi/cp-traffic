VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelIv 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5535
   ClientLeft      =   75
   ClientTop       =   1590
   ClientWidth     =   9915
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
   ScaleWidth      =   9915
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to  Report List"
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
      TabIndex        =   35
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
      TabIndex        =   33
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
      Top             =   4800
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
      Height          =   3690
      Left            =   45
      TabIndex        =   15
      Top             =   1785
      Width           =   9690
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
         ScaleWidth      =   4770
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   255
         Visible         =   0   'False
         Width           =   4770
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   1080
            TabIndex        =   23
            Top             =   0
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   450
            Text            =   "9/6/2019"
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
         Begin VB.ComboBox cbcSet2 
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
            Height          =   315
            Left            =   1440
            TabIndex        =   51
            Top             =   2280
            Width           =   1215
         End
         Begin VB.ComboBox cbcSet1 
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
            Height          =   315
            Left            =   1440
            TabIndex        =   49
            Top             =   1860
            Width           =   1215
         End
         Begin VB.TextBox edcSet2 
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
            Left            =   105
            TabIndex        =   50
            TabStop         =   0   'False
            Text            =   "Minor Set #"
            Top             =   2340
            Width           =   1215
         End
         Begin VB.TextBox edcSet1 
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
            Left            =   120
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   "Major Set #"
            Top             =   1890
            Width           =   1215
         End
         Begin VB.PictureBox plcSelC8 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   44
            Top             =   1560
            Width           =   4380
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3000
               TabIndex        =   47
               Top             =   0
               Value           =   1  'Checked
               Width           =   1020
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1680
               TabIndex        =   46
               Top             =   0
               Value           =   1  'Checked
               Width           =   1275
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Trade"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   45
               Top             =   0
               Value           =   1  'Checked
               Width           =   900
            End
         End
         Begin VB.PictureBox plcSelC7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4380
            TabIndex        =   26
            Top             =   360
            Width           =   4380
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Inventory"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1680
               TabIndex        =   28
               Top             =   0
               Width           =   1425
            End
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Avails"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   720
               TabIndex        =   27
               Top             =   0
               Value           =   -1  'True
               Width           =   1020
            End
         End
         Begin VB.PictureBox plcSelC6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4620
            TabIndex        =   39
            Top             =   1260
            Width           =   4620
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   3480
               TabIndex        =   43
               Top             =   0
               Width           =   960
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2760
               TabIndex        =   42
               Top             =   0
               Width           =   720
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Per Inquiry"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1440
               TabIndex        =   41
               Top             =   0
               Value           =   1  'Checked
               Width           =   1305
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   40
               Top             =   0
               Value           =   1  'Checked
               Width           =   600
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
            TabIndex        =   34
            Top             =   960
            Width           =   4380
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Std"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   36
               Top             =   0
               Value           =   1  'Checked
               Width           =   720
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1560
               TabIndex        =   37
               Top             =   0
               Value           =   1  'Checked
               Width           =   1200
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Remnant"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2880
               TabIndex        =   38
               Top             =   0
               Value           =   1  'Checked
               Width           =   1200
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
            TabIndex        =   29
            Top             =   660
            Width           =   4380
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1680
               TabIndex        =   32
               Top             =   0
               Value           =   1  'Checked
               Width           =   945
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   30
               Top             =   0
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.TextBox edcSelCFrom1 
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
            Left            =   3480
            MaxLength       =   3
            TabIndex        =   25
            Top             =   30
            Width           =   420
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   135
            TabIndex        =   22
            Top             =   60
            Width           =   885
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "# Quarters"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2400
            TabIndex        =   24
            Top             =   60
            Width           =   930
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
         Left            =   5040
         ScaleHeight     =   3420
         ScaleWidth      =   4440
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4445
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   1
            Left            =   240
            TabIndex        =   54
            Top             =   240
            Width           =   4125
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   0
            Left            =   240
            MultiSelect     =   2  'Extended
            TabIndex        =   53
            Top             =   240
            Visible         =   0   'False
            Width           =   4125
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   0
            Width           =   3945
         End
         Begin VB.Label laclbcName 
            Appearance      =   0  'Flat
            Caption         =   "Rate Card"
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
            Left            =   1800
            TabIndex        =   31
            Top             =   3240
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
      Left            =   60
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
         Width           =   1350
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8850
      Top             =   1080
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RptSelIv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptseliv.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy"#
'
' File Name: RptSelIv.Frm   Inventory Valuation
'       5/28/99 Allow 10 character date input, instead of 8 (m/d/yyyy)
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
Dim smRateCardTag As String
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
'Dim ilIndex As Integer
'Dim ilValue As Integer
'    ilIndex = lbcRptType.ListIndex
'    ilValue = Value
'''    If ilValue Then
' '       If igRptCallType = CONTRACTSJOB Then
' '           If (igRptType = 0) And (ilIndex > 1) Then
' '               ilIndex = ilIndex + 1
' '           End If
' '           If ilIndex = CNT_BR Then
''                ckcall.Visible = False
''                imSetAll = False
''                ckcall.Value = False
''                imSetAll = True
''                If rbcSelCSelect(0).Value Then          'advt option
''                    lbcSelection(0).Visible = False     'cnt list box
''                    lbcSelection(5).Visible = False     'advt list box
''                ElseIf rbcSelCSelect(1).Value Then      'agy option
''                    lbcSelection(10).Visible = False    'cnt list box
''                    lbcSelection(8).Visible = False     'agy list box
''                Else                                    'slsp option
''                    lbcSelection(10).Visible = False    'cnt list box
''                    lbcSelection(9).Visible = False     'slsp list box
''                End If
''            ElseIf ilIndex = CNT_VEHCPPCPM Then
''                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
''                llRet = SendMessagebyNum(lbcSelection(2).Hwnd, LB_SELITEMRANGE, ilValue, llRg)
''            End If
''        End If
''    Else                                                'turned All AAS off
''        If igRptCallType = CONTRACTSJOB Then
''            If (igRptType = 0) And (ilIndex > 1) Then
''                ilIndex = ilIndex + 1
''            End If
''            If ilIndex = CNT_BR Then
''                'show the AAS boxes again
' '               mSetupPopAAS                               'setup list box of valid contracts
' '           End If
''        End If
''    End If
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
    'mTerminate True
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
    ilNoJobs = 2
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenReportIV() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If
        ilRet = gCmcGenIv(ilListIndex, imGenShiftKey, smLogUserCode)

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
        Screen.MousePointer = vbHourglass
        'call the prepass code in RptCrIv, but only for the first report
        If igJobRptNo = 1 Then
            gCRQAvailsGenIV
        End If
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
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
        End If
        If igJobRptNo = 2 Then
            'clear avr.btr records to zero the file for the next report
            gCRQAvailsClearIV
        End If

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
'    '                ilRet = gPopRateCardBox(RptSelIv, llDate, RptSelIv!lbcSelection(12), tgRateCardCode(), smRateCardTag, -1)
'    '            End If
'    '        End If
'    ''End Select
'    mSetCommands
'End Sub
Private Sub CSI_CalFrom_GotFocus()
    gCtrlGotFocus CSI_CalFrom
End Sub

Private Sub edcSelCFrom1_Change()
    mSetCommands
End Sub
Private Sub edcSelCFrom1_GotFocus()
    gCtrlGotFocus edcSelCFrom1
End Sub
Private Sub edcSelCFrom1_KeyPress(KeyAscii As Integer)
    Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            If ilListIndex = 17 Then    'Quarterly Avail
                'Filter characters (allow only BackSpace, numbers 0 thru 9
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
    End Select
End Sub
'Private Sub edcSelCTo_Change()
'    Dim ilListIndex As Integer
'    ilListIndex = lbcRptType.ListIndex
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
'    mSetCommands
'End Sub
'Private Sub edcSelCTo_GotFocus()
'    gCtrlGotFocus edcSelCTo
'End Sub
'Private Sub edcSelCTo_KeyPress(KeyAscii As Integer)
'    Exit Sub
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
    RptSelIv.Refresh
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
    'RptSelIv.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgRateCardCode
    PECloseEngine
    
    Set RptSelIv = Nothing   'Remove data segment
    
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
    '9-6-19 Remove all extra controls on form and place the controls appropriately.
    
'    plcSelC3.Visible = True
'    plcSelC5.Visible = True
'    plcSelC6.Visible = True
'    plcSelC7.Visible = True
'    plcSelC8.Visible = True
''    plcSelC1.Visible = False
''    plcSelC2.Visible = False
''    plcSelC4.Visible = False
'    lacSelCFrom.Move 120, 75, 1080
''    lacSelCFrom.Caption = "Start Date"
'    lacSelCFrom.Visible = True
'    edcSelCFrom.Move 1005, 30
'    edcSelCFrom.Visible = True
'    edcSelCFrom.MaxLength = 10
'
'    lacSelCFrom1.Move 2505, 75, 1380
''    lacSelCFrom1.Caption = "# Quarters"
'    lacSelCFrom1.Visible = True
'    edcSelCFrom1.Move 3500, 30
'    edcSelCFrom1.MaxLength = 2
'    edcSelCFrom1.Visible = True
'
''    lacSelCTo1.Visible = False
''    lacSelCTo.Visible = False
''    edcSelCTo1.Visible = False
''    edcSelCTo.Visible = False
'    'plcSelC7.Caption = "Show"
'    plcSelC7.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 60
'    rbcSelC7(0).Left = 720
'    rbcSelC7(0).Width = 1000
'    rbcSelC7(0).Caption = "Avails"
'    rbcSelC7(0).Value = True
'    rbcSelC7(0).Visible = True
'    rbcSelC7(1).Left = 1720
'    rbcSelC7(1).Width = 1200
'    rbcSelC7(1).Caption = "Inventory"
'    rbcSelC7(1).Visible = True
''    rbcSelC7(2).Visible = False
'
'    'plcSelC3.Caption = "Include"
'    plcSelC3.Move 120, plcSelC7.Top + plcSelC7.Height + 30
'    ckcSelC3(0).Move 700, -30, 760
'    ckcSelC3(0).Value = vbChecked
'    ckcSelC3(0).Caption = "Holds"
'    ckcSelC3(0).Visible = True
'    ckcSelC3(1).Move 1600, -30, 900
'    ckcSelC3(1).Value = vbChecked
'    ckcSelC3(1).Caption = "Orders"
'    ckcSelC3(1).Visible = True
'    'plcSelC5.Caption = ""
'    plcSelC5.Move 120, plcSelC3.Top + plcSelC3.Height
'    ckcSelC5(0).Move 700, -30, 700
'    ckcSelC5(0).Value = vbChecked
'    ckcSelC5(0).Caption = "Std"
'    ckcSelC5(0).Visible = True
'    ckcSelC5(1).Move 1600, -30, 1100
'    ckcSelC5(1).Value = vbChecked
'    ckcSelC5(1).Caption = "Reserved"
'    ckcSelC5(1).Visible = True
'    ckcSelC5(2).Move 2900, -30, 1060
'    ckcSelC5(2).Value = vbChecked
'    ckcSelC5(2).Caption = "Remnant"
'    ckcSelC5(2).Visible = True
'
'    'plcSelC6.Caption = ""
'    plcSelC6.Move 120, plcSelC5.Top + plcSelC5.Height
'    ckcSelC6(0).Move 700, -30, 560
'    ckcSelC6(0).Value = vbChecked
'    ckcSelC6(0).Caption = "DR"
'    ckcSelC6(0).Visible = True
'    ckcSelC6(1).Move 1300, -30, 1200
'    ckcSelC6(1).Value = vbChecked
'    ckcSelC6(1).Caption = "Per Inquiry"
'    ckcSelC6(1).Visible = True
'    ckcSelC6(2).Move 2600, -30, 700
'    ckcSelC6(2).Caption = "PSA"
'    ckcSelC6(2).Visible = True
'    ckcSelC6(3).Move 3300, -30, 860
'    ckcSelC6(3).Caption = "Promo"
'    ckcSelC6(3).Visible = True
'    'plcSelC8.Caption = ""
'    plcSelC8.Move 120, plcSelC6.Top + plcSelC6.Height
'    ckcSelC8(0).Move 700, -30, 800
'    ckcSelC8(0).Value = vbChecked
'    ckcSelC8(0).Caption = "Trade"
'    ckcSelC8(0).Visible = True
'    ckcSelC8(1).Move 1700, -30, 960
'    ckcSelC8(1).Caption = "Missed"
'    ckcSelC8(1).Value = vbChecked
'    ckcSelC8(1).Visible = True
'    ckcSelC8(2).Move 2800, -30, 800
'    ckcSelC8(2).Value = vbChecked
'    ckcSelC8(2).Caption = "Fill"
'    ckcSelC8(2).Visible = True
    gPopVehicleGroups RptSelIv!cbcSet1, tgVehicleSets1(), False
    gPopVehicleGroups RptSelIv!cbcSet2, tgVehicleSets2(), True
'    edcSet1.Move 120, plcSelC8.Top + plcSelC8.Height + 90
'    cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 30
'    edcSet2.Move 120, cbcSet1.Top + cbcSet1.Height + 90
'    cbcSet2.Move cbcSet1.Left, edcSet2.Top - 30
mSetCommands
End Sub
Private Sub lbcSelection_Click(Index As Integer)
    Dim ilListIndex As Integer
'    If Not lbcSelection(1).Index = 0 Then
'        mSetCommands
'    Else
    If Index = 0 Then            'vehicle selection
        If Not imAllClicked Then
            ilListIndex = lbcRptType.ListIndex
            ckcAll.Enabled = True
            ckcAll.Visible = True
            ckcAll.Value = vbUnchecked  'False
            lbcSelection(0).Visible = True
        Else
            imSetAll = False
            ckcAll.Value = vbUnchecked  'False
            imSetAll = True
        End If
    End If
'    End If
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
    'dummy = LlSetOption(hdJob, LL_OPTION_HELPAVAILABLE, False)
    'ilDummy = LlSetOption(hdJob, LL_OPTION_SORTVARIABLES, True)
    'ilDummy = LlSetOption(hdJob, LL_OPTION_ONLYONETABLE, Not ilMultiTable)

'    frcOption.Caption = "Inventory Valuation Selection"

    imAllClicked = False
    imSetAll = True
    gCenterStdAlone RptSelIv
'
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
    Dim llDate As Long
'    Dim hmVef As Integer                    'Vehiclee file handle
'    Dim tmVef As VEF                        'VEF record image
'    Dim tmVefSrchKey As INTKEY0             'VEF record image
'    Dim imVefRecLen As Integer              'VEF record length
'    'ReDim tgVefCode(5000) As VEFCODE
'    Dim hmRpf As Integer                    'Rate card file handle
'    Dim tmRpf As RPF                        'RCF record image
'    Dim tmRpfSrchKey As INTKEY0             'RCF record image
'    Dim imRpfRecLen As Integer              'RCF record length
'    'ReDim tgRcfCode(5000) As RCFCODE
    'build list box for Save to file option
    gPopExportTypes cbcFileType     '10-20-01
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0


    Screen.MousePointer = vbHourglass
    lbcRptType.AddItem smSelectedRptName    'Do not change
    frcOption.Enabled = True
    pbcOption.Enabled = True
    frcOption.Visible = True
    pbcOption.Visible = True

    lbcSelection(0).Visible = True          'Draw vehicle list box
    lbcSelection(0).Move 240, 300
    lbcSelection(0).Height = 1440

    'ckcAll.Move 0, 0
    ckcAll.Left = lbcSelection(0).Left
    ckcAll.Caption = "All Vehicles"
    ckcAll.Visible = True
'    lbcSelection(12).Visible = True          'Draw rate card box
'    lbcSelection(12).Move 10, 2000
'    lbcSelection(12).Height = 1440
    lbcSelection(1).Move 240, 2000       'remove all extra list boxes, change index 12 to 1
    lbcSelection(1).Height = 1440
    laclbcName.Caption = "Rate Cards"
    laclbcName.Move 240, lbcSelection(0).Top + lbcSelection(0).Height + 90
    laclbcName.Visible = True
    'Populate vehicle list box
    ilRet = gPopUserVehicleBox(RptSelIv, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(0), tgCSVNameCode(), sgCSVNameCodeTag)    'lbcCSVNameCode)
    'Populate Rate Card list box
'    ilRet = gPopRateCardBox(RptSelIv, llDate, RptSelIv!lbcSelection(12), tgRateCardCode(), smRateCardTag, -1)
    ilRet = gPopRateCardBox(RptSelIv, llDate, RptSelIv!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)


'    hmVef = CBtrvTable(ONEHANDLE)           'Load Vehicle list box
'    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        btrDestroy hmVef
'        Exit Sub
'    End If
'    imVefRecLen = Len(tmVef)
'    btrExtClear hmVef
'    lbcSelection(0).Clear
'    ilRet = btrGetFirst(hmVef, tmVef, imVefRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)
'    For ilLoop = 0 To imVefRecLen Step 1
'        lbcSelection(0).AddItem (tmVef.sName)
'        ilRet = btrGetNext(hmVef, tmVef, imVefRecLen, 0, BTRV_LOCK_NONE)
'    Next ilLoop
'
'    ilRet = btrClose(hmVef)
'    btrDestroy hmVef
'
'
'
'    hmRpf = CBtrvTable(ONEHANDLE)           'Load Rate Card list box
'   ilRet = btrOpen(hmRpf, "", sgDBPath & "Rpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'   If ilRet <> BTRV_ERR_NONE Then
'        btrDestroy hmRpf
'        Exit Sub
'    End If
'    imRpfRecLen = Len(tmRpf)
'    btrExtClear hmRpf
'    lbcSelection(1).Clear
'    ilRet = btrGetFirst(hmRpf, tmRpf, imRpfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'    For ilLoop = 0 To imRpfRecLen Step 1
'    lbcSelection(1).AddItem (tmRpf.sName)
'        ilRet = btrGetNext(hmRpf, tmRpf, imRpfRecLen, 0, BTRV_LOCK_NONE)
'    Next ilLoop
'
'   ilRet = btrClose(hmRpf)
'    btrDestroy hmRpf
'
    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
        RptSelIv.Caption = smSelectedRptName & " Report"
        frcOption.Caption = smSelectedRptName & " Selection"
        slStr = Trim$(smSelectedRptName)
        ilLoop = InStr(slStr, "&")
        If ilLoop > 0 Then
            slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
        End If
        frcOption.Caption = slStr & " Selection"
    End If
    mSetCommands
    Screen.MousePointer = vbDefault
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
    'gInitStdAlone RptSelIv, slStr, ilTestSystem
    igRptType = -1

    'If igStdAloneMode Then
    '    'smSelectedRptName = "Copy Inventory by Advertiser"
    '    smSelectedRptName = "Inventory Valuation Detail"
    '    igRptCallType = -1 'SETDEFSJOB 'PROPOSALPROJECTION 'NYFEED  'COLLECTIONSJOB 'SLSPCOMMSJOB   'LOGSJOB 'COPYJOB 'COLLECTIONSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB 'PROPOSALPROJECTION
    '    'igRptType = 0   '3 'Log     '0   'Summary '3 Program  '1  links
    '    slCommand = "" '"x\x\x\x\2\2/6/95\7\12M\12M\1\26" '"" '"CONT0802.ASC\11/20/94\10:11:0 AM" '"x\x\x\x\2"
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
    Dim slStr As String
    Dim slDate As String
    Dim ilMonth As Integer              'date conversion for MakePlan
    Dim ilVehicles As Integer
    Dim ilRateCard As Integer
    Dim slQuart As String
    Dim ilDiff As Integer
    ilListIndex = lbcRptType.ListIndex

    'Check Vehicle list box for selection
    If Not (ckcAll.Value = vbChecked) Then
        For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
            If lbcSelection(0).Selected(ilLoop) Then
                ilVehicles = True
                Exit For
            End If
        Next ilLoop
    Else
        ilVehicles = True
    End If
    'Check Rate Card list box for selection
    '9-6-19 all extra list boxes removed, change r/c from12  to 1
'    For ilLoop = 0 To lbcSelection(12).ListCount - 1 Step 1
    For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
'        If lbcSelection(12).Selected(ilLoop) Then
        If lbcSelection(1).Selected(ilLoop) Then
            ilRateCard = True
            Exit For
        End If
    Next ilLoop
    'Verify that the #quarters requested does not exceed #quarters left in the year
'        slStr = RptSelIv!edcSelCFrom.Text
        slStr = RptSelIv!CSI_CalFrom.Text               '9-6-19 use csi calendar control vs edit box
        slDate = gObtainYearStartDate(0, slStr)
    If Mid$(slStr, 3, 1) = "/" Then
        ilMonth = Val(Left$(slDate, 2))
    Else
        ilMonth = Val(Left$(slDate, 1))
    End If
    slQuart = RptSelIv!edcSelCFrom1.Text
    If (ilMonth >= 1 And ilMonth <= 3) And (Val(slQuart) <= 4) Then
        ilDiff = True
    ElseIf (ilMonth >= 4 And ilMonth <= 6) And (Val(slQuart) <= 3) Then
        ilDiff = True
    ElseIf (ilMonth >= 7 And ilMonth <= 9) And (Val(slQuart) <= 2) Then
        ilDiff = True
    ElseIf (ilMonth >= 10 And ilMonth <= 11) And (Val(slQuart) = 1) Then
        ilDiff = True
    ElseIf ilMonth = 12 Then
        ilDiff = True
    Else
        ilDiff = False
    End If
    'Besides selecting Vehicle and Rate Card, must also enter Eff.Date and # Qtrs.
    If ilVehicles And ilRateCard And ilDiff = True Then
'        If edcSelCFrom.Text <> "" And edcSelCFrom1.Text <> "" Then
        If CSI_CalFrom.Text <> "" And edcSelCFrom1.Text <> "" Then          '9-6-17 use csi calendar control vs edit box
            ilEnable = True
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
    Unload RptSelIv
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
Private Sub plcSelC8_Paint()
    plcSelC8.CurrentX = 0
    plcSelC8.CurrentY = 0
    plcSelC8.Print ""
End Sub
Private Sub plcSelC7_Paint()
    plcSelC7.CurrentX = 0
    plcSelC7.CurrentY = 0
    plcSelC7.Print "Show"
End Sub
Private Sub plcSelC6_Paint()
    plcSelC6.CurrentX = 0
    plcSelC6.CurrentY = 0
    plcSelC6.Print ""
End Sub
Private Sub plcSelC5_Paint()
    plcSelC5.CurrentX = 0
    plcSelC5.CurrentY = 0
    plcSelC5.Print ""
End Sub
Private Sub plcSelC3_Paint()
    plcSelC3.CurrentX = 0
    plcSelC3.CurrentY = 0
    plcSelC3.Print "Include"
End Sub

