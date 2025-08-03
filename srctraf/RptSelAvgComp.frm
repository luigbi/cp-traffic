VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelAvgCmp 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Average Comparison"
   ClientHeight    =   5685
   ClientLeft      =   165
   ClientTop       =   1305
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
   ScaleHeight     =   5685
   ScaleWidth      =   9270
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   38
      Top             =   615
      Width           =   2055
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
      TabIndex        =   46
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
      TabIndex        =   47
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
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4320
      Top             =   5040
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
      TabIndex        =   35
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   43
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
         TabIndex        =   39
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
         TabIndex        =   42
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   41
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Rate / Spot Price Comparison Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4275
      Left            =   75
      TabIndex        =   44
      Top             =   1365
      Width           =   9090
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
         Height          =   4020
         Left            =   4785
         ScaleHeight     =   4020
         ScaleWidth      =   4230
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   180
         Width           =   4230
         Begin VB.CheckBox ckcAllVehicles 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   0
            Width           =   1905
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   0
            ItemData        =   "RptSelAvgComp.frx":0000
            Left            =   120
            List            =   "RptSelAvgComp.frx":0007
            MultiSelect     =   2  'Extended
            TabIndex        =   34
            Top             =   300
            Width           =   3945
         End
      End
      Begin VB.PictureBox pbcSelC 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         FillColor       =   &H8000000F&
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
         Height          =   4050
         Left            =   75
         ScaleHeight     =   4050
         ScaleWidth      =   4710
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   4710
         Begin VB.CheckBox ckcShowUnitPrice 
            Caption         =   "Display average rate using the rate card price"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   62
            Top             =   3105
            Visible         =   0   'False
            Width           =   4080
         End
         Begin VB.TextBox edcContract 
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
            Left            =   960
            MaxLength       =   9
            TabIndex        =   61
            Top             =   2730
            Width           =   1125
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4575
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   2400
            Width           =   4575
            Begin VB.OptionButton rbcSpotLen 
               Caption         =   "All"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   2190
               TabIndex        =   57
               Top             =   0
               Width           =   705
            End
            Begin VB.OptionButton rbcSpotLen 
               Caption         =   "30 and 60 only"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2925
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   1450
            End
            Begin VB.Label lacSpotLen 
               Appearance      =   0  'Flat
               Caption         =   "Spot Lengths to include"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   0
               TabIndex        =   58
               Top             =   0
               Width           =   2085
            End
         End
         Begin VB.OptionButton rbcAvgRatePrice 
            Caption         =   "Avg Spot Price"
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   1800
            TabIndex        =   51
            Top             =   3675
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.OptionButton rbcAvgRatePrice 
            Caption         =   "Avg Rate"
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   3675
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox edcStartYear 
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
            Left            =   1335
            MaxLength       =   4
            TabIndex        =   9
            Top             =   180
            Width           =   600
         End
         Begin VB.PictureBox plcBillOrCollect 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4575
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1815
            Width           =   4575
            Begin VB.OptionButton rbcUseLines 
               Caption         =   "Pkg"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   1080
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   0
               Width           =   1170
            End
            Begin VB.OptionButton rbcUseLines 
               Caption         =   "Air"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2280
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   885
            End
            Begin VB.Label lacUseLines 
               Appearance      =   0  'Flat
               Caption         =   "Use Lines"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   0
               TabIndex        =   53
               Top             =   0
               Width           =   885
            End
         End
         Begin VB.PictureBox plcTotalsBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4575
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2115
            Width           =   4575
            Begin VB.OptionButton rbcAvgBy 
               Caption         =   "Separate 30/60"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   1080
               TabIndex        =   60
               Top             =   0
               Value           =   -1  'True
               Width           =   1785
            End
            Begin VB.OptionButton rbcAvgBy 
               Caption         =   "Combined"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2925
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   0
               Width           =   1230
            End
            Begin VB.Label lacAvgBy 
               Appearance      =   0  'Flat
               Caption         =   "Avg by"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   0
               TabIndex        =   54
               Top             =   0
               Width           =   850
            End
         End
         Begin VB.PictureBox plcAllTypes 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   120
            ScaleHeight     =   1215
            ScaleWidth      =   4575
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   555
            Width           =   4575
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Non-Polit"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   16
               Left            =   1942
               TabIndex        =   31
               Top             =   930
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Polit"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   15
               Left            =   810
               TabIndex        =   30
               Top             =   930
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "N/C"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   14
               Left            =   1942
               TabIndex        =   29
               Top             =   690
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Rep"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   12
               Left            =   3720
               TabIndex        =   28
               Top             =   690
               Width           =   705
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "AirTime"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   11
               Left            =   2700
               TabIndex        =   27
               Top             =   690
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   3525
               TabIndex        =   25
               Top             =   450
               Width           =   960
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   2700
               TabIndex        =   24
               Top             =   450
               Width           =   765
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "PI"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   1942
               TabIndex        =   23
               Top             =   450
               Value           =   1  'Checked
               Width           =   675
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   795
               TabIndex        =   22
               Top             =   450
               Value           =   1  'Checked
               Width           =   690
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Remnant"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   3375
               TabIndex        =   21
               Top             =   210
               Value           =   1  'Checked
               Width           =   1080
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   1942
               TabIndex        =   20
               Top             =   210
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   795
               TabIndex        =   19
               Top             =   210
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1942
               TabIndex        =   18
               Top             =   -30
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   795
               TabIndex        =   17
               Top             =   -30
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Trade"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   795
               TabIndex        =   26
               Top             =   690
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.Label lacInclude 
               Appearance      =   0  'Flat
               Caption         =   "Include"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   0
               TabIndex        =   52
               Top             =   0
               Width           =   645
            End
         End
         Begin VB.TextBox edcYears 
            Alignment       =   2  'Center
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
            Left            =   3975
            MaxLength       =   3
            TabIndex        =   11
            Text            =   "5"
            Top             =   180
            Width           =   420
         End
         Begin VB.Label lacContract 
            Appearance      =   0  'Flat
            Caption         =   "Contr#"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   32
            Top             =   2760
            Width           =   645
         End
         Begin VB.Label lacStart 
            Appearance      =   0  'Flat
            Caption         =   "Current Year"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   90
            TabIndex        =   8
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label lacYears 
            Appearance      =   0  'Flat
            Caption         =   "# of History Years "
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2370
            TabIndex        =   10
            Top             =   210
            Width           =   1515
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   40
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   36
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
         Width           =   1005
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
Attribute VB_Name = "RptSelAvgCmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptSelAvgCmp.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  ckcAvails_Click                                                                       *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelAvgCmp.Frm - Participant Payables
'   Report to providea means to pay participant fees by combining Billing, Collectiosn, and Business on the Books
'   All by std months, for rolling 12 months.
'
' Release: 4.3
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal X%, ByVal Y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal X%, ByVal Y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAllVehicle As Integer 'True=Set list box; False= don't change list box
Dim imSetAllPart As Integer

Dim imAllVehicleClicked As Integer
Dim imAllPartClicked As Integer


Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name

Dim imCodes() As Integer
Dim smLogUserCode As String
'Import contract report
'Spot week Dump
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
Private Sub ckcAllVehicles_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllVehicles.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllVehicle Then
        imAllVehicleClicked = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllVehicleClicked = False
    End If
    mSetCommands
End Sub
Private Sub cmcBrowse_Click()
    gAdjustCDCFilter imFTSelectedIndex, cdcSetup
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
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim llPacingDate As Long
    Dim ilLastBilledInx As Integer
    
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

    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenAverageCompare() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If

        ilRet = gCmcAverageCompare()        '
        If ilRet = -1 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            'mTerminate
'            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
        ElseIf ilRet = 0 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If

        Screen.MousePointer = vbHourglass
        mCreateAverageCompare     'pre-pass
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
    Next ilJobs
    
    sgVehicleTag = ""
    imGenShiftKey = 0

    Screen.MousePointer = vbHourglass
    gCRGrfClear
    Screen.MousePointer = vbDefault
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    'pbcClickFocus.SetFocus
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

Private Sub edcYears_Change()
    mSetCommands
End Sub

Private Sub edcYears_GotFocus()
    gCtrlGotFocus edcYears
End Sub
Private Sub edcStartYear_Change()
    mSetCommands
End Sub
Private Sub edcStartYear_gotfocus()
    gCtrlGotFocus edcStartYear
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
    RptSelAvgCmp.Refresh
End Sub

Private Sub Form_Click()
    'pbcClickFocus.SetFocus
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
    'RptSelAvgCmp.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgMNFCodeRpt
    Erase imCodes
    PECloseEngine
    
    Set RptSelAvgCmp = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    If Index = 0 Then           'selective vehicles
        imSetAllVehicle = False
        ckcAllVehicles.Value = vbUnchecked
        imSetAllVehicle = True
'    ElseIf Index = 1 Then
'        imSetAllPart = False
'        'ckcAllParticipants.Value = vbUnchecked
'        imSetAllPart = True
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

    RptSelAvgCmp.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imSetAllVehicle = True
    imSetAllPart = True
    
    imAllVehicleClicked = False
    imAllPartClicked = False
    
    If InStr(1, frcOption.Caption, "Rate") > 0 Then
        rbcAvgRatePrice(0).Value = True: rbcAvgBy(0).Enabled = False: rbcAvgBy(1).Enabled = False: ckcShowUnitPrice.Visible = True
    Else
        rbcAvgRatePrice(1).Value = True
    End If
    gCenterStdAlone RptSelAvgCmp
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
Dim ilRet As Integer
    gPopExportTypes cbcFileType     '10-20-01
    pbcSelC.Visible = False

    Screen.MousePointer = vbHourglass

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    'lbcSelection(1).Clear
    'lbcSelection(1).Tag = ""
    ilRet = gObtainVef()
    ilRet = gRptVehPop(RptSelAvgCmp, lbcSelection(0), VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHNTR + VEHREP_W_CLUSTER + VEHREP_WO_CLUSTER + ACTIVEVEH + VEHSPORT)
    'ilRet = gPopMnfPlusFieldsBox(RptSelAvgCmp, lbcSelection(1), tgMnfCodeCT(), sgMNFCodeTagCT, "H1")

    'cbcSort1_Click
    frcOption.Enabled = True
    pbcSelC.Visible = True
    pbcOption.Visible = True
    lbcSelection(0).Visible = True
    'lbcSelection(1).Visible = True
    pbcOption.Visible = True
    pbcOption.Enabled = True
    
    mSetCommands
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSel
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
    'gInitStdAlone RptSelAvgCmp, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Quarterly Booked Spots"
    '    igRptCallType = -1  'unused in standalone exe, CONTRACTSJOB 'SLSPCOMMSJOB   'LOGSJOB 'CONTRACTSJOB 'COPYJOB 'COLLECTIONSJOB'CONTRACTSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB 'PROPOSALPROJECTION
    '    igRptType = -1  'unused in standalone exe   '3 'Log '1   'Contract    '0   'Summary '3 Program  '1  links
    'Else
        ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
        If (ilRet = CP_MSG_NONE) Then
            ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
            ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
            igRptCallType = Val(slStr)      'Function ID (what function calling this report if )
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
    Dim ilEnable As Integer
    Dim ilLoop As Integer
    Dim ilBoxSelected As Integer
    
    ilEnable = True
    
    ilBoxSelected = True
    If lbcSelection(0).SelCount <= 0 Then
        ilBoxSelected = False
        ilEnable = False
        'Exit For
    End If
    
    If ilBoxSelected Then
        If (edcStartYear.Text = "") Then
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
    Unload RptSelAvgCmp
    igManUnload = NO
End Sub

'Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = KEYF1 Then    'Functio key 1= Help
'        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
'        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
'        'Traffic!cdcSetup.Action = 6
'    End If
'End Sub

Private Sub plcAllTypes_Paint()
    plcAllTypes.Cls
    plcAllTypes.CurrentX = 0
    plcAllTypes.CurrentY = 0
    plcAllTypes.Print "Select"
End Sub

Private Sub plcBillOrCollect_Paint()
    plcBillOrCollect.Cls
    plcBillOrCollect.CurrentX = 0
    plcBillOrCollect.CurrentY = 0
    plcBillOrCollect.Print "Pay On"
End Sub

Private Sub plcTotalsBy_Paint()
    plcTotalsBy.Cls
    plcTotalsBy.CurrentX = 0
    plcTotalsBy.CurrentY = 0
    plcTotalsBy.Print "Totals by"
End Sub

Private Sub rbcAvgRatePrice_Click(Index As Integer)
    rbcAvgBy(0).Enabled = False: rbcAvgBy(1).Enabled = False: ckcShowUnitPrice = False: ckcShowUnitPrice.Visible = False
    If Index = 1 Then
        rbcAvgBy(0).Enabled = True: rbcAvgBy(1).Enabled = True
    Else
        ckcShowUnitPrice.Visible = True
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
