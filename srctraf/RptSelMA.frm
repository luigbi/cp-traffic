VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelMA 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Margin Acquisition Report Selection"
   ClientHeight    =   6405
   ClientLeft      =   210
   ClientTop       =   1860
   ClientWidth     =   11475
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
   ScaleHeight     =   6405
   ScaleWidth      =   11475
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   8295
      TabIndex        =   63
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
      TabIndex        =   31
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
      TabIndex        =   28
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4920
      Top             =   1680
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
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   13
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
         TabIndex        =   10
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
         TabIndex        =   12
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Margin Acquisition Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4920
      Left            =   45
      TabIndex        =   20
      Top             =   1440
      Width           =   10890
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
         Height          =   4635
         Left            =   120
         ScaleHeight     =   4635
         ScaleWidth      =   5625
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   5625
         Begin VB.CheckBox ckcSkipSort3 
            Caption         =   "New Page"
            Height          =   210
            Left            =   3000
            TabIndex        =   46
            Top             =   3390
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox ckcSkipSort2 
            Caption         =   "New Page"
            Height          =   210
            Left            =   3000
            TabIndex        =   44
            Top             =   2790
            Width           =   1215
         End
         Begin VB.CheckBox ckcSkipSort1 
            Caption         =   "New Page"
            Height          =   210
            Left            =   3000
            TabIndex        =   40
            Top             =   2190
            Width           =   1215
         End
         Begin V81TrafficReports.CSI_Calendar calStartDate 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            Text            =   "6/17/2011"
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BorderStyle     =   1
            CSI_ShowDropDownOnFocus=   -1  'True
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
            CSI_ForceMondaySelectionOnly=   -1  'True
            CSI_AllowBlankDate=   0   'False
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   2
         End
         Begin VB.ComboBox cbcSort3 
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
            Left            =   120
            TabIndex        =   45
            Top             =   3330
            Width           =   2580
         End
         Begin VB.ComboBox cbcSort2 
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
            Left            =   120
            TabIndex        =   43
            Top             =   2730
            Width           =   2580
         End
         Begin VB.ComboBox cbcSort1 
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
            Left            =   120
            TabIndex        =   39
            Top             =   2130
            Width           =   2580
         End
         Begin VB.PictureBox PlcProposals 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   5535
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   480
            Width           =   5535
            Begin VB.CheckBox ckcProposals 
               Caption         =   "Unapproved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3720
               TabIndex        =   18
               Top             =   0
               Value           =   1  'Checked
               Width           =   1560
            End
            Begin VB.CheckBox ckcProposals 
               Caption         =   "Completed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2280
               TabIndex        =   17
               Top             =   0
               Value           =   1  'Checked
               Width           =   1320
            End
            Begin VB.CheckBox ckcProposals 
               Caption         =   "Working"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1080
               TabIndex        =   16
               Top             =   0
               Value           =   1  'Checked
               Width           =   1080
            End
         End
         Begin VB.PictureBox plcCTypes 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   5535
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   840
            Width           =   5535
            Begin VB.CheckBox ckcCType 
               Caption         =   "Non-Polit"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   11
               Left            =   3000
               TabIndex        =   36
               Top             =   720
               Value           =   1  'Checked
               Width           =   1380
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   19
               Top             =   -30
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1920
               TabIndex        =   21
               Top             =   -30
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Polit"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2160
               TabIndex        =   35
               Top             =   705
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   960
               TabIndex        =   23
               Top             =   225
               Value           =   1  'Checked
               Width           =   1200
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   2160
               TabIndex        =   25
               Top             =   225
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Remnant"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   3480
               TabIndex        =   26
               Top             =   225
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   960
               TabIndex        =   27
               Top             =   450
               Value           =   1  'Checked
               Width           =   735
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Per Inquiry"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   1680
               TabIndex        =   29
               Top             =   450
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   3120
               TabIndex        =   30
               Top             =   450
               Width           =   825
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   3960
               TabIndex        =   32
               Top             =   450
               Width           =   975
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Trades"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   960
               TabIndex        =   33
               Top             =   705
               Value           =   1  'Checked
               Width           =   1020
            End
         End
         Begin VB.TextBox edcNoWeeks 
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
            MaxLength       =   2
            TabIndex        =   15
            Top             =   60
            Width           =   345
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
            Left            =   3000
            TabIndex        =   51
            Top             =   3930
            Width           =   2220
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
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   53
            Top             =   3930
            Width           =   1170
         End
         Begin VB.Label lacSortVG 
            Appearance      =   0  'Flat
            Caption         =   "Vehicle Group"
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
            Height          =   210
            Left            =   3000
            TabIndex        =   52
            Top             =   3705
            Width           =   1290
         End
         Begin VB.Label lacSort3 
            Appearance      =   0  'Flat
            Caption         =   "Sort Field #3"
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
            Height          =   210
            Left            =   120
            TabIndex        =   50
            Top             =   3090
            Width           =   1275
         End
         Begin VB.Label lacSort2 
            Appearance      =   0  'Flat
            Caption         =   "Sort Field #2"
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
            Height          =   210
            Left            =   120
            TabIndex        =   49
            Top             =   2490
            Width           =   1155
         End
         Begin VB.Label lacSort1 
            Appearance      =   0  'Flat
            Caption         =   "Sort Field #1 (major to minor)"
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
            Height          =   210
            Left            =   120
            TabIndex        =   48
            Top             =   1890
            Width           =   1875
         End
         Begin VB.Label lacContract 
            Caption         =   "Contract #"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   3990
            Width           =   975
         End
         Begin VB.Label lacStartDate 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   825
         End
         Begin VB.Label lacNoWeeks 
            Appearance      =   0  'Flat
            Caption         =   "# Weeks"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2520
            TabIndex        =   37
            Top             =   120
            Width           =   765
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
         Height          =   4740
         Left            =   5640
         ScaleHeight     =   4740
         ScaleWidth      =   5130
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   5130
         Begin VB.CheckBox ckcAllSlsp 
            Caption         =   "All Salespeople"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2760
            TabIndex        =   60
            Top             =   2460
            Width           =   1905
         End
         Begin VB.CheckBox ckcAllGroupItems 
            Caption         =   "All Group Items"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2760
            TabIndex        =   56
            Top             =   120
            Width           =   1905
         End
         Begin VB.CheckBox ckcAllAdvt 
            Caption         =   "All Advertisers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   2460
            Width           =   1905
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1920
            Index           =   3
            ItemData        =   "RptSelMA.frx":0000
            Left            =   2760
            List            =   "RptSelMA.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   57
            Top             =   480
            Width           =   2175
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1920
            Index           =   2
            ItemData        =   "RptSelMA.frx":0004
            Left            =   240
            List            =   "RptSelMA.frx":0006
            MultiSelect     =   2  'Extended
            TabIndex        =   55
            Top             =   480
            Width           =   4635
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1920
            Index           =   1
            ItemData        =   "RptSelMA.frx":0008
            Left            =   2760
            List            =   "RptSelMA.frx":000A
            MultiSelect     =   2  'Extended
            TabIndex        =   61
            Top             =   2760
            Width           =   2175
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1920
            Index           =   0
            ItemData        =   "RptSelMA.frx":000C
            Left            =   240
            List            =   "RptSelMA.frx":000E
            MultiSelect     =   2  'Extended
            TabIndex        =   59
            Top             =   2760
            Width           =   2175
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   120
            Width           =   1425
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   8280
      TabIndex        =   64
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   7920
      TabIndex        =   62
      Top             =   120
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
         Width           =   1095
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   10680
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RptSelMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptSelMA.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelMA.Frm   Margin Acquisition report
'
' Release: 6.0
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection(2))
Dim imSetAllAdvt As Integer 'True=Set list box; False= don't change list box
Dim imAllClickedAdvt As Integer  'True=All box clicked (don't call ckcAll within lbcSelection(0))
Dim imSetAllGroupItems As Integer 'True=Set list box; False= don't change list box
Dim imAllClickedGroupItems As Integer  'True=All box clicked (don't call ckcAll within lbcSelection(3))
Dim imSetAllSlsp As Integer 'True=Set list box; False= don't change list box
Dim imAllClickedSlsp As Integer  'True=All box clicked (don't call ckcAll Slspwithin lbcSelection(1))

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
Dim imTerminate As Integer
Dim imSort1 As Integer
Dim imSort2 As Integer
Dim imSort3 As Integer
Dim imPrevSort1 As Integer
Dim imPrevSort2 As Integer
Dim imPrevSort3 As Integer
Private Sub calStartDate_CalendarChanged()
    mSetCommands
End Sub


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

Private Sub cbcSet1_Click()
Dim ilLoop As Integer
Dim ilSetIndex As Integer
Dim ilRet As Integer

    ilLoop = cbcSet1.ListIndex
    ilSetIndex = gFindVehGroupInx(ilLoop, tgVehicleSets1())
    If ilSetIndex > 0 Then
        smVehGp5CodeTag = ""
        ilRet = gPopMnfPlusFieldsBox(RptSelMA, lbcSelection(3), tgSOCode(), smVehGp5CodeTag, "H" & Trim$(str$(ilSetIndex)))
        lbcSelection(3).Visible = True
        ckcAllGroupItems.Visible = True
        lbcSelection(2).Width = 1995
        ckcAllGroupItems.Value = vbUnchecked
    Else
        lbcSelection(3).Visible = False
        ckcAllGroupItems.Visible = False
        lbcSelection(2).Width = 4395
        ckcAllGroupItems.Value = vbUnchecked
    End If
    mSetCommands
End Sub

Private Sub cbcSort1_Click()
    imSort1 = cbcSort1.ListIndex
    If imSort1 = SORT1_GROUP Or imSort2 = SORT2_GROUP Then     'neither sort options 1 or 2 using vehicle groups
        cbcSet1.Visible = True
        lacSortVG.Visible = True
        ckcAllGroupItems.Visible = True
        lbcSelection(3).Visible = True
    Else
        cbcSet1.Visible = False
        lacSortVG.Visible = False
        ckcAllGroupItems.Value = vbChecked
        ckcAllGroupItems.Visible = False
        lbcSelection(3).Visible = False
    End If
    'If (imSort1 = imSort2 - 1) Or ((imSort3 = SORT3_ADVCNT Or imSort3 = SORT3_ADVCNTVEH Or imSort3 = SORT3_VEHCNT) And (imSort1 = SORT1_ADVCNT Or imSort1 = SORT1_ADVCNTVEH Or imSort1 = SORT1_VEHCNT)) Then
    If ((imSort1 = imSort2) And (imSort1 <> SORT1_NONE)) Then
   
        MsgBox "Same sort field selected, choose another", vbOKOnly, "Margin Acquisition"
        cmcGen.Enabled = False
    Else
        imPrevSort1 = imSort1
        mSetCommands
    End If
End Sub
Private Sub cbcSort2_Click()
    imSort2 = cbcSort2.ListIndex
    If imSort1 = SORT1_GROUP Or imSort2 = SORT2_GROUP Then     'neither sort options 1 or 2 using vehicle groups
        cbcSet1.Visible = True
        lacSortVG.Visible = True
        ckcAllGroupItems.Visible = True
        lbcSelection(3).Visible = True
    Else
        cbcSet1.Visible = False
        lacSortVG.Visible = False
        ckcAllGroupItems.Value = vbChecked
        ckcAllGroupItems.Visible = False
        lbcSelection(3).Visible = False
    End If
   
    'If (imSort1 = imSort2 - 1) Or ((imSort3 = SORT3_ADVCNT Or imSort3 = SORT3_ADVCNTVEH Or imSort3 = SORT3_VEHCNT) And (imSort2 = SORT2_ADVCNT Or imSort2 = SORT2_ADVCNTVEH Or imSort2 = SORT2_VEHCNT)) Then
    If ((imSort1 = imSort2) And (imSort2 <> SORT2_NONE)) Then
   
        MsgBox "Same sort field selected, choose another", vbOKOnly, "Margin Acquisition"
        cmcGen.Enabled = False
    Else
        imPrevSort2 = imSort2
        mSetCommands
    End If
End Sub
Private Sub cbcSort3_Click()
'    imSort3 = cbcSort3.ListIndex
'    If (imSort3 = SORT3_ADVCNT Or imSort3 = SORT3_ADVCNTVEH Or imSort3 = SORT3_VEHCNT) And ((imSort2 = SORT2_ADVCNT Or imSort2 = SORT2_ADVCNTVEH Or imSort2 = SORT2_VEHCNT) Or (imSort1 = SORT1_ADVCNT Or imSort1 = SORT1_ADVCNTVEH Or imSort1 = SORT1_VEHCNT)) Then
'        MsgBox "Same sort field selected, choose another", vbOKOnly, "Margin Acquisition"
'        cmcGen.Enabled = False
'    Else
        imPrevSort3 = imSort3
        mSetCommands
'    End If
    
End Sub

Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(2).hwnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    imAllClicked = False
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcAllAdvt_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllAdvt.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllAdvt Then
        imAllClickedAdvt = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).hwnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    imAllClickedAdvt = False
    mSetCommands
End Sub

Private Sub ckcAllGroupItems_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllGroupItems.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllGroupItems Then
        imAllClickedGroupItems = True
        llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(3).hwnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    imAllClickedGroupItems = False
    mSetCommands
End Sub

Private Sub ckcAllSlsp_Click()
   'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllSlsp.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllSlsp Then
        imAllClickedSlsp = True
        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(1).hwnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    imAllClickedSlsp = False
    mSetCommands
End Sub

Private Sub ckcCType_Click(Index As Integer)
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
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
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
        If Not gGenReportMA() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If

        ilRet = gCmcGenMA(imGenShiftKey, smLogUserCode)
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
        gCreateMarginAcquisition
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
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-10
        End If
    Next ilJobs
    imGenShiftKey = 0

    Screen.MousePointer = vbHourglass
    gCRGrfClear
    Screen.MousePointer = vbDefault
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
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

Private Sub edcNoWeeks_Change()
    mSetCommands
End Sub

Private Sub edcNoWeeks_GotFocus()
    gCtrlGotFocus edcNoWeeks
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
    RptSelMA.Refresh
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
    'RptSelMA.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgRptSelBudgetCodeSP
    Erase tgClfSP
    Erase tgCffSP
    Erase imCodes
    PECloseEngine
    
    Set RptSelMA = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    If Index = 0 Then               'advt list box
        If Not imAllClickedAdvt Then
            imSetAllAdvt = False
            ckcAllAdvt.Value = vbUnchecked  'False
            imSetAllAdvt = True
        End If
    ElseIf Index = 1 Then           'slsp list box
        If Not imAllClickedSlsp Then
            imSetAllSlsp = False
            ckcAllSlsp.Value = vbUnchecked  'False
            imSetAllSlsp = True
        End If
    ElseIf Index = 2 Then           'vehicles
        If Not imAllClicked Then
            imSetAll = False
            ckcAll.Value = vbUnchecked  'False
            imSetAll = True
        End If
    ElseIf Index = 3 Then           'vehicle groups
        If Not imAllClickedGroupItems Then
            imSetAllGroupItems = False
            ckcAllGroupItems.Value = vbUnchecked  'False
            imSetAllGroupItems = True
        Else
            ckcAllGroupItems.Value = False
        End If
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

    RptSelMA.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    imAllClickedGroupItems = False
    imSetAllGroupItems = True
    imAllClickedAdvt = False
    imSetAllAdvt = True
    imAllClickedSlsp = False
    imSetAllSlsp = True
   
    lbcSelection(2).Width = 4395            'default vehicles to width of list box area, unless group items required

    gCenterStdAlone RptSelMA
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
    'cbcWhenDay.AddItem "One Time"
    'cbcWhenDay.AddItem "Every M-F"
    'cbcWhenDay.AddItem "Every M-Sa"
    'cbcWhenDay.AddItem "Every M-Su"
    'cbcWhenDay.AddItem "Every Monday"
    'cbcWhenDay.AddItem "Every Tuesday"
    'cbcWhenDay.AddItem "Every Wednesday"
    'cbcWhenDay.AddItem "Every Thursday"
    'cbcWhenDay.AddItem "Every Friday"
    'cbcWhenDay.AddItem "Every Saturday"
    'cbcWhenDay.AddItem "Every Sunday"
    'cbcWhenDay.AddItem "Cal Month End+1"
    'cbcWhenDay.AddItem "Cal Month End+2"
    'cbcWhenDay.AddItem "Cal Month End+3"
    'cbcWhenDay.AddItem "Cal Month End+4"
    'cbcWhenDay.AddItem "Cal Month End+5"
    'cbcWhenDay.AddItem "Std Month End+1"
    'cbcWhenDay.AddItem "Std Month End+2"
    'cbcWhenDay.AddItem "Std Month End+3"
    'cbcWhenDay.AddItem "Std Month End+4"
    'cbcWhenDay.AddItem "Std Month End+5"
    'cbcWhenDay.ListIndex = 0
    'cbcWhenTime.AddItem "Right Now"
    'cbcWhenTime.AddItem "at 10PM"
    'cbcWhenTime.AddItem "at 12AM"
    'cbcWhenTime.AddItem "at 2AM"
    'cbcWhenTime.AddItem "at 4AM"
    'cbcWhenTime.AddItem "at 6AM"
    'cbcWhenTime.ListIndex = 0
    'Setup report output types
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    gPopExportTypes cbcFileType     '10-20-01

    If Not gObtainAgency() Then
        MsgBox "The Agency File Could Not Be Opened "
        Exit Sub
    End If
    
    ilRet = gRptVehPop(RptSelMA, lbcSelection(2), VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHREP_W_CLUSTER + VEHREP_WO_CLUSTER + ACTIVEVEH)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopErr
        gCPErrorMsg ilRet, "mInitReport (gPopUserVehicleBox)", RptSelMA
        On Error GoTo 0
    End If
    
    ilRet = gRptAdvtPop(RptSelMA, lbcSelection(0))      'populate advt
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopErr
        gCPErrorMsg ilRet, "mInitReport (gPopAdvtBox)", RptSelMA
        On Error GoTo 0
    End If
    
    ilRet = gRptSPersonPop(RptSelMA, lbcSelection(1))   'populate salespeople
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopErr
        gCPErrorMsg ilRet, "mInitReport (gPopSalespersonBox)", RptSelMA
        On Error GoTo 0
    End If
    
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    ckcAll.Visible = True
 
    sgMnfVehGrpTag = ""
    gPopVehicleGroups RptSelMA!cbcSet1, tgVehicleSets1(), False
    
'    cbcSort1.Clear
'    cbcSort1.AddItem "Advt/Prod, Contract"
'    cbcSort1.AddItem "Advt/Prod, Contract, Vehicle"
'    cbcSort1.AddItem "Salesperson"
'    cbcSort1.AddItem "Vehicle, Advt/Prod"
'    cbcSort1.AddItem "Vehicle Group"
'    cbcSort1.ListIndex = 0
'    imSort1 = 0
'
'    cbcSort2.Clear
'    cbcSort2.AddItem "None"
'    cbcSort2.AddItem "Advt/Prod, Contract"
'    cbcSort2.AddItem "Advt/Prod, Contract, Vehicle"
'    cbcSort2.AddItem "Salesperson"
'    cbcSort2.AddItem "Vehicle, Advt/Prod"
'    cbcSort2.AddItem "Vehicle Group"
'    cbcSort2.ListIndex = 0
'    imSort2 = 0

'    cbcSort3.Clear
'    cbcSort3.AddItem "None"
'    cbcSort3.AddItem "Advt/Prod, Contract"
'    cbcSort3.AddItem "Advt/Prod, Contract, Vehicle"
'    cbcSort3.AddItem "Vehicle, Advt/Prod"
'    cbcSort3.ListIndex = 0
'    imSort3 = 0

    cbcSort1.Clear
    cbcSort1.AddItem "None"
    cbcSort1.AddItem "Salesperson"
    cbcSort1.AddItem "Vehicle Group"
    cbcSort1.ListIndex = 1
    imSort1 = 1

    cbcSort2.Clear
    cbcSort2.AddItem "None"
    cbcSort2.AddItem "Salesperson"
    cbcSort2.AddItem "Vehicle Group"
    cbcSort2.ListIndex = 2
    imSort2 = 2

    
    cbcSort3.Clear
    cbcSort3.AddItem "Advt/Prod, Contract"
    cbcSort3.AddItem "Advt/Prod, Contract, Vehicle"
    cbcSort3.AddItem "Vehicle, Advt/Prod"
    cbcSort3.ListIndex = 0
    imSort3 = 0
    
    pbcSelC.Visible = True
    pbcOption.Visible = True
    
    mSetCommands
    calStartDate.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
mPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
            MsgBox "Application must be run from the Traffic application", vbCritical, "Reports"
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
    'gInitStdAlone RptSelMA, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Sales vs Plan"
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
    Dim ilSetIndex As Integer
    
    ilEnable = False
    If (calStartDate.Text <> "") And (edcNoWeeks.Text <> "") Then
        ilEnable = False
        If lbcSelection(0).SelCount > 0 And lbcSelection(1).SelCount > 0 And lbcSelection(2).SelCount > 0 Then       'at leat 1 advt & 1 vehicle & slsp must be selected
            ilEnable = True
            'determine if a vehicle group has been selected
            ilLoop = cbcSet1.ListIndex
            ilSetIndex = gFindVehGroupInx(ilLoop, tgVehicleSets1())

            If ilSetIndex > 0 Then
                If lbcSelection(3).SelCount <= 0 Then
                    ilEnable = False
                End If
            End If
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
    Unload RptSelMA
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub



Private Sub plcCTypes_Paint()
    plcCTypes.CurrentX = 0
    plcCTypes.CurrentY = 0
    plcCTypes.Print "Contracts"
End Sub


Private Sub PlcProposals_Paint()
    PlcProposals.CurrentX = 0
    PlcProposals.CurrentY = 0
    PlcProposals.Print "Proposals"
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



