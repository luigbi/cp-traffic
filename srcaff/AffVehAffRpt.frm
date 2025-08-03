VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVehAffRpt 
   Caption         =   "Affiliate Agreement Report"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   9360
   Icon            =   "AffVehAffRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   9360
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4200
      Top             =   375
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6780
      FormDesignWidth =   9360
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      TabIndex        =   10
      Top             =   1650
      Width           =   8895
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6000
         Picture         =   "AffVehAffRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Select Stations from File.."
         Top             =   2640
         Width           =   360
      End
      Begin VB.CheckBox ckcSortSelection 
         Caption         =   "All"
         Height          =   255
         Left            =   6720
         TabIndex        =   57
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ListBox lbcSortSelection 
         Height          =   1815
         ItemData        =   "AffVehAffRpt.frx":0E34
         Left            =   6720
         List            =   "AffVehAffRpt.frx":0E3B
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   56
         Top             =   3000
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.ComboBox cbcSort 
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   450
         Width           =   1485
      End
      Begin VB.CheckBox ckcShowComments 
         Caption         =   "Show Comments"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   4200
         Width           =   1515
      End
      Begin VB.CheckBox ckcContactInfo 
         Caption         =   "Show Contact Info"
         Height          =   255
         Left            =   1920
         TabIndex        =   52
         Top             =   4200
         Width           =   1845
      End
      Begin VB.Frame frcService 
         Caption         =   "Service Agreements"
         Height          =   750
         Left            =   2640
         TabIndex        =   49
         Top             =   3120
         Width           =   1695
         Begin VB.CheckBox ckcService 
            Caption         =   "Service"
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   435
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.CheckBox ckcService 
            Caption         =   "Non-Service"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   210
            Value           =   1  'Checked
            Width           =   1365
         End
      End
      Begin VB.CheckBox chkPledgeOrPgm 
         Caption         =   "Show Program Times (vs Avail Pledge Times)"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   4680
         Width           =   3600
      End
      Begin VB.CheckBox ckcMulticastOnly 
         Caption         =   "Multicast Only"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   4440
         Width           =   1365
      End
      Begin VB.Frame IncludeDormant 
         Caption         =   "Dormant Vehicles"
         Height          =   585
         Left            =   2280
         TabIndex        =   26
         Top             =   2535
         Width           =   2055
         Begin VB.OptionButton rbcDormVeh 
            Caption         =   "No"
            Height          =   210
            Index           =   1
            Left            =   960
            TabIndex        =   28
            Top             =   240
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton rbcDormVeh 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame frcExpired 
         Caption         =   "Expired Agreements"
         Height          =   585
         Left            =   120
         TabIndex        =   31
         Top             =   2535
         Width           =   2100
         Begin VB.OptionButton rbcInclExpired 
            Caption         =   "No"
            Height          =   210
            Index           =   1
            Left            =   960
            TabIndex        =   25
            Top             =   240
            Width           =   700
         End
         Begin VB.OptionButton rbcInclExpired 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin VB.CheckBox ckcStationInfo 
         Caption         =   "Show Station Information"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   33
         Top             =   4440
         Width           =   2265
      End
      Begin VB.ListBox lbcStations 
         Height          =   1815
         ItemData        =   "AffVehAffRpt.frx":0E42
         Left            =   4560
         List            =   "AffVehAffRpt.frx":0E49
         MultiSelect     =   2  'Extended
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3000
         Width           =   4020
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   4560
         TabIndex        =   37
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1815
         ItemData        =   "AffVehAffRpt.frx":0E50
         Left            =   4560
         List            =   "AffVehAffRpt.frx":0E52
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   660
         Width           =   4020
      End
      Begin VB.Frame Frame5 
         Caption         =   "Show"
         Height          =   975
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   2475
         Begin VB.OptionButton optDatePhone 
            Caption         =   "Air Dates + Pledge Days"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   59
            Top             =   660
            Width           =   2205
         End
         Begin VB.OptionButton optDatePhone 
            Caption         =   "Format"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   1110
            TabIndex        =   58
            Top             =   435
            Width           =   1275
         End
         Begin VB.OptionButton optDatePhone 
            Caption         =   "Export Codes"
            Height          =   255
            Index           =   2
            Left            =   1110
            TabIndex        =   48
            Top             =   210
            Width           =   1275
         End
         Begin VB.OptionButton optDatePhone 
            Caption         =   "Phone #'s"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   435
            Width           =   1215
         End
         Begin VB.OptionButton optDatePhone 
            Caption         =   "Air Dates"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   210
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Type"
         Height          =   540
         Left            =   1680
         TabIndex        =   11
         Top             =   210
         Width           =   2655
         Begin VB.OptionButton optSP 
            Caption         =   "People"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   13
            Top             =   240
            Width           =   795
         End
         Begin VB.OptionButton optSP 
            Caption         =   "Stations"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.OptionButton optSP 
            Caption         =   "Both"
            Height          =   195
            Index           =   2
            Left            =   1920
            TabIndex        =   14
            Top             =   240
            Width           =   700
         End
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4560
         TabIndex        =   34
         Top             =   360
         Width           =   1935
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   270
         Left            =   3210
         TabIndex        =   16
         Top             =   1065
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/12/2022"
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
      Begin V81Affiliate.CSI_Calendar CalEnterFrom 
         Height          =   270
         Left            =   1605
         TabIndex        =   17
         Top             =   1425
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/12/2022"
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
         CSI_DefaultDateType=   2
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   270
         Left            =   1605
         TabIndex        =   15
         Top             =   1065
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/12/2022"
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
      Begin V81Affiliate.CSI_Calendar CalEnterTo 
         Height          =   270
         Left            =   3210
         TabIndex        =   18
         Top             =   1425
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/12/2022"
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
         CSI_DefaultDateType=   2
      End
      Begin V81Affiliate.CSI_Calendar CalStartBetween1 
         Height          =   270
         Left            =   1605
         TabIndex        =   19
         Top             =   1775
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/12/2022"
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
         CSI_DefaultDateType=   2
      End
      Begin V81Affiliate.CSI_Calendar CalStartBetween2 
         Height          =   270
         Left            =   3210
         TabIndex        =   20
         Top             =   1775
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/12/2022"
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
         CSI_DefaultDateType=   2
      End
      Begin V81Affiliate.CSI_Calendar CalEndBetween1 
         Height          =   270
         Left            =   1605
         TabIndex        =   21
         Top             =   2100
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/12/2022"
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
         CSI_DefaultDateType=   2
      End
      Begin V81Affiliate.CSI_Calendar CalEndBetween2 
         Height          =   270
         Left            =   3210
         TabIndex        =   22
         Top             =   2100
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "12/12/2022"
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
         CSI_DefaultDateType=   2
      End
      Begin VB.Frame frcDates 
         Caption         =   "Agreements"
         Height          =   1695
         Left            =   120
         TabIndex        =   36
         Top             =   810
         Width           =   4215
         Begin VB.Label lacEndthru 
            Caption         =   "And"
            Height          =   285
            Left            =   2625
            TabIndex        =   46
            Top             =   1335
            Width           =   375
         End
         Begin VB.Label lacStartthru 
            Caption         =   "And"
            Height          =   285
            Left            =   2625
            TabIndex        =   45
            Top             =   960
            Width           =   435
         End
         Begin VB.Label lacEndBetween 
            Caption         =   "Ending between"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   1335
            Width           =   1650
         End
         Begin VB.Label lacStartBetween 
            Caption         =   "Starting between"
            Height          =   285
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lacTo 
            Caption         =   "To"
            Height          =   225
            Left            =   2625
            TabIndex        =   42
            Top             =   615
            Width           =   375
         End
         Begin VB.Label lacFrom 
            Caption         =   "Entered- From"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   615
            Width           =   1080
         End
         Begin VB.Label Label4 
            Caption         =   "End"
            Height          =   225
            Left            =   2625
            TabIndex        =   40
            Top             =   255
            Width           =   315
         End
         Begin VB.Label Label3 
            Caption         =   "Active- Start"
            Height          =   225
            Left            =   120
            TabIndex        =   38
            Top             =   255
            Width           =   1035
         End
      End
      Begin VB.Label lacSortBy 
         Caption         =   "Sort by-"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   195
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   5355
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   5115
      TabIndex        =   8
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   225
      Width           =   2685
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   255
      TabIndex        =   0
      Top             =   0
      Width           =   3585
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffVehAffRpt.frx":0E54
         Left            =   1335
         List            =   "AffVehAffRpt.frx":0E56
         TabIndex        =   4
         Top             =   765
         Width           =   2040
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Mail List"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1125
         Width           =   1335
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   825
         Width           =   870
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Value           =   -1  'True
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmVehAffRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*  frmVehAffRpt - compares vehicles and their affiliations, sorted by either one
'*
'*  Created July,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'
'   8-11-04 Add selectivity to gather agreements starting between X & Y dates;
'           Add selectivity to gather agreements ending between X & Y dates
'****************************************************************************
Option Explicit

Private hmMail As Integer
Private smToFile As String
Private imChkStationIgnore As Integer
Private imChkListBoxIgnore As Integer
Private imChkListOtherIgnore As Integer
Private imSortBy As Integer
Private bmSortListTest As Boolean
Private tmAmr As AMR

Private rst_Agreement As ADODB.Recordset
Private Const SORTBY_AUD = 0
Private Const SORTBY_MKTNAME = 1
Private Const SORTBY_MKTRANK = 2
Private Const SORTBY_OWNER_STATION = 3
Private Const SORTBY_OWNER_VEHICLE = 4
Private Const SORTBY_STATION = 5
Private Const SORTBY_VEHICLE = 6

Private Sub mEnableDisableContactStationInfo()
    ' Enable Show Station and Contact Information if STATION and FORMAT are selected    Date: 8/8/2018  FYM
    If (optDatePhone(3).Value = True And cbcSort.ListIndex = 5) Then
        ckcStationInfo.Enabled = False: ckcStationInfo.Value = vbUnchecked
        ckcContactInfo.Enabled = False: ckcContactInfo.Value = vbUnchecked
        ckcShowComments.Enabled = False: ckcShowComments.Value = vbUnchecked
    Else
        ckcStationInfo.Enabled = True
        ckcContactInfo.Enabled = True
        ckcShowComments.Enabled = True
        'optDatePhone(3).Enabled = False
        optDatePhone(3).Visible = True
        optDatePhone(3).Value = False
    End If
End Sub

Private Sub mEnableGenerateReportButton()
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/13/2018  FYM
    If ckcSortSelection.Visible Then
'        If (CalOnAirDate.Text <> "" And CalOffAirDate.Text <> "" And _
'            ((lbcVehAff.SelCount > 0 And lbcStations.SelCount > 0 And lbcSortSelection.SelCount > 0) Or _
'            (ckcAllStations.Value = vbChecked And chkListBox.Value = vbChecked And ckcSortSelection.Value = vbChecked))) Then
         If ((lbcVehAff.SelCount > 0 And lbcStations.SelCount > 0 And lbcSortSelection.SelCount > 0) Or _
            (ckcAllStations.Value = vbChecked And chkListBox.Value = vbChecked And ckcSortSelection.Value = vbChecked)) Then
            cmdReport.Enabled = True
        Else
            cmdReport.Enabled = False
        End If
    Else
'        If (CalOnAirDate.Text <> "" And CalOffAirDate.Text <> "" And _
'            ((lbcVehAff.SelCount > 0 And lbcStations.SelCount > 0) Or _
'            (ckcAllStations.Value = vbChecked And chkListBox.Value = vbChecked))) Then
         If ((lbcVehAff.SelCount > 0 And lbcStations.SelCount > 0) Or _
            (ckcAllStations.Value = vbChecked And chkListBox.Value = vbChecked)) Then
            cmdReport.Enabled = True
        Else
            cmdReport.Enabled = False
        End If
    End If
End Sub

Private Sub CalOffAirDate_CalendarChanged()
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub CalOnAirDate_CalendarChanged()
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub cbcSort_Click()
Dim ilLoop As Integer
    Dim iIndex As Integer
    
    lbcSortSelection.Clear
    ckcSortSelection.Visible = False
    ckcSortSelection.Value = vbUnchecked
    lbcSortSelection.Visible = False
    lbcStations.Width = lbcVehAff.Width
    ckcStationInfo.Enabled = True

    bmSortListTest = False
    imSortBy = cbcSort.ListIndex
 
    If imSortBy = SORTBY_STATION Then
        'added SHOW option: FORMAT; should be enable/visible when SORT = STATION, VEHICLE      Date: 8/8/2018   FYM
        optDatePhone(3).Enabled = True
        'optDatePhone(3).Visible = True
        
        ' Enable Show Station and Contact Information if STATION and FORMAT are selected    Date: 8/8/2018  FYM
        mEnableDisableContactStationInfo
    ElseIf imSortBy = SORTBY_MKTNAME Then
        ckcSortSelection.Caption = "All Markets"
        For ilLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
            lbcSortSelection.AddItem Trim$(tgMarketInfo(ilLoop).sName)
            lbcSortSelection.ItemData(lbcSortSelection.NewIndex) = tgMarketInfo(ilLoop).lCode
        Next ilLoop
        ckcSortSelection.Visible = True
        lbcSortSelection.Visible = True
        lbcStations.Width = (lbcVehAff.Width / 2) - 360     'make station list half width in order to show the market list
        bmSortListTest = True
    
        'added SHOW option: FORMAT; should be enabled/visible when SORT = STATION, VEHICLE      Date: 8/8/2018   FYM
        optDatePhone(3).Enabled = False
        optDatePhone(3).Visible = True
        optDatePhone(3).Value = False
    ElseIf imSortBy = SORTBY_OWNER_STATION Or imSortBy = SORTBY_OWNER_VEHICLE Then

        ckcSortSelection.Caption = "All Owners"
        For ilLoop = 0 To UBound(tgOwnerInfo) - 1 Step 1
            lbcSortSelection.AddItem Trim$(tgOwnerInfo(ilLoop).sName)
            lbcSortSelection.ItemData(lbcSortSelection.NewIndex) = tgOwnerInfo(ilLoop).lCode
        Next ilLoop
        ckcSortSelection.Visible = True
        lbcSortSelection.Visible = True
        lbcStations.Width = (lbcVehAff.Width / 2) - 360     'make station list half width in order to show the owner list
        bmSortListTest = True
        
        'added SHOW option: FORMAT; should be enabled/visible when SORT = STATION, VEHICLE      Date: 8/8/2018   FYM
        optDatePhone(3).Enabled = False
        'optDatePhone(3).Visible = True
        optDatePhone(3).Value = False
    Else
        'added SHOW option: FORMAT; should be enabled/visible when SORT = STATION, VEHICLE      Date: 8/8/2018   FYM
        mEnableDisableContactStationInfo
        If imSortBy = SORTBY_VEHICLE Then
            ckcStationInfo.Enabled = False: ckcStationInfo.Value = vbUnchecked
        End If
        optDatePhone(3).Enabled = False
        'optDatePhone(3).Visible = True
        optDatePhone(3).Value = False
    End If
    
    If optDatePhone(0).Value Then
        optDatePhone_Click 0
    ElseIf optDatePhone(1).Value Then
        optDatePhone_Click 1
    ElseIf optDatePhone(2).Value Then
        optDatePhone_Click 2
    ElseIf optDatePhone(3).Value Then
        'added SHOW option: FORMAT      Date:8/13/2018  FYM
        optDatePhone_Click 3
    End If
    
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
     
End Sub

Private Sub chkListBox_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehAff.ListCount > 0 Then
        imChkListBoxIgnore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkListBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton

End Sub

Private Sub ckcAllStations_Click()
 Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkStationIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcStations.ListCount > 0 Then
        imChkStationIgnore = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkStationIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub ckcSortSelection_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkListOtherIgnore Then
        Exit Sub
    End If
    If ckcSortSelection.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcSortSelection.ListCount > 0 Then
        imChkListOtherIgnore = True
        lRg = CLng(lbcSortSelection.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSortSelection.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkListOtherIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub cmdDone_Click()
    Unload frmVehAffRpt
End Sub

Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim ilLoop As Integer
    Dim iRet As Integer
    Dim sCode As String
    Dim bm As Variant
    Dim sName As String
    Dim slSortSelection As String           'list of selected markets or owners, if option
    Dim sVehicles As String
    Dim sStations As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateRange As String
    Dim sStationType As String
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim sMail As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    Dim slEnterFrom As String
    Dim slEnterTo As String
    Dim slStartBetween1 As String   'show all agreements whose start date is within X & Y dates
    Dim slStartBetween2 As String
    Dim slEndBetween1 As String     'show all agreements whose end dates is within X & Y dates
    Dim slEndBetween2 As String
    Dim slStartBetween As String    'sql query
    Dim slEndBetween As String      'sql query
    Dim slDescription As String
    Dim slEnteredRange As String
    Dim slMulticastOnly As String
    Dim slService As String
    Dim ilInclChoiceCodes As Integer
    Dim ilInclVehicleCodes As Integer
    Dim ilInclStationCodes As Integer
    Dim llUseChoiceCodes()  As Long
    Dim ilUseVehicleCodes() As Integer
    Dim ilUseStationCodes() As Integer
    Dim llCount As Long
    Dim blFound As Boolean
    Dim llValue As Long
    Dim llOwnerInx As Long
    Dim ilShttInx As Integer
    Dim llTemp As Long
    Dim slOwnerName As String
    Dim ilMktInx As Integer
    Dim ilMktRepInx As Integer
    Dim ilServRepInx As Integer
    Dim llVefInx As Long
        
        On Error GoTo ErrHand
        
        sStartDate = Trim$(CalOnAirDate.Text)
        If sStartDate = "" Then
            sStartDate = "1/1/1970"
        End If
        sEndDate = Trim$(CalOffAirDate.Text)
        If sEndDate = "" Then
            sEndDate = "12/31/2069"
        End If
        If gIsDate(sStartDate) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
            CalOnAirDate.SetFocus
            Exit Sub
        End If
        If gIsDate(sEndDate) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
            CalOffAirDate.SetFocus
            Exit Sub
        End If
        
        'Validate Entered From/To dates
        slEnterFrom = Trim$(CalEnterFrom.Text)
        If slEnterFrom = "" Then
            slEnterFrom = "1/1/1970"
        End If
        slEnterTo = Trim$(CalEnterTo.Text)
        If slEnterTo = "" Then
            slEnterTo = "12/31/2069"
        End If
        If gIsDate(slEnterFrom) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
            CalEnterFrom.SetFocus
            Exit Sub
        End If
        If gIsDate(slEnterTo) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
            CalEnterTo.SetFocus
            Exit Sub
        End If
        
        'Validate agreements starting between X & Y dates
        slStartBetween1 = Trim$(CalStartBetween1.Text)
        If slStartBetween1 = "" Then
            slStartBetween1 = "1/1/1970"
        End If
        slStartBetween2 = Trim$(CalStartBetween2.Text)
        If slStartBetween2 = "" Then
            slStartBetween2 = "12/31/2069"
        End If
        
        If gIsDate(slStartBetween1) = False Then
            Beep
            gMsgBox "Please enter a valid agreement start date span (m/d/yy)", vbCritical
            CalStartBetween2.SetFocus
            Exit Sub
        End If
        If gIsDate(slStartBetween2) = False Then
            Beep
            gMsgBox "Please enter a valid agreement start date span (m/d/yy)", vbCritical
            CalStartBetween2.SetFocus
            Exit Sub
        End If
        
        
         'Validate Agreemnts ending betweening  X & Y dates
        slEndBetween1 = Trim$(CalEndBetween1.Text)
        If slEndBetween1 = "" Then
            slEndBetween1 = "1/1/1970"
        End If
        slEndBetween2 = Trim$(CalEndBetween2.Text)
        If slEndBetween2 = "" Then
            slEndBetween2 = "12/31/2069"
        End If
        
        If gIsDate(slEndBetween1) = False Then
            Beep
            gMsgBox "Please enter a valid agreement end date span (m/d/yy)", vbCritical
            CalEndBetween2.SetFocus
            Exit Sub
        End If
        If gIsDate(slEndBetween2) = False Then
            Beep
            gMsgBox "Please enter a valid agreement end date span (m/d/yy)", vbCritical
            CalEndBetween2.SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        'CRpt1.Connect = "DSN = " & sgDatabaseName
      
        If optRptDest(0).Value = True Then
            ilRptDest = 0
        ElseIf optRptDest(1).Value = True Then
            ilRptDest = 1
        ElseIf optRptDest(2).Value = True Then
            ilRptDest = 2
            ilExportType = cboFileType.ListIndex    '3-15-04
            sgCrystlFormula1 = "E" 'Show Export Codes
    
        ElseIf optRptDest(3).Value = True Then
            iRet = OpenMsgFile(hmMail, smToFile)
            If iRet = False Then
                Exit Sub
            End If
        Else
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
        cmdDone.Enabled = False
        cmdReturn.Enabled = False

        gUserActivityLog "S", sgReportListName & ": Prepass"
        'Retrieve information from the list box
        If optSP(0).Value Then
            sStationType = "shttType = 0"
        ElseIf optSP(1).Value Then
            sStationType = "shttType = 1"
        Else
            sStationType = ""
        End If
        
        'get the Generation date and time to filter data for Crystal
        sgGenDate = Format$(gNow(), "m/d/yyyy")             '7-10-13 use global gen date/time for crystal filtering
        sgGenTime = Format$(gNow(), sgShowTimeWSecForm)
    
        sStartDate = Format(sStartDate, "m/d/yyyy")
        sEndDate = Format(sEndDate, "m/d/yyyy")
       
        'create sql query to get agreements starting between 2 dates
        slStartBetween = " attOnAir >=" & "'" & Format$(slStartBetween1, sgSQLDateForm) & "'" & " And attOnAir <=" & "'" + Format$(slStartBetween2, sgSQLDateForm) & "'"
        'create sql query to get agreements ending between 2 dates
        slEndBetween = " (attOffAir >=" & "'" & Format$(slEndBetween1, sgSQLDateForm) & "'" & " And attOffAir <=" & "'" + Format$(slEndBetween2, sgSQLDateForm) & "'" & ") or " & "(attDropDate >=" & "'" & Format$(slEndBetween1, sgSQLDateForm) & "'" & " And attDropDate <=" & "'" + Format$(slEndBetween2, sgSQLDateForm) & "'" & ") "
        'create sql query to get agreements active between 2 spans
        sDateRange = " attOffAir >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & " And attDropDate >=" & "'" + Format$(sStartDate, sgSQLDateForm) & "'" & " And attOnAir <=" & "'" & Format$(sEndDate, sgSQLDateForm) & "'"
        slEnteredRange = " attEnterDate >= " & "'" & Format$(slEnterFrom, sgSQLDateForm) & "'" & " And attEnterDate <= " & "'" & Format$(slEnterTo, sgSQLDateForm) & "'"
        
        sVehicles = ""
        sStations = ""
        slMulticastOnly = ""
        slService = ""
        
        If ckcMulticastOnly.Value = vbChecked Then
            slMulticastOnly = " and (attMulticast = 'Y') "
            sgCrystlFormula8 = "Y"
        Else
            sgCrystlFormula8 = "N"
        End If
        
        If ckcService(0).Value = vbChecked Then             'non-service
            slService = " and (attServiceAgreement <> 'Y') "
            sgCrystlFormula11 = "'N'"
        End If
             
        If ckcService(1).Value = vbChecked Then             'non-service
            If ckcService(0).Value = vbUnchecked Then         'include only service agreements
                slService = " and (attServiceAgreement = 'Y') "
                sgCrystlFormula11 = "'S'"
            Else            'both checked, no filter
                slService = ""
                sgCrystlFormula11 = "'B'"
            End If
        End If
        
        If chkPledgeOrPgm.Value = vbChecked Then            '1-26-12 show Pgm times (vs pledge start/end pgm time)
            sgCrystlFormula9 = "P"                          'Show program start/end times
        Else
            If optDatePhone(4).Value = True Then                   '3-4-20 show avail + pledge days
                sgCrystlFormula9 = "D"
            Else
                sgCrystlFormula9 = "A"                          'Show start time of first avail and end time of last avail from DAT
            End If
        End If
        
        If ckcContactInfo.Value = vbChecked Then            '2-20-15 show contact info
            sgCrystlFormula12 = "Y"                          'Show contact info
        Else
            sgCrystlFormula12 = "N"                          'hide contact info
        End If
        
        If ckcShowComments.Value = vbChecked Then            '2-20-15 show agreement commnets
            sgCrystlFormula13 = "Y"
        Else
            sgCrystlFormula13 = "N"
        End If
        
        
        sVehicles = ""
        sStations = ""
        slSortSelection = ""          'market or owner selection
        
        'ReDim ilUseVehicleCodes(1 To 1) As Integer
        ReDim ilUseVehicleCodes(0 To 0) As Integer
        'ReDim ilUseStationCodes(1 To 1) As Integer
        ReDim ilUseStationCodes(0 To 0) As Integer
        'ReDim llUseChoiceCodes(1 To 1) As Long
        ReDim llUseChoiceCodes(0 To 0) As Long
        gObtainCodes lbcVehAff, ilInclVehicleCodes, ilUseVehicleCodes()        'build array of which codes to incl/excl
        For ilLoop = LBound(ilUseVehicleCodes) To UBound(ilUseVehicleCodes) - 1
            If Trim$(sVehicles) = "" Then
                If ilInclVehicleCodes = True Then                          'include the list
                    sVehicles = "attvefcode IN (" & Str(ilUseVehicleCodes(ilLoop))
                Else                                                        'exclude the list
                    sVehicles = "attvefcode Not IN (" & Str(ilUseVehicleCodes(ilLoop))
                End If
            Else
                sVehicles = sVehicles & "," & Str(ilUseVehicleCodes(ilLoop))
            End If
        Next ilLoop
        If sVehicles <> "" Then
            sVehicles = sVehicles & ")"
        End If
        gObtainCodes lbcStations, ilInclStationCodes, ilUseStationCodes()        'build array of which advt codes to incl/excl
        For ilLoop = LBound(ilUseStationCodes) To UBound(ilUseStationCodes) - 1
            If Trim$(sStations) = "" Then
                If ilInclStationCodes = True Then                          'include the list
                    sStations = "attshfcode IN (" & Str(ilUseStationCodes(ilLoop))
                Else                                                        'exclude the list
                    sStations = "attshfcode Not IN (" & Str(ilUseStationCodes(ilLoop))
                End If
            Else
                sStations = sStations & "," & Str(ilUseStationCodes(ilLoop))
            End If
        Next ilLoop
        If sStations <> "" Then
            sStations = sStations & ")"
        End If
    

        If bmSortListTest Then                 '3rd list box to test
            gObtainCodesLong lbcSortSelection, ilInclChoiceCodes, llUseChoiceCodes()        'build array of which advt codes to incl/excl
    
            slSortSelection = ""
            If ckcSortSelection.Value = vbUnchecked Then
                For ilLoop = LBound(llUseChoiceCodes) To UBound(llUseChoiceCodes) - 1
                    If Trim$(slSortSelection) = "" Then
                        If ilInclChoiceCodes = True Then                          'include the list
                            slSortSelection = " IN (" & Str(llUseChoiceCodes(ilLoop))
                        Else                                                        'exclude the list
                            slSortSelection = " Not IN (" & Str(llUseChoiceCodes(ilLoop))
                        End If
                    Else
                        slSortSelection = slSortSelection & "," & Str(llUseChoiceCodes(ilLoop))
                    End If
                Next ilLoop
                If slSortSelection <> "" Then
                    slSortSelection = slSortSelection & ")"
                End If
            End If
        End If
    
    
        If optRptDest(3).Value = True Then  'Mail List
            SQLQuery = "SELECT DISTINCT  shttCallLetters, shttFax from "
        Else
            SQLQuery = "SELECT * from"
        End If
        'SQLQuery = " Select * from"
        SQLQuery = SQLQuery + " shtt INNER JOIN  att ON shttCode = attShfCode "
        SQLQuery = SQLQuery + "LEFT OUTER JOIN mkt ON shttMktCode = mktCode "
        SQLQuery = SQLQuery + " INNER JOIN   VEF_Vehicles ON attVefCode = vefCode "
        SQLQuery = SQLQuery + " Where (" & sDateRange & ")" & " and (" & slEnteredRange & ") and (" & slStartBetween & ")" & " AND (" & slEndBetween & ")"
        If rbcInclExpired(1).Value Then  'If True don't show expired agreements
            SQLQuery = SQLQuery + " AND " & "(attOffAir >=" & "'" & Format$(Date, sgSQLDateForm) & "'" & ")"
            SQLQuery = SQLQuery + " AND " & "(attDropDate >=" & "'" & Format$(Date, sgSQLDateForm) & "'" & ")"
        End If
        
        If rbcDormVeh(1).Value Then  'If True don't show include dormant vehicles
            SQLQuery = SQLQuery + " AND " & " vefstate <> 'D'"
        End If
    
        If sStationType <> "" Then
            SQLQuery = SQLQuery + " AND (" & sStationType & ")"
        End If
        If sStations <> "" Then
            SQLQuery = SQLQuery + " AND (" & sStations & ")"
        End If
        If sVehicles <> "" Then     '12-13-00
            SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
        End If
        SQLQuery = SQLQuery + slMulticastOnly + slService
        
        Set rst_Agreement = gSQLSelectCall(SQLQuery)
        llCount = 0
        While Not rst_Agreement.EOF
            
            'filter owner, market if applicable
            blFound = True
            If bmSortListTest Then
                If imSortBy = SORTBY_MKTNAME Or imSortBy = SORTBY_OWNER_VEHICLE Or imSortBy = SORTBY_OWNER_STATION Then
                    If imSortBy = SORTBY_MKTNAME Then
                        llValue = rst_Agreement!shttMktCode
                    Else            'owner by vehicle or station
                        llValue = rst_Agreement!shttOwnerArttCode
                    End If
                    blFound = gTestIncludeExcludeLong(llValue, ilInclChoiceCodes, llUseChoiceCodes())
                End If
            End If
            If blFound Then         'create the prepas record containing the agreement code to print
                If optRptDest(3).Value = True Then  'Mail List
                    'D.S. 09/15/04 did not need the extra ")", causes SQL error
                    'SQLQuery = SQLQuery + ")"
                    On Error GoTo ErrHand
                    'gUserActivityLog "E", sgReportListName & ": Prepass"
                    sMail = """" & Trim$(rst_Agreement!shttCallLetters) & """" & "," & """" & "1-" & Trim$(rst_Agreement!shttFax) & """"
                    Print #hmMail, sMail
                Else
                    
                    tmAmr.sOwner = ""
                    tmAmr.sVehicleName = ""
                    tmAmr.sMarket = ""
                    tmAmr.iRank = 0
                    tmAmr.sSalesRep = ""
                    tmAmr.sServRep = ""
                    
                    llVefInx = gBinarySearchVef(CLng(rst_Agreement!vefCode))
                    If llVefInx <> -1 Then
                        tmAmr.sVehicleName = Trim$(tgVehicleInfo(llVefInx).sVehicleName)
                    End If
                    
                    ilShttInx = gBinarySearchStationInfoByCode(rst_Agreement!shttCode)
                    ilMktInx = gBinarySearchMkt(CLng(tgStationInfoByCode(ilShttInx).iMktCode))
                    If ilMktInx <> -1 Then
                        tmAmr.sMarket = Trim$(tgMarketInfo(ilMktInx).sName)
                        tmAmr.iRank = tgMarketInfo(ilMktInx).iRank
                    End If
        
                    llOwnerInx = gBinarySearchOwner(CLng(tgStationInfoByCode(ilShttInx).lOwnerCode))
                    If llOwnerInx >= 0 Then
                        tmAmr.sOwner = Trim(tgOwnerInfo(llOwnerInx).sName)
                    End If
                    
                    'try the agreement mkt rep code first.  if does not exist, default to station mkt rep code
                    ilMktRepInx = gBinarySearchRepInfo(CLng(rst_Agreement!attMktRepUstCode), tgMarketRepInfo())
                    If ilMktRepInx <> -1 Then           'use the agreement mkt sales rep
                        tmAmr.sSalesRep = Trim$(tgMarketRepInfo(ilMktRepInx).sName)
                    Else
                        'agreement mkt rep doesnt exist, use station mkt rep
                        ilMktRepInx = gBinarySearchRepInfo(CLng(tgStationInfoByCode(ilShttInx).iMktRepUstCode), tgMarketRepInfo())
                        If ilMktRepInx <> -1 Then
                            tmAmr.sSalesRep = Trim$(tgMarketRepInfo(ilMktRepInx).sName)
                        End If
                    End If
                    
                    'try the agreement service rep code first.  if does not exist, default to station service rep code
                    ilServRepInx = gBinarySearchRepInfo(CLng(rst_Agreement!attServRepUstCode), tgServiceRepInfo())
                    If ilServRepInx <> -1 Then           'use the agreement mkt sales rep
                        tmAmr.sServRep = Trim$(tgServiceRepInfo(ilServRepInx).sName)
                    Else
                        'agreement mkt rep doesnt exist, use station mkt rep
                        ilServRepInx = gBinarySearchRepInfo(CLng(tgStationInfoByCode(ilShttInx).iServRepUstCode), tgServiceRepInfo())
                        If ilServRepInx <> -1 Then
                            tmAmr.sServRep = Trim$(tgServiceRepInfo(ilServRepInx).sName)
                        End If
                    End If
                    
                    tmAmr.lSmtCode = rst_Agreement!attCode
                    SQLQuery = "Insert Into amr ( "
                    SQLQuery = SQLQuery & "amrGenDate, "
                    SQLQuery = SQLQuery & "amrGenTime, "
                    SQLQuery = SQLQuery & "amrSmtCode, "
                    SQLQuery = SQLQuery & "amrRank, "
                    SQLQuery = SQLQuery & "amrMarket, "
                    SQLQuery = SQLQuery & "amrOwner, "
                    SQLQuery = SQLQuery & "amrVehicleName, "
                    SQLQuery = SQLQuery & "amrSalesRep, "
                    SQLQuery = SQLQuery & "amrServRep "
                    SQLQuery = SQLQuery & ") "
                    SQLQuery = SQLQuery & "Values ( "
                    SQLQuery = SQLQuery & "'" & Format$(sgGenDate, sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & ", "
                    SQLQuery = SQLQuery & tmAmr.lSmtCode & ", "
                    SQLQuery = SQLQuery & tmAmr.iRank & ", "
                    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sMarket)) & "', "
                    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sOwner)) & "', "
                    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sVehicleName)) & "', "
                    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sSalesRep)) & "', "
                    SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sServRep)) & "' "
                    SQLQuery = SQLQuery & ") "
                    On Error GoTo ErrHand
                    
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "VehAffRpt-cmdReport_Click"
                        Exit Sub
                    End If
                    On Error GoTo 0
                End If
            End If
            llCount = llCount + 1
            rst_Agreement.MoveNext
        Wend


        If optRptDest(3).Value = True Then
            gUserActivityLog "E", sgReportListName & ": Prepass"
            Close hmMail
            On Error GoTo 0
            Screen.MousePointer = vbDefault
            gMsgBox "Output Sent To: " & smToFile, vbInformation
            Exit Sub
        End If
        
        SQLQuery = "SELECT * from amr "
    
        SQLQuery = SQLQuery + " Inner Join att on amrsmtcode = attcode  INNER JOIN  shtt ON shttCode = attShfCode "
        'SQLQuery = SQLQuery + "LEFT OUTER JOIN mkt ON shttMktCode = mktCode "
        'SQLQuery = SQLQuery + " INNER JOIN   VEF_Vehicles ON attVefCode = vefCode "
                
        If optDatePhone(2).Value Then           'show export codes
            SQLQuery = SQLQuery + " Inner Join vef_Vehicles on attvefcode = vefcode Inner join vpf_Vehicle_Options on vefCode = vpfvefKCode "
        ElseIf optDatePhone(3).Value Then           'show FORMAT codes      Date: 8/13/2018 FYM
            SQLQuery = SQLQuery + " LEFT OUTER Join FMT_Station_Format on shttfmtcode = fmtcode "
        Else
            SQLQuery = SQLQuery + "LEFT OUTER JOIN ust ust ON shttMktRepUstCode = ust.ustCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN ust ustserv ON shttServRepUstCode = ustserv.ustCode "
        
            SQLQuery = SQLQuery + " LEFT OUTER JOIN CEF_Comments_Events CEF_Comments_Events ON ust.ustEMailCefCode = CEF_Comments_Events.cefCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN ust ustatt ON attMktRepUstCode = ustatt.ustCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN ust ustattserv ON attServRepUstCode = ustattserv.ustCode "
        
            SQLQuery = SQLQuery + "LEFT OUTER JOIN CEF_Comments_Events CEF_Comments_Events_Att ON ustatt.ustEMailCefCode = CEF_Comments_Events_Att.cefCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN CEF_Comments_Events CEF_Comments_Events_AttServ ON ustattserv.ustEMailCefCode = CEF_Comments_Events_AttServ.cefCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN CEF_Comments_Events CEF_Comments_Events_Serv ON ustserv.ustEMailCefCode = CEF_Comments_Events_Serv.cefCode "
        End If
        
        SQLQuery = SQLQuery & " Where (amrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' AND amrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    
        If imSortBy = SORTBY_VEHICLE Then
            slRptName = "afVhStVh.rpt"
            slExportName = "AgreeVehicleRpt"
        ElseIf imSortBy = SORTBY_OWNER_VEHICLE Then
            slRptName = "afAgreeOwnerVh.rpt"
            slExportName = "AgreeOwnerVhRpt"
        ElseIf imSortBy = SORTBY_STATION Then
            slRptName = "afVhStSt.rpt"
            slExportName = "AgreeStationRpt"
            
            'if sort by STATION, VEHICLE and show "FORMAT", use AfAgreeStnFmt.rpt
            'Date: 8/13/2018 FYM
            If optDatePhone(3).Value = True Then
                slRptName = "AfAgreeStnFmt.rpt"
                slExportName = "afAgreeStationFmt"
            End If
            
            If ckcStationInfo.Value Then
                sgCrystlFormula7 = "Y" 'ShowStationInfo
            Else
                sgCrystlFormula7 = "N" 'ShowStationInfo
            End If
        Else
            slRptName = "afAgreeOther.rpt"
            slExportName = "AgreeOtherRpt"
            If ckcStationInfo.Value Then
                sgCrystlFormula7 = "Y" 'ShowStationInfo
            Else
                sgCrystlFormula7 = "N" 'ShowStationInfo
            End If
            
        End If
    
        'A = audience & station, N = Mkt Nme & Station, R = Mkt Rank & Station, O = Owner & Station, W = Owner & Station, V = Vehicle, S = Station,
        If imSortBy = SORTBY_AUD Then
            sgCrystlFormula10 = "A"
        ElseIf imSortBy = SORTBY_MKTNAME Then
            sgCrystlFormula10 = "N"
        ElseIf imSortBy = SORTBY_MKTRANK Then
            sgCrystlFormula10 = "R"
        ElseIf imSortBy = SORTBY_OWNER_STATION Then
            sgCrystlFormula10 = "O"                 'owner/station
        ElseIf imSortBy = SORTBY_OWNER_VEHICLE Then
            sgCrystlFormula10 = "W"
        ElseIf imSortBy = SORTBY_STATION Then
            sgCrystlFormula10 = "S"
        ElseIf imSortBy = SORTBY_VEHICLE Then
            sgCrystlFormula10 = "V"
        End If
        
        If optDatePhone(0).Value Then
            sgCrystlFormula1 = "N" 'ShowPhone
        ElseIf optDatePhone(1).Value Then
            sgCrystlFormula1 = "Y" 'ShowPhone
        ElseIf optDatePhone(2).Value Then
            'added SHOW option: FORMAT; needed to make sure this code executes only for EXPORT CODES option
            'Date: 8/13/2018 FYM
            sgCrystlFormula1 = "E" 'Show Export Codes
            slRptName = "AfAgreeExpCodes.rpt"
            slExportName = "AfAgreeExpCodes"
        Else
            sgCrystlFormula1 = "N"      'set default value to "N"   Date: 8/14/208  FYM
        End If
    
    
        
        dFWeek = CDate(sStartDate)
        'StartDate
        sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
        dFWeek = CDate(sEndDate)
        'EndDate
        sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
        
        slDescription = gDateDescription(slEnterFrom, slEnterTo)
        sgCrystlFormula4 = Trim$(slDescription)         'entered date span
        slDescription = gDateDescription(slStartBetween1, slStartBetween2)
        sgCrystlFormula5 = Trim$(slDescription)         'agreements starting between 2 date spans
        slDescription = gDateDescription(slEndBetween1, slEndBetween2)
        sgCrystlFormula6 = Trim$(slDescription)         'agreements ending between 2 date spans
    
        
        
        gUserActivityLog "E", sgReportListName & ": Prepass"
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
    
        gUserActivityLog "S", sgReportListName & ": Clear amr"
    
        'remove all the records just printed
        SQLQuery = "DELETE FROM amr "
        SQLQuery = SQLQuery & " WHERE (amrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and amrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
        cnn.BeginTrans
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "VehAffRpt-cmdReport_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
            
        cmdReport.Enabled = True            'give user back control to gen, done buttons
        cmdDone.Enabled = True
        cmdReturn.Enabled = True
        
        gUserActivityLog "E", sgReportListName & ": Clear amr"
    
        Screen.MousePointer = vbDefault
        
        Exit Sub

    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmVehAffRpt-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmVehAffRpt
End Sub

'TTP 9943 - Add ability to import stations for report selectivity
Private Sub cmdStationListFile_Click()
    Dim slCurDir As String
    slCurDir = CurDir
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    
    ' Import from the Selected File
    gSelectiveStationsFromImport lbcStations, ckcAllStations, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
    cbcSort.AddItem "Audience (Desc),Station "
    cbcSort.AddItem "DMA Mkt Name,Station"
    cbcSort.AddItem "DMA Mkt Rank,Station"
    cbcSort.AddItem "Owner,Station"
    cbcSort.AddItem "Owner,Vehicle"
    cbcSort.AddItem "Station,Vehicle"
    cbcSort.AddItem "Vehicle,Station"
    cbcSort.ListIndex = 6               'default to Vehicle option
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmVehAffRpt
    gCenterForm frmVehAffRpt

    cmdReport.Enabled = False   'enable only after ALL filters are set  Date: 8/4/2018    FYM
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    frmVehAffRpt.Caption = "Affiliate Agreement Report - " & sgClientName
    imChkListBoxIgnore = False
    'SQLQuery = "SELECT vef.vefName from vef WHERE ((vef.vefvefCode = 0 AND vef.vefType = 'C') OR vef.vefType = 'L' OR vef.vefType = 'A')"
    'SQLQuery = SQLQuery + " ORDER BY vef.vefName"
    'Set rst = gSQLSelectCall(SQLQuery)
    'While Not rst.EOF
    '    grdVehAff.AddItem "" & rst(0).Value & ""
    '    rst.MoveNext
    'Wend
    slDate = Format$(gNow(), "m/d/yyyy")
    Do While Weekday(slDate, vbSunday) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop
    CalOnAirDate.Text = slDate
    CalOffAirDate.Text = DateAdd("d", 6, slDate)
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        ''grdVehAff.AddItem "" & Trim$(tgVehicleInfo(iLoop).sVehicle) & "|" & tgVehicleInfo(iLoop).iCode
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    
    chkListBox.Value = 0    'chged from false to 0 10-22-99
    
    CalEndBetween1.ZOrder (0)
    CalStartBetween1.ZOrder (0)
    CalEnterFrom.ZOrder (0)
    CalOnAirDate.ZOrder (0)
    CalEndBetween2.ZOrder (0)
    CalStartBetween2.ZOrder (0)
    CalEnterTo.ZOrder (0)
    CalOffAirDate.ZOrder (0)

    lbcStations.Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            'If tgStationInfo(iLoop).iType = 0 Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            'End If
        End If
    Next iLoop
    gPopExportTypes cboFileType     '3-15-04
    gPopRepInfo "M", tgMarketRepInfo()
    gPopRepInfo "S", tgServiceRepInfo()

    cboFileType.Enabled = False
    ckcStationInfo.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_Agreement.Close
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmVehAffRpt = Nothing
End Sub

Private Sub lbcSortSelection_Click()
    If imChkListOtherIgnore Then
        Exit Sub
    End If
    If ckcSortSelection.Value = vbChecked Then
        imChkListOtherIgnore = True
        'chkListBox.Value = False
        ckcSortSelection.Value = vbUnchecked    'chged from false to 0 10-22-99
        imChkListOtherIgnore = False
    End If

    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub lbcStations_Click()
    If imChkStationIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = vbChecked Then
        imChkStationIgnore = True
        ckcAllStations.Value = vbUnchecked    'chged from false to 0 10-22-99
        imChkStationIgnore = False
    End If

    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub lbcVehAff_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = vbChecked Then
        imChkListBoxIgnore = True
        chkListBox.Value = vbUnchecked    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If

    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub optDatePhone_Click(Index As Integer)
'    ckcStationInfo.Visible = True
'    chkPledgeOrPgm.Visible = True
   
    ckcContactInfo.Enabled = True
    ckcShowComments.Enabled = True
    ckcStationInfo.Enabled = True
    chkPledgeOrPgm.Enabled = True
    
    ' Enable Show Station and Contact Information if STATION and FORMAT are selected    Date: 8/8/2018  FYM
    mEnableDisableContactStationInfo
    If Index = 4 Then           '3-3-20 if showing Air Dates + pledge days, cannot show the program times.  instead the plege days will show along with the pledge days
        chkPledgeOrPgm.Enabled = False
        chkPledgeOrPgm.Value = vbUnchecked
        ckcStationInfo.Enabled = False
        ckcStationInfo.Value = vbUnchecked
    End If
    If Index = 0 Then               'show air dates
        If imSortBy = SORTBY_VEHICLE Or imSortBy = SORTBY_OWNER_VEHICLE Then    'anythning by major vehicle sort doesnt allow
            ckcStationInfo.Enabled = False
            ckcStationInfo.Value = vbUnchecked
            
        End If
    ElseIf Index = 1 Then           'show phone #
        If imSortBy = SORTBY_VEHICLE Or imSortBy = SORTBY_OWNER_VEHICLE Then    'anythning by major vehicle sort doesnt allow
            ckcStationInfo.Enabled = False
            ckcStationInfo.Value = vbUnchecked
            chkPledgeOrPgm.Enabled = False
            chkPledgeOrPgm.Value = vbUnchecked
        Else
            chkPledgeOrPgm.Enabled = False
            chkPledgeOrPgm.Value = vbUnchecked
        End If
    ElseIf Index = 2 Then           'show export codes
        ckcContactInfo.Enabled = False
        ckcContactInfo.Value = vbUnchecked
        ckcShowComments.Enabled = False
        ckcShowComments.Value = vbUnchecked
        ckcStationInfo.Enabled = False
        ckcStationInfo.Value = vbUnchecked
        chkPledgeOrPgm.Enabled = False
        chkPledgeOrPgm.Value = vbUnchecked
    ElseIf Index = 3 Then
        ' Enable Show Station and Contact Information if STATION and FORMAT are selected    Date: 8/8/2018  FYM
        mEnableDisableContactStationInfo
    End If
    
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub
