VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelCA 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avails Combo Report Selection"
   ClientHeight    =   5970
   ClientLeft      =   165
   ClientTop       =   1305
   ClientWidth     =   9945
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
   ScaleHeight     =   5970
   ScaleWidth      =   9945
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "rptselca.frx":0000
      Left            =   6000
      List            =   "rptselca.frx":0002
      TabIndex        =   66
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   20
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
      Left            =   7080
      TabIndex        =   62
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
      TabIndex        =   64
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
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4245
      Top             =   3615
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
         TabIndex        =   10
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
      TabIndex        =   11
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   16
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
         TabIndex        =   13
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
         TabIndex        =   15
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Avails Combo Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4440
      Left            =   75
      TabIndex        =   17
      Top             =   1440
      Width           =   9690
      Begin VB.ComboBox cbcDemo 
         BackColor       =   &H00FFFF00&
         Height          =   330
         ItemData        =   "rptselca.frx":0004
         Left            =   4200
         List            =   "rptselca.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   180
         Width           =   1115
         _ExtentX        =   1958
         _ExtentY        =   450
         Text            =   "9/19/2022"
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
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4095
         ScaleWidth      =   5310
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   5310
         Begin VB.CheckBox ckcIncludeAvailGroupSections 
            Caption         =   "Billboard, Drop-in, Extra sections"
            Height          =   225
            Left            =   0
            TabIndex        =   98
            Top             =   3960
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3435
         End
         Begin VB.OptionButton rbcGrossNet 
            Caption         =   "Net"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   1320
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   3720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.OptionButton rbcGrossNet 
            Caption         =   "Gross"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   3720
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   855
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalTo 
            Height          =   255
            Left            =   3480
            TabIndex        =   26
            Top             =   0
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   450
            Text            =   "9/19/2022"
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
         Begin VB.PictureBox plcMinor 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   90
            ScaleHeight     =   255
            ScaleWidth      =   4125
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   3570
            Visible         =   0   'False
            Width           =   4125
            Begin VB.OptionButton rbcMinor 
               Caption         =   "Daypart"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1890
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton rbcMinor 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   825
               TabIndex        =   81
               Top             =   0
               Width           =   960
            End
            Begin VB.OptionButton rbcMinor 
               Caption         =   "Week"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3045
               TabIndex        =   83
               TabStop         =   0   'False
               Top             =   0
               Width           =   1230
            End
         End
         Begin VB.PictureBox plcInterm 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   90
            ScaleHeight     =   255
            ScaleWidth      =   4155
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   3390
            Visible         =   0   'False
            Width           =   4155
            Begin VB.OptionButton rbcInterm 
               Caption         =   "Week"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3045
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   1230
            End
            Begin VB.OptionButton rbcInterm 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   825
               TabIndex        =   77
               Top             =   0
               Width           =   930
            End
            Begin VB.OptionButton rbcInterm 
               Caption         =   "Daypart"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1890
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   0
               Width           =   1020
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
            MaxLength       =   10
            TabIndex        =   25
            Top             =   0
            Width           =   1020
         End
         Begin VB.PictureBox plcCTypes 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   4455
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   720
            Width           =   4455
            Begin VB.CheckBox ckcCType 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1440
               TabIndex        =   28
               Top             =   -30
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2280
               TabIndex        =   29
               Top             =   -30
               Value           =   1  'Checked
               Width           =   8870
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Feed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3240
               TabIndex        =   75
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   480
               TabIndex        =   30
               Top             =   210
               Value           =   1  'Checked
               Width           =   1080
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   1560
               TabIndex        =   31
               Top             =   210
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Remnant"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   2760
               TabIndex        =   32
               Top             =   210
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   3840
               TabIndex        =   33
               Top             =   210
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Per Inquiry"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   480
               TabIndex        =   34
               Top             =   435
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   1695
               TabIndex        =   35
               Top             =   435
               Width           =   705
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   2400
               TabIndex        =   36
               Top             =   435
               Width           =   870
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Trades"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   3360
               TabIndex        =   37
               Top             =   435
               Value           =   1  'Checked
               Width           =   900
            End
         End
         Begin VB.PictureBox plcSpots 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   4380
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1440
            Width           =   4380
            Begin VB.CheckBox ckcSpots 
               Caption         =   "MG"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   3000
               TabIndex        =   84
               Top             =   425
               Value           =   1  'Checked
               Width           =   780
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Spinoff"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   1920
               TabIndex        =   48
               Top             =   435
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Recapturable"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   480
               TabIndex        =   47
               Top             =   435
               Value           =   1  'Checked
               Width           =   1440
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "N/C"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   3000
               TabIndex        =   46
               Top             =   210
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "-Fill"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   2280
               TabIndex        =   45
               Top             =   210
               Value           =   1  'Checked
               Width           =   645
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "+Fill"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   1440
               TabIndex        =   44
               Top             =   210
               Value           =   1  'Checked
               Width           =   705
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Bonus"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   480
               TabIndex        =   43
               Top             =   210
               Value           =   1  'Checked
               Width           =   870
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "ADU"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   3720
               TabIndex        =   42
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1080
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3000
               TabIndex        =   41
               Top             =   -30
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Charge"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2040
               TabIndex        =   40
               Top             =   -30
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1080
               TabIndex        =   39
               Top             =   -30
               Value           =   1  'Checked
               Width           =   930
            End
         End
         Begin VB.PictureBox plcTotals 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1980
            ScaleHeight     =   255
            ScaleWidth      =   2850
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   2820
            Visible         =   0   'False
            Width           =   2850
            Begin VB.OptionButton rbcTotals 
               Caption         =   "Day"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   825
               TabIndex        =   58
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton rbcTotals 
               Caption         =   "Week"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1470
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   0
               Width           =   885
            End
         End
         Begin VB.ComboBox cbcGroup 
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
            Left            =   1560
            TabIndex        =   52
            Top             =   2400
            Width           =   1500
         End
         Begin VB.PictureBox plcGameType 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            ScaleHeight     =   240
            ScaleWidth      =   5025
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   2640
            Width           =   5025
            Begin VB.CheckBox ckcGameType 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   840
               TabIndex        =   54
               Top             =   -30
               Width           =   1275
            End
            Begin VB.CheckBox ckcGameType 
               Caption         =   "Postponed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2160
               TabIndex        =   55
               Top             =   -30
               Width           =   1395
            End
         End
         Begin VB.CheckBox ckcShowNamedAvails 
            Caption         =   "Show Avail Names"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   49
            Top             =   2160
            Value           =   1  'Checked
            Width           =   2115
         End
         Begin VB.CheckBox ckcMultimedia 
            Caption         =   "Include Multimedia"
            Height          =   225
            Left            =   60
            TabIndex        =   56
            Top             =   2805
            Width           =   1995
         End
         Begin VB.PictureBox plcUnsoldOnly 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            ScaleHeight     =   240
            ScaleWidth      =   4665
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   3000
            Width           =   4665
            Begin VB.CheckBox ckcUnsoldOnly 
               Caption         =   "MM Only"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2565
               TabIndex        =   63
               Top             =   0
               Width           =   1395
            End
            Begin VB.CheckBox ckcUnsoldOnly 
               Caption         =   "Avails Only"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1185
               TabIndex        =   74
               Top             =   0
               Width           =   1275
            End
         End
         Begin VB.CheckBox ckcNewPage 
            Caption         =   "New page each group"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1965
            TabIndex        =   50
            Top             =   2100
            Visible         =   0   'False
            Width           =   2250
         End
         Begin VB.PictureBox plcMajor 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   105
            ScaleHeight     =   255
            ScaleWidth      =   4305
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   3150
            Visible         =   0   'False
            Width           =   4305
            Begin VB.OptionButton rbcMajor 
               Caption         =   "Daypart"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1890
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton rbcMajor 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   825
               TabIndex        =   71
               Top             =   0
               Width           =   1050
            End
            Begin VB.OptionButton rbcMajor 
               Caption         =   "Week"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3045
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   0
               Width           =   1230
            End
         End
         Begin VB.PictureBox plcLengths 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   4455
            TabIndex        =   91
            Top             =   360
            Width           =   4455
            Begin VB.TextBox edcLength 
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
               Index           =   4
               Left            =   3720
               MaxLength       =   3
               TabIndex        =   96
               Top             =   0
               Width           =   435
            End
            Begin VB.TextBox edcLength 
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
               Index           =   3
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   95
               Top             =   0
               Width           =   435
            End
            Begin VB.TextBox edcLength 
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
               Index           =   2
               Left            =   2760
               MaxLength       =   3
               TabIndex        =   94
               Top             =   0
               Width           =   435
            End
            Begin VB.TextBox edcLength 
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
               Index           =   1
               Left            =   2280
               MaxLength       =   3
               TabIndex        =   93
               Text            =   "10"
               Top             =   0
               Width           =   435
            End
            Begin VB.TextBox edcLength 
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
               Index           =   0
               Left            =   1800
               MaxLength       =   3
               TabIndex        =   92
               Text            =   "30"
               Top             =   0
               Width           =   435
            End
            Begin VB.Label lacLength 
               Appearance      =   0  'Flat
               Caption         =   "Lengths to Highlight"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   0
               TabIndex        =   97
               Top             =   30
               Width           =   1815
            End
         End
         Begin VB.Label lacDemo 
            Caption         =   "Demo"
            Height          =   255
            Left            =   3480
            TabIndex        =   100
            Top             =   3840
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lacByGrossNet 
            Appearance      =   0  'Flat
            Caption         =   "By"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            TabIndex        =   87
            Top             =   3720
            Visible         =   0   'False
            Width           =   200
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "End"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3000
            TabIndex        =   24
            Top             =   60
            Width           =   405
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Event Dates-Start"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   60
            Width           =   1560
         End
         Begin VB.Label lacGroup 
            Appearance      =   0  'Flat
            Caption         =   "Vehicle Groups"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   51
            Top             =   2430
            Width           =   1335
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
         Height          =   4170
         Left            =   5280
         ScaleHeight     =   4170
         ScaleWidth      =   4335
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   165
         Width           =   4335
         Begin VB.CheckBox ckcAllVGItems 
            Caption         =   "All Group Items"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2160
            TabIndex        =   86
            Top             =   2100
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   3
            ItemData        =   "rptselca.frx":0008
            Left            =   2160
            List            =   "rptselca.frx":000F
            MultiSelect     =   2  'Extended
            TabIndex        =   85
            Top             =   2400
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.CheckBox ckcAllNamedAvails 
            Caption         =   "All Named Avails"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   68
            Top             =   2100
            Width           =   1875
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   2
            ItemData        =   "rptselca.frx":0016
            Left            =   300
            List            =   "rptselca.frx":0018
            MultiSelect     =   2  'Extended
            TabIndex        =   67
            Top             =   2400
            Width           =   3885
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   1
            ItemData        =   "rptselca.frx":001A
            Left            =   300
            List            =   "rptselca.frx":0021
            TabIndex        =   9
            Top             =   2400
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1710
            Index           =   0
            ItemData        =   "rptselca.frx":0028
            Left            =   300
            List            =   "rptselca.frx":002A
            MultiSelect     =   2  'Extended
            TabIndex        =   8
            Top             =   240
            Width           =   3885
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   0
            Width           =   1635
         End
         Begin VB.Label lacRC 
            Appearance      =   0  'Flat
            Caption         =   "Rate Card"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2880
            TabIndex        =   65
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   21
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   19
      Top             =   105
      Width           =   2805
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   75
      TabIndex        =   3
      Top             =   60
      Width           =   1890
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Export"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   960
         Width           =   1395
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Save to File"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   1395
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Print"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   750
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Label lblExportStatus 
      Caption         =   "lblExportStatus"
      Height          =   255
      Left            =   2160
      TabIndex        =   101
      Top             =   120
      Width           =   3855
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
Attribute VB_Name = "RptSelCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptselca.frm on Wed 6/17/09 @ 12:56 P
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
' File Name: RptSelCA.Frm - Avails Combo Report for Sports and non-Sports (by day/week)
'
'
' Release: 4.7  9/25/00
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
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imSetAllnamedAvails As Integer
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imAllclickedNamedAvails As Integer
Dim imSetAllVGItems As Integer
Dim imAllClickedVGItems As Integer
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
'Import contract report
'Spot week Dump
Dim imTerminate As Integer
'Dim tmSRec As LPOPREC
'Rate Card
Dim smRateCardTag As String

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

Private Sub cbcGroup_Click()
    Dim illoop As Integer
    Dim ilSetIndex As Integer
    Dim ilRet As Integer

    illoop = cbcGroup.ListIndex
    ilSetIndex = gFindVehGroupInx(illoop, tgVehicleSets1())
    If ilSetIndex > 0 Then
        smVehGp5CodeTag = ""
        ilRet = gPopMnfPlusFieldsBox(RptSelCA, lbcSelection(3), tgSOCode(), smVehGp5CodeTag, "H" & Trim$(str$(ilSetIndex)))
    Else
        lbcSelection(3).Clear
        ckcAllVGItems.Value = vbUnchecked
    End If
    mSetCommands
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
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcAllNamedAvails_Click()
  'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllNamedAvails.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    ilValue = Value
    If imSetAllnamedAvails Then
        imAllclickedNamedAvails = True
        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllclickedNamedAvails = False
    End If
    mSetCommands
End Sub

Private Sub ckcAllVGItems_Click()
    Dim Value As Integer
    Value = False
    If ckcAllVGItems.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    ilValue = Value
    If imSetAllVGItems Then
        imAllClickedVGItems = True
        llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClickedVGItems = False
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
    Dim slReportName As String
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

    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    ilListIndex = lbcRptType.ListIndex

    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        slReportName = "AvailsCombo.Rpt"
        If ilListIndex = AVAILSCOMBO_NONSPORTS Then     'avails combo by day or week (non-sports)  not coded yet
            slReportName = "AvailsComboNS.Rpt"
        End If
        If rbcOutput(3).Value = False Then 'Not using the Special Export option in TTP 10434 - Event and Sports export (WWO)
            If Not gOpenPrtJob(slReportName) Then
                igGenRpt = False
                frcOutput.Enabled = igOutput
                frcCopies.Enabled = igCopies
                'frcWhen.Enabled = igWhen
                frcFile.Enabled = igFile
                frcOption.Enabled = igOption
                'frcRptType.Enabled = igReportType
                Exit Sub
            End If
        End If
        If rbcOutput(3).Value = False Then 'Not using the Special Export option in TTP 10434 - Event and Sports export (WWO)
            ilRet = gCmcGenCA()
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
        End If
        Screen.MousePointer = vbHourglass

        gCreateComboAvails
        Screen.MousePointer = vbDefault
        If rbcOutput(0).Value Then 'Display
            DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
            igDestination = 0
            Report.Show vbModal
        ElseIf rbcOutput(1).Value Then 'Print
            ilCopies = Val(edcCopies.Text)
            ilRet = gOutputToPrinter(ilCopies)
        ElseIf rbcOutput(2).Value Then 'Save to File
            slFileName = edcFileName.Text
            'ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
        End If
    Next ilJobs
    imGenShiftKey = 0
    'Export - TTP 10434 - Event and Sports export (WWO)
    If Not rbcOutput(3).Value Then
        Screen.MousePointer = vbHourglass
        gCRGrfClear
    End If
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

Private Sub CSI_CalFrom_CalendarChanged()
    Dim slDate As String
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilLen As Integer
    ilLen = Len(CSI_CalFrom.Text)
    If ilLen >= 4 Then
        slDate = CSI_CalFrom.Text           'retrieve jan thru dec year
        slDate = gObtainStartStd(slDate)
        llDate = gDateValue(slDate)

        'populate Rate Cards and bring in Rcf, Rif, and Rdf
        ilRet = gPopRateCardBox(RptSelCA, llDate, RptSelCA!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
    End If
    mSetCommands
End Sub

Private Sub CSI_CalFrom_Change()
    Dim slDate As String
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilLen As Integer
    ilLen = Len(CSI_CalFrom.Text)
    If ilLen >= 4 Then
        slDate = CSI_CalFrom.Text           'retrieve jan thru dec year
        slDate = gObtainStartStd(slDate)
        llDate = gDateValue(slDate)

        'populate Rate Cards and bring in Rcf, Rif, and Rdf
        ilRet = gPopRateCardBox(RptSelCA, llDate, RptSelCA!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
    End If
    mSetCommands
End Sub

Private Sub CSI_CalFrom_GotFocus()
    gCtrlGotFocus CSI_CalFrom
End Sub

Private Sub CSI_CalTo_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalTo_GotFocus()
    gCtrlGotFocus CSI_CalTo
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
    If ((KeyAscii <> KEYBACKSPACE) And (KeyAscii <= 32)) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KeyDown) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcLength_GotFocus(Index As Integer)
    gCtrlGotFocus edcLength(Index)

End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

'Private Sub edcSelCFrom_Change()
'Dim slDate As String
'Dim llDate As Long
'Dim ilRet As Integer
'Dim ilLen As Integer
'    ilLen = Len(edcSelCFrom)
'    If ilLen >= 4 Then
'        slDate = edcSelCFrom           'retrieve jan thru dec year
'        slDate = gObtainStartStd(slDate)
'        llDate = gDateValue(slDate)
'
'        'populate Rate Cards and bring in Rcf, Rif, and Rdf
'        ilRet = gPopRateCardBox(RptSelCA, llDate, RptSelCA!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
'    End If
'    mSetCommands
'End Sub
'Private Sub edcSelCFrom_GotFocus()
'    gCtrlGotFocus edcSelCFrom
'End Sub

Private Sub edcSelCFrom1_Change()
    mSetCommands
End Sub

Private Sub edcSelCFrom1_GotFocus()
    gCtrlGotFocus edcSelCFrom1
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
    Me.Refresh
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
    lblExportStatus.Caption = ""
    'RptSelCA.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase imCodes
    PECloseEngine
    
    Set RptSelCA = Nothing   'Remove data segment
    
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcRptType_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRet                         llRg                          ilValue                   *
'*                                                                                        *
'******************************************************************************************

    Dim ilListIndex As Integer
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim ilLoopOnListBox As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilAnfCode As Integer
    Dim slName As String

    ilListIndex = lbcRptType.ListIndex
    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    If ilListIndex = AVAILSCOMBO_SPORTS Then
        'TTP 10434 - Event and Sports export (WWO)
        rbcOutput(3).Visible = True
        edcSelCFrom1.Visible = False
        CSI_CalTo.Visible = True
        mSellConvVirtVehPop 0, True            'get sports vehicles
'        edcSelCFrom1.MaxLength = 10
        ilRet = gAvailsPop(RptSelCA, lbcSelection(2), tgNamedAvail())       'show the named avails for selectivity
        'default the selection on if user named avails flag is set
        For illoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1
            If tgAvailAnf(illoop).sRptDefault = "Y" Then                'set as selected
                For ilLoopOnListBox = 0 To lbcSelection(2).ListCount
                    slNameCode = tgNamedAvail(ilLoopOnListBox).sKey
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slName, 3, "|", slName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilAnfCode = Val(slCode)
                    If ilAnfCode = tgAvailAnf(illoop).iCode Then        'the named avail is set to show on the Combo aVails, default selected
                        lbcSelection(2).Selected(ilLoopOnListBox) = True
                        Exit For
                    End If
                Next ilLoopOnListBox
            End If
        Next illoop
        ckcShowNamedAvails.Top = plcSpots.Top + plcSpots.Height
        lacGroup.Caption = "Group"
        lacGroup.Move ckcShowNamedAvails.Left + ckcShowNamedAvails.Width + 240, ckcShowNamedAvails.Top + 15, 720
        cbcGroup.Move lacGroup.Left + lacGroup.Width - 120, lacGroup.Top - 30

        plcGameType.Move 120, cbcGroup.Top + cbcGroup.Height + 30
        ckcMultimedia.Move 120, plcGameType.Top + plcGameType.Height
        plcUnsoldOnly.Move 120, ckcMultimedia.Top + ckcMultimedia.Height
        lbcSelection(3).Visible = False
        ckcAllVGItems.Visible = False
        ckcAllNamedAvails.Value = vbChecked
    Else                        'Avails Combo by Day/Week
        'TTP 10434 - Event and Sports export (WWO)
        rbcOutput(3).Visible = False
        CSI_CalTo.Visible = False
        edcSelCFrom1.Visible = True
        mSellConvVirtVehPop 0, False            'dont get sports vehicles
        lacSelCFrom.Caption = "Dates - Start"
'        edcSelCFrom.Left = 1320
        CSI_CalFrom.Left = 1320
        lacSelCFrom1.Caption = "# Weeks"
'        lacSelCFrom1.Left = edcSelCFrom.Left + edcSelCFrom.Width + 240
        lacSelCFrom1.Left = CSI_CalFrom.Left + CSI_CalFrom.Width + 240
        lacSelCFrom1.Width = 840
'        edcSelCFrom1.Left = lacSelCFrom1.Left + lacSelCFrom1.Width
        CSI_CalTo.Left = lacSelCFrom1.Left + lacSelCFrom1.Width
        edcSelCFrom1.Width = 480
        edcSelCFrom1.MaxLength = 2
        plcUnsoldOnly.Visible = False
        ckcMultimedia.Visible = False
        ckcShowNamedAvails.Visible = False
        ckcNewPage.Move 120, plcSpots.Top + plcSpots.Height
        ckcNewPage.Visible = True
        plcGameType.Visible = False

        lacGroup.Caption = "Group"
        lacGroup.Move ckcNewPage.Left + ckcNewPage.Width + 240, ckcNewPage.Top + 15, 720
        cbcGroup.Move lacGroup.Left + lacGroup.Width - 120, lacGroup.Top - 30

        plcTotals.Move 120, cbcGroup.Top + cbcGroup.Height
        plcTotals.Visible = True
        rbcTotals(0).Visible = True
        rbcTotals(1).Visible = True
        rbcTotals(0).Value = True
        If rbcTotals(0).Value Then
            rbcTotals_Click 0
        Else
            rbcTotals(0).Value = True
        End If

        plcMajor.Move plcTotals.Left, plcTotals.Top + plcTotals.Height
        If rbcMajor(0).Value Then              'default major to vehicle
            rbcMajor_Click 0
        Else
            rbcMajor(0).Value = True
        End If
        plcInterm.Move plcMajor.Left, plcMajor.Top + plcMajor.Height
        If rbcInterm(1).Value Then             'default intermediate to daypart
            rbcInterm_Click 2
        Else
            rbcInterm(1).Value = True
        End If
        plcMinor.Move plcInterm.Left, plcInterm.Top + plcInterm.Height
        If rbcMinor(2).Value Then             'default minor to day/week card
            rbcMinor_Click 2
        Else
            rbcMinor(2).Value = True
        End If
        plcMajor.Visible = True
        plcInterm.Visible = True
        plcMinor.Visible = True

        lbcSelection(1).Visible = True
        lbcSelection(2).Visible = False
        ckcAllNamedAvails.Visible = False
        lacRC.Move ckcAllNamedAvails.Left, ckcAllNamedAvails.Top
        lacRC.Visible = True
        ckcAllVGItems.Visible = True
        lbcSelection(3).Visible = True
        
        If (tgSpf.sAvailEqualize = "3" Or tgSpf.sAvailEqualize = "6") And ilListIndex <> AVAILSCOMBO_SPORTS Then
            edcLength(4).Visible = False
        End If

        'Date: 1/10/2020 added Gross/Net option for schedule and rate card display
        lacByGrossNet.Visible = True: rbcGrossNet(0).Visible = True: rbcGrossNet(1).Visible = True
        
        lacByGrossNet.Move plcInterm.Left, plcInterm.Top + plcInterm.Height + 300
        rbcGrossNet(0).Move lacByGrossNet.Left + lacByGrossNet.Width + 250, lacByGrossNet.Top
        rbcGrossNet(1).Move rbcGrossNet(0).Left + rbcGrossNet(0).Width + 10, rbcGrossNet(0).Top
        
    End If

End Sub

Private Sub lbcSelection_Click(Index As Integer)
    If Not imAllClicked Then
        If Index = 0 Then           'vehicle list box
            imSetAll = False
            ckcAll.Value = vbUnchecked  '9-12-02 False
            imSetAll = True
        ElseIf Index = 2 Then           'named avails
            imSetAllnamedAvails = False
            ckcAllNamedAvails.Value = vbUnchecked  '9-12-02 False
            imSetAllnamedAvails = True
        ElseIf Index = 3 Then               'all vehicle group items
            imSetAllVGItems = False
            ckcAllVGItems.Value = vbUnchecked  '9-12-02 False
            imSetAllVGItems = True
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
Dim illoop As Integer
Dim slStr As String
Dim ilListIndex As Integer
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

    RptSelCA.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    illoop = InStr(slStr, "&")
    If illoop > 0 Then
        slStr = Left$(slStr, illoop - 1) & "&&" & Mid$(slStr, illoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    imSetAllVGItems = True
    imAllClickedVGItems = False
'    ckcAll.Move 120, 0
    imAllclickedNamedAvails = False
    imSetAllnamedAvails = True
    'lbcSelection(0).Move 120, ckcAll.Height + 30, 4200, 1500    '1650
    
    'lacRC.Move ckcAll.Left, lbcSelection(0).Top + lbcSelection(0).Height + 30
    'ckcAllNamedAvails.Move ckcAll.Left, lbcSelection(0).Top + lbcSelection(0).Height + 30
    'lbcSelection(1).Move lbcSelection(0).Left, lacRC.Top + lacRC.Height + 30, 4200, 1500
    'lbcSelection(2).Move lbcSelection(0).Left, ckcAllNamedAvails.Top + ckcAllNamedAvails.Height + 30, 4200, 1500
    lbcSelection(0).Visible = True
    'lbcSelection(1).Visible = True
        
    'Demo's for TTP 10434 - Event and Sports export (WWO)
    ilRet = gPopMnfPlusFieldsBox(RptSelCA, cbcDemo, tgDemoCode(), sgDemoCodeTag, "D")
    If cbcDemo.ListCount > 0 And cbcDemo.ListIndex < 0 Then
        If cbcDemo.ListIndex < 0 Then cbcDemo.ListIndex = 0
    End If
    gCenterStdAlone RptSelCA
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
    pbcSelC.Visible = False


    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True
    lacSelCFrom.Visible = True
'    edcSelCFrom.Visible = True
    ckcAll.Visible = True
    pbcOption.Visible = True
    pbcOption.Enabled = True
    gPopVehicleGroups RptSelCA!cbcGroup, tgVehicleSets1(), True

    ilRet = gObtainSAF()
    edcLength(0) = tgSaf(0).iRptLenDefault(0)
    edcLength(1) = tgSaf(0).iRptLenDefault(1)
    edcLength(2) = tgSaf(0).iRptLenDefault(2)
    edcLength(3) = tgSaf(0).iRptLenDefault(3)
    edcLength(4) = tgSaf(0).iRptLenDefault(4)

    lbcRptType.AddItem "Event & Sports Avails", AVAILSCOMBO_SPORTS
    lbcRptType.AddItem "Avails Combo by Day/Week", AVAILSCOMBO_NONSPORTS
    If lbcRptType.ListCount > 0 Then
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
    'gInitStdAlone RptSelCA, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Avails Combo Report"
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
'**********************************************************
'*                                                        *
'*      Procedure Name:mSellConvVirtVehPop                *
'*                                                        *
'*             Created:6/16/93       By:D. LeVine         *
'*            Modified:              By:                  *
'*                                                        *
'*            Comments: Populate the selection combo      *
'*                      box                               *
'*      <input>  ilIndex = index to list box array that   *
'*               contains vehicle selection               *
'                ilObtainSports - get sports only,else    *
'*               all other sellingtype vehicles (non-sports) *
'*******************************************************
Private Sub mSellConvVirtVehPop(ilIndex As Integer, ilObtainSports As Integer)
    Dim ilRet As Integer

    If ilObtainSports Then
        ilRet = gPopUserVehicleBox(RptSelCA, VEHSPORT + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)  'lbcCSVNameCode)
    Else
        ilRet = gPopUserVehicleBox(RptSelCA, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHEXCLUDESPORT + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)   'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelCA
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVirtVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim illoop As Integer
    Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    ilEnable = False
'    If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") Then
    If (CSI_CalFrom.Text <> "") Then                '9-4-19 And (edcSelCFrom1.Text <> "") Then
        ilEnable = True
        'atleast one budget must be selected
        If ilEnable Then
            If ilListIndex = AVAILSCOMBO_NONSPORTS Then
                ilEnable = False
                If CSI_CalFrom.Text <> "" Then      'Date: 1/15/2020 added CSI calendar control for date entries --> edcSelCFrom1.Text <> "" Then
                    For illoop = 0 To lbcSelection(1).ListCount - 1 Step 1      'budget entry must be selected
                        If lbcSelection(1).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
            Else
                'event (game)
                If (CSI_CalTo.Text = "") Then
                    ilEnable = False
                End If
            End If
            If ilEnable Then
                ilEnable = False
                'Check vehicle selection
                For illoop = 0 To lbcSelection(0).ListCount - 1 Step 1      'vehicle entry must be selected
                    If lbcSelection(0).Selected(illoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next illoop
            End If
        End If
    End If

    If ilEnable Then
        If ilListIndex = AVAILSCOMBO_NONSPORTS Then
            'if vehicle group selected, see if items selected
            If cbcGroup.ListIndex > 0 Then
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
            
        ElseIf rbcOutput(2).Value Then     'Save As
            If (imFTSelectedIndex >= 0) And (edcFileName.Text <> "") Then
                ilEnable = True
            Else
                ilEnable = False
            End If
            
        ElseIf rbcOutput(3).Value Then 'Export - TTP 10434 - Event and Sports export (WWO)
            ilEnable = True
            
        End If
    End If
    cmcGen.Enabled = ilEnable
End Sub

'Export - TTP 10434 - Event and Sports export (WWO)
Private Sub SetCommandsForExport()
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex

    'Setup UI for Report
    If ilListIndex = AVAILSCOMBO_SPORTS Then
        If rbcOutput(0).Value Or rbcOutput(1).Value Or rbcOutput(2).Value Then   'Display, Print,Save to File
            plcLengths.Visible = True
            plcLengths.Top = 360
            plcCTypes.Top = 720 'Contract Types
            RptSelCA.plcSpots.Top = 1440 'Spot Types
            ckcShowNamedAvails.Top = 2175 'Show Avail Names
            lacGroup.Top = 2190
            cbcGroup.Top = 2160
            ckcShowNamedAvails.Visible = True
            lacGroup.Visible = True
            cbcGroup.Visible = True
            ckcMultimedia.Top = 2745
            ckcMultimedia.Visible = True
            plcUnsoldOnly.Visible = True
            plcUnsoldOnly.Top = 2970
            plcGameType.Top = 2505 'Include
            plcGameType.Visible = True 'Include
            ckcIncludeAvailGroupSections.Visible = False 'Billboard, Drop-in, Extra sections
            cbcDemo.Visible = False
            lacDemo.Visible = False
            lbcSelection(0).Height = 1710 'Vechicles:
            ckcAllNamedAvails.Visible = True 'Named Avails
            lbcSelection(2).Visible = True 'Named Avails
            cmcGen.Caption = "Generate Report"
            
        ElseIf rbcOutput(3).Value Then 'Export - TTP 10434 - Event and Sports export (WWO)
            plcLengths.Top = 360
            plcLengths.Visible = False
            plcCTypes.Top = 420 'Contract Types
            RptSelCA.plcSpots.Top = 1300 'Spot Types
            ckcShowNamedAvails.Visible = False
            lacGroup.Visible = False
            cbcGroup.Visible = False
            ckcShowNamedAvails.Top = 1855 'Show Avail Names
            lacGroup.Top = 1855
            cbcGroup.Top = 1855
            ckcMultimedia.Top = 2745
            ckcMultimedia.Visible = False
            plcUnsoldOnly.Visible = False
            plcUnsoldOnly.Top = 2970
            plcGameType.Top = 2225 'Include
            plcGameType.Visible = True 'Include
            ckcIncludeAvailGroupSections.Top = 2440 'Billboard, Drop-in, Extra sections
            ckcIncludeAvailGroupSections.Left = 960 'Billboard, Drop-in, Extra sections
            ckcIncludeAvailGroupSections.Visible = True 'Billboard, Drop-in, Extra sections
            cbcDemo.Top = 3060
            cbcDemo.Left = 1080
            cbcDemo.Visible = True
            lacDemo.Left = 100
            lacDemo.Top = 2880
            lacDemo.Visible = True
            lbcSelection(0).Height = 4000 'Vechicles:
            ckcAllNamedAvails.Visible = False 'Named Avails
            lbcSelection(2).Visible = False 'Named Avails
            cmcGen.Caption = "Export to Excel"
            
        End If
    End If


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
    Unload RptSelCA
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcGameType_Paint()
    plcGameType.CurrentX = 0
    plcGameType.CurrentY = 0
    plcGameType.Print "Include"
End Sub

Private Sub plcInterm_Paint()
    plcInterm.CurrentX = 0
    plcInterm.CurrentY = 0
    plcInterm.Print "Interm"
End Sub
Private Sub plcMajor_Paint()
    plcMajor.CurrentX = 0
    plcMajor.CurrentY = 0
    plcMajor.Print "Major"
End Sub

Private Sub plcMinor_Paint()
    plcMinor.CurrentX = 0
    plcMinor.CurrentY = 0
    plcMinor.Print "Minor"
End Sub

Private Sub plcUnsoldOnly_Paint()
    plcUnsoldOnly.CurrentX = 0
    plcUnsoldOnly.CurrentY = 0
    plcUnsoldOnly.Print "Show Unsold"
End Sub

Private Sub rbcInterm_Click(Index As Integer)
    Index = Index
End Sub

Private Sub rbcMajor_Click(Index As Integer)
    Index = Index
End Sub

Private Sub rbcMinor_Click(Index As Integer)
    Index = Index
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
                lbcRptType_Click
            Case 1  'Print
                frcFile.Visible = False
                frcFile.Enabled = False
                frcCopies.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcCopies.Visible = True
                lbcRptType_Click
            Case 2  'File
                frcCopies.Visible = False
                frcCopies.Enabled = False
                frcFile.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcFile.Visible = True
                lbcRptType_Click
            Case 3  'Export - TTP 10434 - Event and Sports export (WWO)
                frcFile.Enabled = False
                frcCopies.Visible = False   'Print Box
                frcFile.Visible = False     'Save to File Box
                frcCopies.Enabled = False
        End Select
    End If
    mSetCommands
    SetCommandsForExport
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

Private Sub rbcTotals_Click(Index As Integer)
    If Index = 0 Then
        rbcMajor(2).Caption = "Day"
        rbcInterm(2).Caption = "Day"
        rbcMinor(2).Caption = "Day"
    Else
        rbcMajor(2).Caption = "Week"
        rbcInterm(2).Caption = "Week"
        rbcMinor(2).Caption = "Week"
    End If
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcTotals_Paint()
    plcTotals.CurrentX = 0
    plcTotals.CurrentY = 0
    plcTotals.Print "Totals by"
End Sub
Private Sub plcSpots_Paint()
    plcSpots.CurrentX = 0
    plcSpots.CurrentY = 0
    plcSpots.Print "Spot Types"
End Sub
Private Sub plcCTypes_Paint()
    plcCTypes.CurrentX = 0
    plcCTypes.CurrentY = 0
    plcCTypes.Print "Contract Types"
End Sub
