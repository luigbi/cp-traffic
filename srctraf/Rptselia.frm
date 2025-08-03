VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelIA 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5610
   ClientLeft      =   420
   ClientTop       =   4485
   ClientWidth     =   9705
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
   ScaleHeight     =   5610
   ScaleWidth      =   9705
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   55
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8145
      TabIndex        =   16
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
      TabIndex        =   50
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
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   -60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcNoSort 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2205
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   5115
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcMultiCntr 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2055
      Sorted          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4830
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.FileListBox lbcFileName 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3060
      Pattern         =   "*.Dal"
      TabIndex        =   45
      Top             =   4935
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox lbcLnCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5070
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcCntrCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2130
      Sorted          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5055
      Visible         =   0   'False
      Width           =   960
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4005
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
         ForeColor       =   &H00FFFF00&
         Height          =   300
         Left            =   1065
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   15
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   9
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
      TabIndex        =   17
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   24
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
         TabIndex        =   21
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
         TabIndex        =   23
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Vehicle Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4005
      Left            =   45
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
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
         Height          =   3600
         Left            =   60
         ScaleHeight     =   3600
         ScaleMode       =   0  'User
         ScaleWidth      =   4875
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   255
         Width           =   4875
         Begin V81TrafficReports.CSI_Calendar CSI_CalTo 
            Height          =   255
            Left            =   3000
            TabIndex        =   74
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            Text            =   "9/11/19"
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BorderStyle     =   1
            CSI_ShowDropDownOnFocus=   0   'False
            CSI_InputBoxBoxAlignment=   1
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
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   1200
            TabIndex        =   73
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            Text            =   "9/11/19"
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BorderStyle     =   1
            CSI_ShowDropDownOnFocus=   0   'False
            CSI_InputBoxBoxAlignment=   1
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
         Begin VB.CheckBox ckcShowScript 
            Caption         =   "Show Script "
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2280
            TabIndex        =   72
            Top             =   1585
            Width           =   2055
         End
         Begin VB.CheckBox ckcUseCountAff 
            Caption         =   "Use Affidavit with Times+Counts"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   71
            Top             =   3360
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.CheckBox ckcInclCreativeTitle 
            Caption         =   "Show Creative Title"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   70
            Top             =   1585
            Width           =   2115
         End
         Begin VB.CheckBox ckcShowRate 
            Caption         =   "Show Rates"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   68
            Top             =   1335
            Width           =   1515
         End
         Begin VB.PictureBox plcSortBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   4140
            TabIndex        =   65
            Top             =   600
            Width           =   4140
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Prod/ISCI"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1920
               TabIndex        =   67
               Top             =   0
               Width           =   1365
            End
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   840
               TabIndex        =   66
               Top             =   0
               Value           =   -1  'True
               Width           =   1155
            End
         End
         Begin VB.PictureBox plcShowInvNo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2280
            ScaleHeight     =   240
            ScaleWidth      =   1980
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1335
            Visible         =   0   'False
            Width           =   1980
            Begin VB.CheckBox ckcShowInvNo 
               Caption         =   "Show Invoice #"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   0
               TabIndex        =   62
               Top             =   0
               Value           =   1  'Checked
               Width           =   1680
            End
         End
         Begin VB.PictureBox plcContrFeed 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   120
            ScaleHeight     =   480
            ScaleWidth      =   4380
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   2865
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcContrFeed 
               Caption         =   "Promo Spots"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2160
               TabIndex        =   64
               Top             =   240
               Width           =   1440
            End
            Begin VB.CheckBox ckcContrFeed 
               Caption         =   "PSA Spots"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   840
               TabIndex        =   63
               Top             =   240
               Width           =   1200
            End
            Begin VB.CheckBox ckcContrFeed 
               Caption         =   "Contract Spots"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   840
               TabIndex        =   60
               Top             =   0
               Value           =   1  'Checked
               Width           =   1665
            End
            Begin VB.CheckBox ckcContrFeed 
               Caption         =   "Feed Spots"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2520
               TabIndex        =   59
               Top             =   0
               Value           =   1  'Checked
               Width           =   1440
            End
         End
         Begin VB.PictureBox plcSelC10 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4260
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1080
            Visible         =   0   'False
            Width           =   4260
            Begin VB.CheckBox ckcSelC10 
               Caption         =   "Skip to new page each vehicle"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   3600
            End
         End
         Begin VB.PictureBox plcSelC7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   4380
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   4380
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Both"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3000
               TabIndex        =   13
               Top             =   0
               Width           =   1185
            End
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Summary"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1680
               TabIndex        =   12
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Detail"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   720
               TabIndex        =   11
               Top             =   0
               Width           =   1125
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
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   2565
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Regional"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2730
               TabIndex        =   69
               Top             =   0
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Local"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1770
               TabIndex        =   34
               Top             =   0
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "National"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   570
               TabIndex        =   33
               Top             =   0
               Visible         =   0   'False
               Width           =   1200
            End
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
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   27
            Top             =   2205
            Width           =   1170
         End
         Begin VB.PictureBox plcSelC1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4140
            TabIndex        =   5
            Top             =   360
            Visible         =   0   'False
            Width           =   4140
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Advt"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   600
               TabIndex        =   6
               Top             =   0
               Width           =   715
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Agency"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1440
               TabIndex        =   7
               Top             =   0
               Width           =   1005
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Salesperson"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   2520
               TabIndex        =   8
               Top             =   0
               Width           =   1755
            End
         End
         Begin VB.Label lacContract 
            Appearance      =   0  'Flat
            Caption         =   "Contract #"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   57
            Top             =   2265
            Width           =   900
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Dates- Start"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   135
            TabIndex        =   31
            Top             =   60
            Width           =   1110
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "End"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2520
            TabIndex        =   30
            Top             =   60
            Visible         =   0   'False
            Width           =   405
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
         Height          =   3645
         Left            =   4965
         ScaleHeight     =   3587.371
         ScaleMode       =   0  'User
         ScaleWidth      =   4575
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CheckBox CkcAllveh 
            Caption         =   "All Contracts"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2220
            TabIndex        =   35
            Top             =   0
            Width           =   1800
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   9
            ItemData        =   "Rptselia.frx":0000
            Left            =   120
            List            =   "Rptselia.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   53
            Top             =   360
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   8
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   52
            Top             =   360
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   7
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   51
            Top             =   360
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   6
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   46
            Top             =   360
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   5
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   42
            Top             =   360
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   4
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   41
            Top             =   360
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   3
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   40
            Top             =   360
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   2
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   39
            Top             =   360
            Visible         =   0   'False
            Width           =   4350
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   1
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   4395
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   0
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   0
            Width           =   1600
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   56
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   54
      Top             =   150
      Width           =   2805
   End
   Begin VB.Frame frcoutput 
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
         Width           =   1305
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
Attribute VB_Name = "RptSelIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselia.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelIA.Frm  (duplicated from RptSelIA
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
Dim imSetAllAAS As Integer 'True=Set list box; False= don't change list box
Dim imAllClickedAAS As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSetAllVeh As Integer
Dim imAllClickedVeh As Integer
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
'Library calendar file- used to obtain post log date status
'User- used to obtain discrepancy contract that was currently being processed
'      this is used if the system gos down
Dim hmSpf As Integer            'Site file handle
Dim tmSpf As SPF                'SPF record image
Dim imSpfRecLen As Integer        'SPF record length
'Log
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
    Dim slCntrStatus As String
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        If rbcSelCSelect(0).Value Then
            If Value Then
                If lbcSelection(5).ListCount > 0 Then
                    llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                    Screen.MousePointer = vbHourglass
                    '8-24-11 found no need to populate all the contracts if ALL selected.
                    'speed up selectivity process
                    'mCntrPop slCntrStatus, 1        'get only orders (w/o revisions)
                    lbcSelection(0).Visible = True
                    lbcSelection(5).Visible = True
                    'plcContrFeed.Visible = True        'selective contracts chosen
                    If tgSpf.sSystemType = "R" Then
                        'option to include Contract spots and feed spots if Radio system; otherwise
                        'only contract spots exist
                        'plcContrFeed.Visible = True


                        '3-21-08 local, natl, regional question always asked
                        'If igFoundNatl And igFoundLocl Then         'locals & natl question, adjust position
                            plcContrFeed.Move 120, plcSelC3.Top + plcSelC3.Height
                        'Else
                        '    lacContract.Move 120, lacContract.Top + lacContract.Height
                        'End If
                    End If
                    'ckcContrFeed(0).Value = vbChecked   'default contracts spots on
                    'ckcContrFeed(1).Value = vbChecked  'default feed spots on
                    Screen.MousePointer = vbDefault
                    imSetAllVeh = False
                    llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
                CkcAllveh.Visible = False
            Else            'all advt selected
                lbcSelection(0).Visible = True
                lbcSelection(5).Visible = True
                'plcContrFeed.Visible = False        'selective contracts chosen
                ckcContrFeed(0).Value = vbChecked   'default contracts spots on
                ckcContrFeed(1).Value = vbChecked  'default feed spots on
            End If
            llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            lbcSelection(0).Visible = False
        Else
            llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        End If
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub
Private Sub CkcAllVeh_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If CkcAllveh.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value

    If imSetAllVeh Then
        If lbcSelection(0).ListCount Then
            llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
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
    If igGenRpt Then
        Exit Sub
    End If
    igGenRpt = True
    igOutput = frcoutput.Enabled
    igCopies = frcCopies.Enabled
    igFile = frcFile.Enabled
    igOption = frcOption.Enabled
    frcoutput.Enabled = False
    frcCopies.Enabled = False
    frcFile.Enabled = False
    frcOption.Enabled = False

    igUsingCrystal = True

    If rbcSelC7(2).Value Then
        ilNoJobs = 2
        ilStartJobNo = 1
    Else
        ilNoJobs = 1
        ilStartJobNo = 1
    End If
    'dan  11/04/08 display multiple reports at one time
    Set ogReport = New CReportHelper
    ogReport.iLastPrintJob = ilNoJobs
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenReportIA() Then
            igGenRpt = False
            frcoutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If
        ilRet = gCmcGenIA(imGenShiftKey, smLogUserCode)
        '-1 is a Crystal failure of gSetSelection or gSEtFormula
        If ilRet = -1 Then
            igGenRpt = False
            frcoutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
        ElseIf ilRet = 0 Then   '0 = invalid input data, stay in
            igGenRpt = False
            frcoutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        ElseIf ilRet = 2 Then           'successful return from bridge reports
            igGenRpt = False
            frcoutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
        End If
       '1 falls thru - successful crystal report
        Screen.MousePointer = vbHourglass
        If igJobRptNo = 1 Then
            gInvAffRpt
        End If
        Screen.MousePointer = vbDefault
        'coment out if below to make single reports
        If ilJobs >= ogReport.iLastPrintJob Then
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
        End If
    Next ilJobs
    Set ogReport = Nothing      'dan
    imGenShiftKey = 0
    Screen.MousePointer = vbHourglass
    gIvrClear
    Screen.MousePointer = vbDefault
    igGenRpt = False
    frcoutput.Enabled = igOutput
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
    If ((KeyAscii <> KEYBACKSPACE) And (KeyAscii <= 32)) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KEYDOWN) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcSelCTo_Change()
    mSetCommands
End Sub

Private Sub edcSelCTo_GotFocus()
    gCtrlGotFocus edcSelCTo
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
    RptSelIA.Refresh
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
    'RptSelIA.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set RptSelIA = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub



Private Sub lbcRptType_Click()

    mMorelbcRptType
    mSetCommands
End Sub
Private Sub lbcSelection_Click(Index As Integer)

    Dim slCntrStatus As String
    Dim slName As String

    Dim ilIdx As Integer
    If Not imAllClicked Then
'        If Not imAllClicked Then
            If Index = 5 Then
                slCntrStatus = "HO"
                mCntrPop slCntrStatus, 1        'get only orders (w/o revisions)
                If tgSpf.sSystemType = "R" Then     'station system (vs network or syndicator); may need to see network spots
                    slName = "99999999|999-999||999||[Feed Spots]\0"
                    lbcCntrCode.AddItem slName, 0
                    lbcSelection(0).AddItem "[Feed Spots]", 0 'Add ID to list box"
                    'plcContrFeed.Visible = False        'selective contracts chosen, turn off generic
                    'selectivity question for contract & feed spots.  Use the contract list box for selection
                    ckcContrFeed(0).Value = vbChecked   'default contracts spots on
                    ckcContrFeed(1).Value = vbChecked  'default feed spots on
                End If

                lbcSelection(0).Visible = True
                imSetAllVeh = False
                CkcAllveh.Value = vbUnchecked   'False
                imSetAllVeh = True
                CkcAllveh.Visible = False
                For ilIdx = 0 To lbcSelection(5).ListCount - 1 Step 1
                    If lbcSelection(5).Selected(ilIdx) Then
                        CkcAllveh.Visible = True
                        Exit For
                    End If
                Next
                imSetAll = False
                ckcAll.Value = vbUnchecked  'False
                imSetAll = True
                If imTerminate Then
                    cmcCancel_Click
                    Exit Sub
                End If
            Else                    '8-27-19
                imSetAllVeh = False
                CkcAllveh.Value = vbUnchecked   'False
                imSetAllVeh = True
            End If
'        Else
'            imSetAll = False
'            ckcAll.Value = vbUnchecked  'False
'            imSetAll = True
'        End If
    End If
    mSetCommands
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    gCtrlGotFocus lbcSelection(Index)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAdvtPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Advertiser list box   *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAdvtPop(lbcSelection As Control)
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(RptSelIA, lbcSelection, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(RptSelIA, lbcSelection, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", RptSelIA
        On Error GoTo 0
    End If
    Exit Sub
mAdvtPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAgencyPop                      *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Agency list box       *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAgencyPop(lbcSelection As Control)
'
'   mAgencyPop
'   Where:
'
    Dim ilRet As Integer
    'ilRet = gPopAgyBox(RptSelIA, lbcSelection, Traffic!lbcAgency)
    ilRet = gPopAgyBox(RptSelIA, lbcSelection, tgAgency(), sgAgencyTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgencyPopErr
        gCPErrorMsg ilRet, "mAgencyPop (gPopAgyBox)", RptSelIA
        On Error GoTo 0
    End If
    Exit Sub
mAgencyPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCntrPop                        *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                    :7/10/96 -Use new contract status*
'                                                      *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mCntrPop(slCntrStatus As String, ilHOState As Integer)
'
'   mCntrPop
'   Where:
'       slcntrStatus(I)- O; H; W; C; I; D or blank for all
'       ilHOState(I) - 1 only get cnt (w/o revision) H & O only
'                      2 combo - get latest orders includ revisions (H O G or N) if G or N, show instead of the H or O
'                      3 everything - revision & orders (HOGNWCI) if GNWCI, show over the H or O
'
'
    Dim ilRet As Integer 'btrieve status
    Dim slNameCode As String  'Name and code
    Dim slCode As String    'Code number
    Dim ilCurrent As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slName As String
    Dim llCntrNo As Long
    Dim ilShow As Integer
    Dim slCntrType As String
    Dim ilAdfCode As Integer
    Dim llLen As Long
    Dim ilErr As Integer
    Dim slShow As String
    Dim ilRevNo As Integer
    Dim ilVerNo As Integer
    Dim ilExtRevNo As Integer
    Dim slRevNo As String
    llLen = 0
    ilErr = False
    lbcSelection(0).Clear
    lbcCntrCode.Clear
    For ilLoop = 0 To lbcSelection(5).ListCount - 1 Step 1
        If lbcSelection(5).Selected(ilLoop) Then
            sgMultiCntrCodeTag = ""             'init the date stamp so the box will be populated
            sgMultiCntrCodeTagIA = ""           'make sure the cntracts are populated for reentrant problem
            ReDim tgMultiCntrCodeIA(0 To 0) As SORTCODE
            lbcMultiCntr.Clear
            slNameCode = tgAdvertiser(ilLoop).sKey  'Traffic!lbcAdvertiser.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
            'ilCurrent = 1   '0=Current; 1=All
            'ilFilter = Val(slCode)   'by contract #; -101=by advertiser
            'ilVehCode = -1  'All vehicles
            'ilRet = gPopCntrBox(RptSelIA, ilCntrType, ilFilter, ilCurrent, 0, ilVehCode, lbcMultiCntr, lbcMultiCntrCode, True, False, False, False)
            'slCntrType = ""                                 'all Types
            slCntrType = "C"
            If tgUrf(0).sResvType <> "H" Then
                slCntrType = slCntrType & "V"
            End If
            If tgUrf(0).sRemType <> "H" Then
                slCntrType = slCntrType & "T"
            End If
            If tgUrf(0).sDRType <> "H" Then
                slCntrType = slCntrType & "R"
            End If
            If tgUrf(0).sPIType <> "H" Then
                slCntrType = slCntrType & "Q"
            End If
            If tgUrf(0).sPSAType <> "H" Then
                slCntrType = slCntrType & "S"
            End If
            If tgUrf(0).sPromoType <> "H" Then
                slCntrType = slCntrType & "M"
            End If
            If slCntrType = "CVTRQSM" Then
                slCntrType = ""
            End If
            'ilShow = 1
            ilShow = 5                  'show # and advt name
            ilCurrent = 1
            ilAdfCode = Val(slCode)
            'load up list box with contracts with matching adv
            'ilRet = gPopCntrForAASBox(RptSelIA, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, lbcMultiCntrCode)
            
            ilRet = gPopCntrForAASBox(RptSelIA, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, tgMultiCntrCodeIA(), sgMultiCntrCodeTagIA)
            If ilRet <> CP_MSG_NOPOPREQ Then
                On Error GoTo mCntrPopErr
                gCPErrorMsg ilRet, "mCntrPop (gPopCntrBox)", RptSelIA
                On Error GoTo 0
            End If
            For ilIndex = 0 To UBound(tgMultiCntrCodeIA) - 1 Step 1 'lbcMultiCntrCode.ListCount - 1 Step 1
                slName = Trim$(tgMultiCntrCodeIA(ilIndex).sKey)  'lbcMultiCntrCode.List(ilIndex)
                If Not gOkAddStrToListBox(slName, llLen, True) Then
                    ilErr = True
                    Exit For
                End If
                lbcCntrCode.AddItem slName  'lbcMultiCntrCode.List(ilIndex)
            Next ilIndex

            If ilErr Then
                Exit For
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To lbcCntrCode.ListCount - 1 Step 1
        slNameCode = lbcCntrCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 1, "|", slCode)
        llCntrNo = 99999999 - CLng(slCode)
        slShow = Trim$(str$(llCntrNo))
        ilRet = gParseItem(slName, 2, "|", slCode)
        ilRet = gParseItem(slCode, 1, "-", slRevNo)
        ilRevNo = 999 - CLng(slRevNo)
        ilRet = gParseItem(slCode, 2, "-", slRevNo)
        ilExtRevNo = 999 - CLng(slRevNo)
        ilRet = gParseItem(slName, 4, "|", slCode)
        ilVerNo = 999 - CLng(slCode)
        ilRet = gParseItem(slName, 5, "|", slCode)
        If (slCode = "W") Or (slCode = "C") Or (slCode = "I") Or (slCode = "D") Then
            If (ilRevNo > 0) Then
                slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
            Else
                slShow = slShow & " V" & Trim$(str$(ilVerNo))
            End If
        Else
            slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
        End If
        ilRet = gParseItem(slName, 6, "|", slCode)

        If Trim$(slCode) = "[Feed Spots]" Then
            slShow = slCode
        Else
            slShow = slShow & " " & slCode
        End If

        lbcSelection(0).AddItem Trim$(slShow)  'Add ID to list box
    Next ilLoop
    
    Screen.MousePointer = vbDefault
    Exit Sub
mCntrPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim ilMultiTable As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slStamp As String

    Screen.MousePointer = vbHourglass

    'D.S. 1-23-02 Gather the sales sources for local/national question.  This applies only to those sites whose
    'sales source origins have combination of local & national
    igFoundLocl = False
    igFoundNatl = False
    igFoundRegl = False
    'ReDim tlMnf(1 To 1) As MNF
    ReDim tlMnf(0 To 0) As MNF
    ilRet = gObtainMnfForType("S", slStamp, tlMnf())
    For ilLoop = LBound(tlMnf) To UBound(tlMnf) - 1
        If tlMnf(ilLoop).iGroupNo = 1 Then      'Local
            igFoundLocl = True
        End If
        If tlMnf(ilLoop).iGroupNo = 2 Then      'Regional
            igFoundRegl = True
        End If
        If tlMnf(ilLoop).iGroupNo = 3 Then      'National
            igFoundNatl = True
        End If
    Next ilLoop
    'End D.S. 1-23-02

    imFirstActivate = True
    slStr = ""
    'ReDim tgMkMnf(1 To 1) As MNF
    ReDim tgMkMnf(0 To 0) As MNF
    ilRet = gObtainMnfForType("H3", slStr, tgMkMnf())
    rbcSelC7(0).Value = True
    imSetAllVeh = True
    'Start Crystal report engine
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        Exit Sub
    End If
    'Set options for report generate
'VB6**    hdJob = rpcRpt.hJob
    ilMultiTable = True

    RptSelIA.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    imSetAllAAS = True
    imAllClickedAAS = False
    imAllClickedVeh = False
    imSetAllVeh = True
    'cbcSel.Move 120, 30
    plcSelC3.Height = 240
    plcSelC3.Top = 1600
    plcSelC3.Left = 120
'    lacSelCFrom.Move 120, 75
'    lacSelCFrom1.Move 2400, 75
'    edcSelCFrom.Move 1500, 30
'    edcSelCFrom1.Move 3240, 30
'    edcSelCTo.Move 1500, 345
'
    plcSelC1.Move 120, 675
   
  
    lbcSelection(0).Move 15, ckcAll.Height + 30, 4380, 3270
    lbcSelection(1).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
'    pbcSelC.Move 90, 255, 4515, 3660
    gCenterStdAlone RptSelIA
End Sub
'
'
'           mInitControls - set controls to proper positions, sizes
'                   hidden, shown, etc.
'
'           Created :  11/28/98 D Hosaka
'
Private Sub mInitControls()
    lbcSelection(0).Move 15, ckcAll.Height + 30, 4380, 3270
    lbcSelection(1).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(2).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(3).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(5).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(5).Move lbcSelection(5).Left, lbcSelection(5).Top, lbcSelection(5).Width / 2 - 30, lbcSelection(5).Height  '1110
    lbcSelection(6).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(0).Move lbcSelection(5).Left + lbcSelection(5).Width + 60, lbcSelection(0).Top, lbcSelection(0).Width / 2 - 30, lbcSelection(0).Height '840
    lbcSelection(7).Move lbcSelection(5).Left, lbcSelection(5).Top, lbcSelection(5).Width, lbcSelection(5).Height       'advt
    lbcSelection(8).Move lbcSelection(5).Left, lbcSelection(5).Top, lbcSelection(5).Width, lbcSelection(5).Height       'agy
    lbcSelection(9).Move lbcSelection(5).Left, lbcSelection(5).Top, lbcSelection(5).Width, lbcSelection(5).Height       'slsp
    'lbcSelection(10).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height      'cnt
    'lbcSelection(11).Move lbcSelection(6).Left, lbcSelection(6).Top, lbcSelection(0).Width, lbcSelection(6).Height      'demo

    lbcSelection(0).Visible = False
    lbcSelection(1).Visible = False
    lbcSelection(2).Visible = False
    lbcSelection(3).Visible = False
    lbcSelection(4).Visible = False
    lbcSelection(5).Visible = False
    lbcSelection(6).Visible = False
    lbcSelection(7).Visible = False
    lbcSelection(8).Visible = False
    lbcSelection(9).Visible = False
    'lbcSelection(10).Visible = False
    'lbcSelection(11).Visible = False
    'lbcSelection(12).Visible = False

    lacSelCFrom.Visible = False
'    edcSelCFrom.Visible = False            '8-26-19 use csi calendar control vs edit box
    lacSelCFrom1.Visible = False
'    edcSelCFrom1.Visible = False

    edcSelCTo.Visible = False           'contract #
    plcSelC1.Visible = False

    plcSelC1.Visible = False

    plcSelC7.Visible = False
    
    '8-24-11 remove changing of controls
'    lacSelCFrom.Move 120, 75, 1380
'    edcSelCFrom.Move 1500, 30, 1350
'
'    edcSelCTo.Move 1500, 345, 1350
'    edcSelCTo.MaxLength = 10    '8 5/27/99 changed for short form date m/d/yyyy
'
'    edcSelCFrom.MaxLength = 10  '8  5/27/99 changed for short form date m/d/yyyy
'    edcSelCFrom1.Move 3240, 30
'    lacSelCFrom1.Move 2340, 75
'    edcSelCTo.Text = ""
    
    plcSelC1.Top = 675
    plcSelC1.Left = 120

    rbcSelCSelect(2).Top = 0
    plcSelC1.Height = 240
    plcSelC3.Height = 240

    plcSelC3.Visible = False



    'ckcSelC6(0).Enabled = vbChecked 'True
    'pbcSelC.Height = 3195
    edcSelCTo.Text = ""
    ckcAll.Move lbcSelection(1).Left            'readjust 'Check All' location to be above left most list box
    'ckcAllAAS.Move ckcAll.Left, ckcAll.Top
    ckcAll.Enabled = True

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitReport                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'
'             5/5/98 - Set frcOption.Visible to false
'             at developement, at run time set to
'             True when all questions on screen have
'             been formatted.  This way the screen comes
'             up all at once, rather than pieces.
'
'       6-16-00 Remove all references to Contract "BR"
'               and Insertion Orders (reports are coded
'               in rptselct)
'
'*                                                     *
'*******************************************************
Private Sub mInitReport()
    Dim ilIndex As Integer
    Dim ilRet As Integer
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    gPopExportTypes cbcFileType     '10-20-01
    'pbcSelC.Visible = False
    'sgPhoneImage = mkcPhone.Text
    lbcRptType.Clear

    lgOrigCntrNo = 0
    hmSpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        imSpfRecLen = Len(tmSpf)
        ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            lgOrigCntrNo = tmSpf.lDiscCurrCntrNo
        End If
        ilRet = btrClose(hmSpf)
        btrDestroy hmSpf
    End If
    'RptSelIA.Caption = "Contract Report Selection"
    mAdvtPop lbcSelection(5)    'Called to initialize Traffic!Advertiser required be mCntrPop
    If imTerminate Then
        Exit Sub
    End If
    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    Screen.MousePointer = vbHourglass
    mAgencyPop lbcSelection(1)
    If imTerminate Then
        Exit Sub
    End If
    'mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
                                        'populate when needed
    'mSellConvVehPop 3
    If imTerminate Then
        Exit Sub
    End If
    'mSellConvVirtVehPop 6, False
    'lbcselection(11) used for demos (cpp/cpm report) and single select budgets (tieout report), populate when needed
    'ilRet = gPopMnfPlusFieldsBox(RptSelIA, lbcSelection(11), lbcDemoCode, "D")
    lbcRptType.AddItem "Proposals/Contracts", 0         '0=proposal
    lbcRptType.AddItem "Paperwork Summary", 1           '1=paperwork summary (contract summaries)

    'If tgUrf(0).islfCode = 0 Then           'its a slsp thats is asking for this report,
                                            'don't allow them to exclude reserves
        ilIndex = 2
        If igRptType = 0 Then   'Proposal
            'rbcRptType(2).Visible = False
        Else    'Contract
            'rbcRptType(2).Caption = "Spots by Advt"
            lbcRptType.AddItem "Spots by Advertiser", ilIndex   '2=spots by advt
            ilIndex = ilIndex + 1
        End If
        lbcRptType.AddItem "Spots by Date & Time", ilIndex      '3=spots by date & time
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Business Booked by Contract", ilIndex  '4=projection (named changed to Business Booked)
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Contract Recap", ilIndex            '5=contr recap
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Spot Placements", ilIndex           '6=Spot placements
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Spot Discrepancies", ilIndex        '7=spot discrepancies
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "MG's", ilIndex                      '8=makegood
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Sales Spot Tracking", ilIndex       '9=sales spot traking
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Commercial Changes", ilIndex        '10=coml changes
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Contract History", ilIndex          '11 Contract history
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Affiliate Spot Tracking", ilIndex   '12 affil spot traking
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Spot Sales", ilIndex                '13=spot sales
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Missed Spots", ilIndex              '14=missed spots
        ilIndex = ilIndex + 1

        'lbcRptType.AddItem "Business Booked by Spot", ilIndex    '15=spot projection (name changed to Business Booked)
        '1-31-00 name chg from business booked by spot to spot business booked
        lbcRptType.AddItem "Spot Business Booked", ilIndex    '15=spot projection (name changed to Business Booked)
        ilIndex = ilIndex + 1
        'spot reprints - used
        lbcRptType.AddItem "Business Booked by Spot Reprint", ilIndex   '16= Business booked reprint
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Avails", ilIndex                    '17=quarterly summary & detail avails
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Average Spot Prices", ilIndex       '18=avg spot prices
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Advertiser Units Ordered", ilIndex  '19=advt units ordered
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Sales Analysis by CPP & CPM", ilIndex '20=sales analysis by cpp & cpm
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Average Rate", ilIndex            '21=Average Rate
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Tie-Out", ilIndex                  '22=Detail Tie Out
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Billed and Booked", ilIndex        '23=Billed & booked by advt, Slsp, owner, vehicle
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Weekly Sales Activity by Quarter", ilIndex   '24=Sales Activity
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Sales Comparison", ilIndex         'Sales Comparison by Advt, Slsp, Agy, comp code, Bus code
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Weekly Sales Activity by Month", ilIndex       'Cumulative Activity Report (pacing)
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Average Prices to Make Plan", ilIndex       'Avg Prices needed to make plan
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "CPP/CPM by Vehicle", ilIndex        'Curent cpp/cpm by vehicle
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Sales Analysis Summary", ilIndex        'Sales Analysis Summary
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Insertion Orders", ilIndex
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Makegood Revenue", ilIndex
        ilIndex = ilIndex + 1
        lbcRptType.AddItem "Affidavit of Performance", ilIndex
        ilIndex = ilIndex + 1
    'End If
    'frcOption.Caption = "Contract Selection"
    ckcAll.Caption = "All xContracts"
    frcOption.Enabled = True
'    pbcSelC.Height = pbcSelC.Height - 60
    'lbcSelection(0).Move lbcSelection(0).Left, pbcSelC.Height + 150, lbcSelection(0).Width, ckcAll.Top - pbcSelC.Height
    'rbcSelCSelect(0).Value = True   'Advertiser/Contract #
    'lbcRptType.ListIndex = 0
    pbcSelC.Visible = True
    pbcOption.Visible = True
    lacSelCFrom.Visible = False

'    edcSelCFrom.Visible = False
    edcSelCTo.Visible = False
    ckcAll.Visible = False


    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
        'RptSelIA.Caption = smSelectedRptName & " Report"
        'frcOption.Caption = smSelectedRptName & " Selection"
        'slStr = Trim$(smSelectedRptName)
        'ilLoop = InStr(slStr, "&")
        'If ilLoop > 0 Then
        '    slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
        'End If
        'frcOption.Caption = slStr & " Selection"
    End If
    mSetCommands
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSel
End Sub
'***********************************************************************
'*                                                                     *
'*      Procedure Name:mMorelbcRptType                                 *
'*                                                                     *
'*             Created:5/17/93       By:D. LeVine                      *
'*            Modified:              By:D. Smith                       *
'*                                                                     *
'***********************************************************************
Private Sub mMorelbcRptType()
    Dim slMonth As String
    Dim slDate As String
    Dim slTime As String
    Dim slDay As String
    Dim slYear As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilTop As Integer

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    
    '8-23-11 change to start/end date input
'    slMonth = gMonthName(slDate)
'    edcSelCFrom.MaxLength = 3           'Jan, ....dec
'    edcSelCFrom.Text = Trim$(slMonth)
'    edcSelCFrom1.MaxLength = 4          'year 1996..2000...
'    edcSelCFrom1.Text = Trim$(slYear)
'    lbcSelection(0).Clear
'    lbcSelection(0).Tag = ""
'
'    If tgSpf.sRUseCorpCal = "Y" Then       'if Using corp cal,retain in memory
'        ilRet = gObtainCorpCal()
'    End If

    'ilRet = gObtainVef()
    mInitControls           'set controls to proper positions, widths, hidden, shown, etc.
    'mSellConvVirtVehPop 6, False
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
                                        'populate when needed
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    ckcAll.Visible = True
    ckcAll.Enabled = True
    
    '8-24-11 remove changing of controls, form already has placement
'    lacSelCFrom.Left = 120
'    lacSelCFrom.Width = 600
'    lacSelCFrom1.Move 2105, 75
'    edcSelCFrom.Move 800, edcSelCFrom.Top, 945
'    edcSelCFrom1.Move 2620, edcSelCFrom.Top, 945

'    8-27-19 remove references to edit boxes for start/end dates, use csi calendar controls
'    edcSelCTo.Move 1110, 1150
'    edcSelCTo.MaxLength = 10    '8  5/27/99 changed for short form date m/d/yyyy
'    edcSelCFrom.MaxLength = 10  '8  5/27/99 changed for short form date m/d/yyyy
'    edcSelCFrom1.MaxLength = 10 '8  5/27/99 changed for short form date m/d/yyyy

    '8-23-11 remove month/year input, change to start/end dates due to end of contract billing
    'lacSelCFrom.Caption = "Month"
    'lacSelCFrom1.Caption = "Year"
    lacSelCFrom.Visible = True
    lacSelCFrom1.Visible = True

'    edcSelCFrom.Visible = True
'    edcSelCFrom1.Visible = True
    edcSelCTo.Visible = True

    'plcSelC3.Visible = False
'    plcSelC1.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 60
    plcSelC1.Move 120, CSI_CalFrom.Top + CSI_CalFrom.Height + 60
    rbcSelCSelect(0).Caption = "Advt"
    rbcSelCSelect(0).Move 600, 0, 675
    rbcSelCSelect(1).Caption = "Agency"
    rbcSelCSelect(1).Move 1290, 0, 980
    rbcSelCSelect(1).Visible = True
    rbcSelCSelect(2).Caption = "Salesperson"
    rbcSelCSelect(2).Move 2250, 0, 1350   '2220
    rbcSelCSelect(1).Enabled = True
    rbcSelCSelect(2).Enabled = True
    rbcSelCSelect(2).Visible = True
    plcSelC1.Visible = True
    lbcSelection(3).Visible = False
    lbcSelection(1).Visible = False
    lbcSelection(2).Visible = False
    lbcSelection(5).Visible = False
    lbcSelection(0).Visible = False
    lbcSelection(4).Visible = False
    'lbcSelection(6).Visible = False
    If rbcSelCSelect(0).Value Then
        lbcSelection(1).Visible = False
        lbcSelection(2).Visible = False
        lbcSelection(5).Visible = True
        lbcSelection(0).Visible = True
        ckcAll.Caption = "All Advertisers"
        ckcAll.Visible = True
    ElseIf rbcSelCSelect(1).Value Then
        lbcSelection(0).Visible = False
        lbcSelection(2).Visible = False
        lbcSelection(5).Visible = False
        lbcSelection(1).Visible = True
        ckcAll.Caption = "All Agencies"
        ckcAll.Visible = True
    ElseIf rbcSelCSelect(2).Value Then
        lbcSelection(0).Visible = False
        lbcSelection(1).Visible = False
        lbcSelection(5).Visible = False
        lbcSelection(2).Visible = True
        ckcAll.Caption = "All Salespeople"
        ckcAll.Visible = True
    End If
    rbcSelCSelect(0).Value = True           'force default to advt

    plcSortBy.Move 120, plcSelC1.Top + plcSelC1.Height

    'Set the default to show Status column, status will always show,
    'Dont use plcselc7 for any new options.  this is tested in prepass
    plcSelC7.Move 120, plcSortBy.Top + plcSortBy.Height   '780
    'plcSelC7.Caption = "Show"
    rbcSelC7(0).Caption = "Detail"
    rbcSelC7(1).Caption = "Summary"
    rbcSelC7(2).Caption = "Both"
    rbcSelC7(0).Left = 620
    'rbcSelC7(1).Left = 720
    rbcSelC7(1).Left = 1420
    rbcSelC7(2).Left = 2580
    plcSelC7.Visible = True
    rbcSelC7(0).Visible = True
    rbcSelC7(1).Visible = True
    rbcSelC7(2).Visible = True
    plcSelC7.Visible = True

'    plcSelC10.Move 120, plcSelC7.Top + plcSelC7.Height, 4000
'    ckcSelC10(0).Move 0, 0, 4000
    ckcSelC10(0).Caption = "Skip to new page each vehicle"

    ckcSelC10(0).Visible = True
    plcSelC10.Visible = True

    'ckcShowRate.Move 120, plcSelC10.Top + plcSelC10.Height

    'option to show inv #
'    plcShowInvNo.Move 120, ckcShowRate.Top + ckcShowRate.Height
    plcShowInvNo.Visible = True

    lacContract.Move 120, ckcInclCreativeTitle.Top + ckcInclCreativeTitle.Height + 60
    edcSelCTo.Move 1110, ckcInclCreativeTitle.Top + ckcInclCreativeTitle.Height + 30

    'D.S. 11-23-02
    'set default values to true
    ckcSelC3(0).Value = vbChecked
    ckcSelC3(1).Value = vbChecked
    ckcSelC3(2).Value = vbChecked
    plcSelC3.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
    'If igFoundNatl And igFoundLocl Then        '3-21-08 always show all types
        ckcSelC3(0).Left = 1290 '1650
        ckcSelC3(0).Width = 1800
        ckcSelC3(0).Top = -15
        ckcSelC3(0).Caption = "National"
        ckcSelC3(0).Visible = True

        ckcSelC3(1).Top = ckcSelC3(0).Top
        ckcSelC3(1).Left = 2290 '2650
        ckcSelC3(1).Width = 840
        ckcSelC3(1).Caption = "Local"
        ckcSelC3(1).Visible = True

        ckcSelC3(2).Top = ckcSelC3(0).Top
        ckcSelC3(2).Left = 3130 '2650
        ckcSelC3(2).Width = 1080
        ckcSelC3(2).Caption = "Regional"
        ckcSelC3(2).Visible = True

        'plcSelC3.Caption = "Show Sales Source"
        plcSelC3.Visible = True
    'End If
    'End D.S. 1-23-02
    plcContrFeed.Visible = True
    If tgSpf.sSystemType = "R" Then
        'option to include Contract spots and feed spots if Radio system; otherwise
        'only contract spots exist
        'plcContrFeed.Visible = True

        '3-21-08 local, natl and regional question always asked
        'If igFoundNatl And igFoundLocl Then         'locals & natl question, adjust position
            plcContrFeed.Move 120, plcSelC3.Top + plcSelC3.Height
        'Else
        '    lacContract.Move 120, lacContract.Top + lacContract.Height
        'End If
        ilTop = plcSelC3.Top + plcSelC3.Height
    Else
        plcContrFeed.Move 120, plcSelC3.Top + plcSelC3.Height
        ckcContrFeed(1).Visible = False
        ckcContrFeed(1).Value = vbUnchecked
        ilTop = plcContrFeed.Top + plcContrFeed.Height
    End If

    If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
        For ilLoop = LBound(tgVff) To UBound(tgVff) - 1         'any that is using station inv imports?
            If tgVff(ilLoop).sPostLogSource = "S" Then
                ckcUseCountAff.Visible = True
                ckcUseCountAff.Value = vbChecked
                ckcUseCountAff.Move 120, ilTop + 30
                Exit For
            End If
        Next ilLoop
    End If
    
    pbcSelC.Visible = True
    pbcOption.Visible = True
    CkcAllveh.Visible = False
    frcOption.Visible = True  'Make all visible

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
    'gInitStdAlone RptSelIA, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Affidavit of Performance"  '"Spot Placements"'"Spot Discrepancies" '"Spot Sales"
    '    igRptCallType = -1
    '    igRptType = -1   '3 'Log '1   'Contract    '0   'Summary '3 Program  '1  links
    'Else
        ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
        If (ilRet = CP_MSG_NONE) Then
            ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
            ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
            igRptCallType = Val(slStr)
     'smSelectedRptName = "Affidavit of Performance"  '"Spot Placements"'"Spot Discrepancies" '"Spot Sales"
     'igRptCallType = CONTRACTSJOB 'SLSPCOMMSJOB   'LOGSJOB 'CONTRACTSJOB 'COPYJOB 'COLLECTIONSJOB'CONTRACTSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB 'PROPOSALPROJECTION
     'igRptType = 1   '3 'Log '1   'Contract    '0   'Summary '3 Program  '1  links
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
'       8-25-01 dh Set Generate button enabled if single
'           contract and the month or year changed
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    Dim ilEnable As Integer
    Dim ilLoop As Integer
    Dim MnthYrFlag As Integer
    'Make sure the a month and year exist before allowing generate button
'    If ((edcSelCFrom.Text = "") Or (edcSelCFrom1.Text = "")) Then
    If ((CSI_CalFrom.Text = "") Or (CSI_CalTo.Text = "")) Then          '8-27-19 use csi cal controls vs edit boxes

        MnthYrFlag = False
    Else
        MnthYrFlag = True
    End If
    If gSetCheck(ckcAll.Value) Then
        ilEnable = True
    Else
        If edcSelCTo.Text <> "" Then                   'something entered in singel contract
            ilEnable = True
        Else
            If rbcSelCSelect(0).Value Then                  'advt, get selective cnts
                'Can't use SelCount as property does not exist for ListBoxbox
                If gSetCheck(ckcAll.Value) Then
                    ilEnable = True
                Else
                    For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                        If lbcSelection(0).Selected(ilLoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next ilLoop
                End If
            End If
            If rbcSelCSelect(1).Value Then                    'agy
                For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
                    If lbcSelection(1).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
            ElseIf rbcSelCSelect(2).Value Then               'slsp
                For ilLoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                    If lbcSelection(2).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
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
    cmcGen.Enabled = ilEnable And MnthYrFlag
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
    'ilRet = gPopSalespersonBox(RptSelIA, 0, True, True, lbcSelection, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(RptSelIA, 0, True, True, lbcSelection, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gPopSalespersonBox)", RptSelIA
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
    Unload RptSelIA
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcContrFeed_Paint()
    plcContrFeed.Cls
    plcContrFeed.CurrentX = 0
    plcContrFeed.CurrentY = 0
    plcContrFeed.Print "Include"
End Sub

Private Sub plcSortBy_Paint()
    plcSortBy.CurrentX = 0
    plcSortBy.CurrentY = 0
    plcSortBy.Print "Sort by"
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
Private Sub rbcSelCSelect_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelCSelect(Index).Value
    'End of coded added

    Dim ilIdx As Integer

    If Value Then
        Select Case Index
            Case 0  'Advertiser/Contract #
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(5).Visible = True
                For ilIdx = 0 To lbcSelection(5).ListCount - 1 Step 1
                    If lbcSelection(5).Selected(ilIdx) Then
                        CkcAllveh.Visible = True
                        Exit For
                    Else
                        CkcAllveh.Visible = False
                    End If
                Next
                ckcAll.Caption = "All Advertisers"
                ckcAll.Visible = True
            Case 1  'Agency
                lbcSelection(0).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(5).Visible = False
                lbcSelection(1).Visible = True
                CkcAllveh.Visible = False
                ckcAll.Visible = False
                ckcAll.Caption = "All Agencies"
                ckcAll.Visible = True
            Case 2  'Salesperson
                lbcSelection(0).Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(5).Visible = False
                lbcSelection(2).Visible = True
                CkcAllveh.Visible = False
                ckcAll.Visible = False
                ckcAll.Caption = "All Salespeople"
                ckcAll.Visible = True
        End Select

        lbcSelection(0).Height = 3250           'cnt list box
        lbcSelection(5).Height = 3250           'advt list box
        lbcSelection(2).Height = 3250           'slsp list box
        lbcSelection(1).Height = 3250           'agy list box
        Select Case Index
            Case 0  'Advertiser/Contract #
                lbcSelection(0).Visible = True
            Case 1  'Agency
            Case 2  'Salesperson
            'Case 3  'vehicles
        End Select
        mSetCommands
    End If
End Sub

Private Sub rbcSortBy_Click(Index As Integer)
    If Index = 0 Then
        ckcSelC10(0).Caption = "Skip to new page each vehicle"
        ckcShowRate.Visible = True  'false  9-17-08
        rbcSelC7(1).Enabled = True
        rbcSelC7(2).Enabled = True
    Else
        ckcSelC10(0).Caption = "Skip to new page each Product/ISCI"
        ckcShowRate.Visible = True
        rbcSelC7(0).Value = True
        rbcSelC7(1).Enabled = False
        rbcSelC7(2).Enabled = False
    End If
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcSelC10_Paint()
    plcSelC10.Cls
    plcSelC10.CurrentX = 0
    plcSelC10.CurrentY = 0
    plcSelC10.Print ""
End Sub
Private Sub plcSelC7_Paint()
    plcSelC7.CurrentX = 0
    plcSelC7.CurrentY = 0
    plcSelC7.Print "Show"
End Sub
Private Sub plcSelC3_Paint()
    plcSelC3.CurrentX = 0
    plcSelC3.CurrentY = 0
    plcSelC3.Print "Sales Source"
End Sub
Private Sub plcSelC1_Paint()
    plcSelC1.CurrentX = 0
    plcSelC1.CurrentY = 0
    plcSelC1.Print "Select"
End Sub
