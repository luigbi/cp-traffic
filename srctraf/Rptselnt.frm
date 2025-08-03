VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelNT 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5670
   ClientLeft      =   495
   ClientTop       =   1485
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
   ScaleHeight     =   5670
   ScaleWidth      =   9270
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   17
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6690
      TabIndex        =   50
      Top             =   -90
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Timer tmcDone 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8835
      Top             =   -150
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
      Left            =   8025
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   -75
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
      Left            =   8310
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   -90
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
      ScaleWidth      =   90
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4245
      Width           =   90
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   3600
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
      Left            =   2055
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
         Left            =   1095
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "1"
         Top             =   315
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   165
         TabIndex        =   7
         Top             =   810
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   345
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
      Caption         =   "NTR"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4120
      Left            =   90
      TabIndex        =   14
      Top             =   1425
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
         Height          =   3825
         Left            =   45
         ScaleHeight     =   3825
         ScaleWidth      =   4770
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4770
         Begin VB.CheckBox ckcInclNonPolit 
            Caption         =   "Non Polit"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   74
            Top             =   2880
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CheckBox ckcInclPolit 
            Caption         =   "Polit"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2160
            TabIndex        =   73
            Top             =   2880
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   810
         End
         Begin V81TrafficReports.CSI_Calendar csi_CalTo 
            Height          =   255
            Left            =   2280
            TabIndex        =   72
            Top             =   0
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   450
            Text            =   "01/10/2024"
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
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   600
            TabIndex        =   71
            Top             =   0
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   450
            Text            =   "01/10/2024"
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
         Begin VB.CheckBox ckcMinorSplit 
            Caption         =   "Split"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   70
            Top             =   1320
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.CheckBox ckcMajorSplit 
            Caption         =   "Split"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   69
            Top             =   960
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.CheckBox ckcUseAcqCost 
            Caption         =   "Use Acquisition Cost"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   105
            TabIndex        =   65
            Top             =   3195
            Visible         =   0   'False
            Width           =   2610
         End
         Begin VB.CheckBox ckcInclHardCost 
            Caption         =   "Hard Cost"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   64
            Top             =   2880
            Visible         =   0   'False
            Width           =   1170
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
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   55
            Top             =   3480
            Width           =   1035
         End
         Begin VB.CheckBox ckcSkipPage 
            Caption         =   "Skip to new page each new group"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   2400
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3210
         End
         Begin VB.CheckBox ckcShowDescr 
            Caption         =   "Show NTR Description"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   2640
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.TextBox lacVehGroup 
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
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "Vehicle Group"
            Top             =   840
            Width           =   1215
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
            TabIndex        =   33
            Top             =   1320
            Width           =   1500
         End
         Begin VB.TextBox lacSet2 
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
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "Minor Sort"
            Top             =   1440
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
            TabIndex        =   31
            Top             =   960
            Width           =   1500
         End
         Begin VB.TextBox lacSet1 
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
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "Major Sort"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.PictureBox plcBillBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4200
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2160
            Width           =   4200
            Begin VB.OptionButton rbcBillBy 
               Caption         =   "Calendar"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2520
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   0
               Width           =   1185
            End
            Begin VB.OptionButton rbcBillBy 
               Caption         =   "Corporate"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   0
               TabIndex        =   48
               Top             =   0
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton rbcBillBy 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1320
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   0
               Width           =   1185
            End
         End
         Begin VB.PictureBox plcGrossNet 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   3480
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3480
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   960
               TabIndex        =   46
               Top             =   0
               Width           =   825
            End
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.PictureBox plcTotalsBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4440
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1680
            Width           =   4440
            Begin VB.OptionButton rbcTotalsBy 
               Caption         =   "Summary"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3240
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   0
               Width           =   1110
            End
            Begin VB.OptionButton rbcTotalsBy 
               Caption         =   "Contract"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   960
               TabIndex        =   41
               Top             =   0
               Value           =   -1  'True
               Width           =   1110
            End
            Begin VB.OptionButton rbcTotalsBy 
               Caption         =   "Advertiser"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2040
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   0
               Width           =   1305
            End
         End
         Begin VB.TextBox edcPeriods 
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
            Left            =   4080
            MaxLength       =   2
            TabIndex        =   26
            Top             =   0
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.TextBox edcDate2 
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
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   24
            Top             =   0
            Width           =   810
         End
         Begin VB.TextBox edcDate1 
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
            Left            =   600
            MaxLength       =   10
            TabIndex        =   22
            Top             =   0
            Width           =   795
         End
         Begin VB.ComboBox cbcVehGroup 
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
            TabIndex        =   29
            Top             =   840
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.PictureBox plcMajorSort 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   105
            ScaleHeight     =   480
            ScaleWidth      =   4380
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   360
            Width           =   4380
            Begin VB.OptionButton rbcMajorSort 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   5
               Left            =   3120
               TabIndex        =   39
               Top             =   240
               Width           =   1020
            End
            Begin VB.OptionButton rbcMajorSort 
               Caption         =   "Salesperson"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   1560
               TabIndex        =   38
               Top             =   240
               Width           =   1380
            End
            Begin VB.OptionButton rbcMajorSort 
               Caption         =   "Owner"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   600
               TabIndex        =   37
               Top             =   240
               Width           =   1020
            End
            Begin VB.OptionButton rbcMajorSort 
               Caption         =   "NTR Type"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3000
               TabIndex        =   36
               Top             =   0
               Width           =   1335
            End
            Begin VB.OptionButton rbcMajorSort 
               Caption         =   "Agency"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1920
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   1050
            End
            Begin VB.OptionButton rbcMajorSort 
               Caption         =   "Advertiser"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   600
               TabIndex        =   34
               Top             =   0
               Width           =   1380
            End
         End
         Begin VB.Label lacInclude 
            Caption         =   "Include"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   2890
            Width           =   615
         End
         Begin VB.Label lacContract 
            Appearance      =   0  'Flat
            Caption         =   "Contract #"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   56
            Top             =   3480
            Width           =   1020
         End
         Begin VB.Label lacPeriods 
            Appearance      =   0  'Flat
            Caption         =   "# Periods"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3120
            TabIndex        =   25
            Top             =   60
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lacDate2 
            Appearance      =   0  'Flat
            Caption         =   "End"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1800
            TabIndex        =   23
            Top             =   60
            Width           =   495
         End
         Begin VB.Label lacDate1 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   21
            Top             =   60
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
         Height          =   3780
         Left            =   4710
         ScaleHeight     =   3780
         ScaleWidth      =   4245
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4245
         Begin VB.CheckBox ckcAllMinor 
            Caption         =   "All Advertisers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   960
            TabIndex        =   67
            Top             =   0
            Visible         =   0   'False
            Width           =   2985
         End
         Begin VB.ListBox lbcSelection 
            Height          =   3210
            Index           =   6
            ItemData        =   "Rptselnt.frx":0000
            Left            =   120
            List            =   "Rptselnt.frx":0002
            TabIndex        =   66
            Top             =   360
            Visible         =   0   'False
            Width           =   4155
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   5
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   63
            Top             =   360
            Visible         =   0   'False
            Width           =   4160
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   4
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   62
            Top             =   360
            Visible         =   0   'False
            Width           =   4160
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   3
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   61
            Top             =   360
            Visible         =   0   'False
            Width           =   4160
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   2
            Left            =   120
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   60
            Top             =   360
            Visible         =   0   'False
            Width           =   4160
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   1
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   59
            Top             =   360
            Visible         =   0   'False
            Width           =   4160
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   0
            ItemData        =   "Rptselnt.frx":0004
            Left            =   120
            List            =   "Rptselnt.frx":0006
            MultiSelect     =   2  'Extended
            TabIndex        =   58
            Top             =   300
            Width           =   4160
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Advertisers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   0
            Width           =   4065
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   18
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   16
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
         Width           =   1275
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
         Width           =   930
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
Attribute VB_Name = "RptSelNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselnt.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelNT.Frm   NTR Revenue reports
'       4-3-02
'
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSetAllMinor As Integer        '11-14-18
Dim imAllClickedMinor As Integer

Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
Dim smLogUserCode As String
Dim imTerminate  As Integer
Dim smMajorSortCaption As String    '4-2-03
Dim smTotalsByCaption As String
Dim imSortBy As Integer     'for NTR Recap or Billed & Booked:
'                           0=advt, 1=agy, 2=ntr type, 3=owner, 4=slsp, 5=vehicle, 6=bill date
Dim imPrevMinorInx As Integer
Dim imPrevMajorInx As Integer
Dim imSortMajorInx As Integer
Dim imSortMinorInx As Integer
Dim tmItf() As ITF
Dim smITFTag As String

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
    Dim ilListIndex As Integer
    Dim ilSortBy As Integer
    Dim illoop As Integer
    Dim ilRet As Integer
    Dim ilTemp As Integer
    
    ilListIndex = lbcRptType.ListIndex
'    For ilLoop = 0 To 6         'set all the list boxes invisible except for the selected one
'        lbcSelection(ilLoop).Visible = False
'    Next ilLoop
    imSortMajorInx = cbcSet1.ListIndex
    If ilListIndex = CNT_NTRBB Or CNT_MULTIMBB Then
        ilTemp = cbcSet2.ListIndex
        If cbcSet1.ListIndex = ilTemp - 1 Then
            'error, cannot have the same sort parameter defined for 2 sort fields
            MsgBox "Same sort selection as Sort #2; select another"
            cbcSet1.ListIndex = imPrevMajorInx
            Exit Sub
        End If
 
        If cbcSet2.ListIndex = 0 Then
            lbcSelection(0).Visible = False                 'adv
            lbcSelection(1).Visible = False                 'agy
            lbcSelection(2).Visible = False                 'ntr
            lbcSelection(3).Visible = False                 'owner
            lbcSelection(4).Visible = False                 'slsp
            lbcSelection(5).Visible = False                 'vehicle
            lbcSelection(6).Visible = False                 'multimedia
            
            lbcSelection(0).Height = 3270
            lbcSelection(1).Height = 3270
            lbcSelection(2).Height = 3270
            lbcSelection(3).Height = 3270
            lbcSelection(4).Height = 3270
            lbcSelection(5).Height = 3270
            lbcSelection(6).Height = 3270
        End If
        
             'changing the major sort selection, only turn off the previous list box
        If imPrevMajorInx = 0 Then          'advt
            lbcSelection(0).Visible = False
        ElseIf imPrevMajorInx = 1 Then        'agy
            lbcSelection(1).Visible = False
        ElseIf imPrevMajorInx = 2 Then        'NTR or multimedia
            If ilListIndex = CNT_NTRBB Then
                lbcSelection(2).Visible = False         'ntr
            Else
                lbcSelection(6).Visible = False     'multimedia
            End If
        ElseIf imPrevMajorInx = 3 Then          'owner
            lbcSelection(3).Visible = False
        ElseIf imPrevMajorInx = 4 Then          'slsp
            lbcSelection(4).Visible = False
        ElseIf imPrevMajorInx = 5 Then      'vehicle
            lbcSelection(5).Visible = False
        End If
        
        ilSortBy = cbcSet1.ListIndex
        If ilSortBy = 0 Then
            lbcSelection(0).Visible = True
            lbcSelection(0).Move 120, ckcAll.Height + 30
            ckcAll.Caption = "All Advertisers"
            ckcMajorSplit.Visible = False
            imSortBy = 0                    'flag for Advt sort
        ElseIf ilSortBy = 1 Then             'Agency
            ckcAll.Caption = "All Agencies"
            lbcSelection(1).Visible = True
            lbcSelection(1).Move 120, ckcAll.Height + 30
            imSortBy = 1                    'flag for agency sort
            ckcMajorSplit.Visible = False
        ElseIf ilSortBy = 2 Then             'ntr types
            If ilListIndex = CNT_MULTIMBB Then
                ckcAll.Caption = "All Multimedia Types"
                lbcSelection(6).Visible = True
                lbcSelection(6).Move 120, ckcAll.Height + 30
           Else
                ckcAll.Caption = "All NTR Types"
                lbcSelection(2).Visible = True
                lbcSelection(2).Move 120, ckcAll.Height + 30
                ckcMajorSplit.Visible = False
            End If
            imSortBy = 2                    'flag for NTR type sort
        ElseIf ilSortBy = 3 Then            'owner
            ckcAll.Caption = "All Owners"
            lbcSelection(3).Visible = True
            lbcSelection(3).Move 120, ckcAll.Height + 30
            ckcMajorSplit.Visible = False
            imSortBy = 3                    'flag for owners sort
        ElseIf ilSortBy = 4 Then             'salesperson
            ckcAll.Caption = "All Salespeople"
            lbcSelection(4).Visible = True
            lbcSelection(4).Move 120, ckcAll.Height + 30
            imSortBy = 4                    'flag for salespeople sort
            ckcMajorSplit.Visible = True
            ckcMinorSplit.Visible = False
        ElseIf ilSortBy = 5 Then           'vehicle
            ckcAll.Caption = "All Vehicles"
            lbcSelection(5).Visible = True
            lbcSelection(5).Move 120, ckcAll.Height + 30
            imSortBy = 5                    'flag for vehicle sort
            ckcMajorSplit.Visible = False
        End If
        imPrevMajorInx = cbcSet1.ListIndex

        ckcAll.Value = vbUnchecked
    End If
End Sub

Private Sub cbcSet2_Click()
    imSortMinorInx = cbcSet2.ListIndex
    If imSortMinorInx = 0 Then                  'no minor set selected
        lbcSelection(0).Move 120, 285, 4260, 3270   '3-18-16
        lbcSelection(1).Move 120, 285, 4260, 3270
        lbcSelection(2).Move 120, 285, 4260, 3270
        lbcSelection(3).Move 120, 285, 4260, 3270
        lbcSelection(4).Move 120, 285, 4260, 3270
        lbcSelection(5).Move 120, 285, 4260, 3270
        lbcSelection(6).Move 120, 285, 4260, 3270
        
        ckcAllMinor.Visible = False
       
        If imPrevMinorInx = 0 Then
            Exit Sub
        'turn off only the list box that was selected for the minor sort set
        'dont want to turn off the major sort selection
        ElseIf imPrevMinorInx = 1 Then          'advt
            lbcSelection(0).Visible = False
        ElseIf imPrevMinorInx = 2 Then         'agy
            lbcSelection(1).Visible = False
        ElseIf imPrevMinorInx = 3 Then         'ntr/multi
            If lbcRptType.ListIndex = CNT_NTRBB Then
                lbcSelection(2).Visible = False
            Else
                lbcSelection(6).Visible = False     'multimedia
            End If
        ElseIf imPrevMinorInx = 4 Then         'owner
            lbcSelection(3).Visible = False
        ElseIf imPrevMinorInx = 5 Then         'slsp
            lbcSelection(4).Visible = False
        ElseIf imPrevMinorInx = 6 Then         'vehicle
            lbcSelection(5).Visible = False
        End If
        imPrevMinorInx = 0
        imPrevMinorInx = cbcSet2.ListIndex
'                ckcAllMinor.Value = vbUnchecked
        Exit Sub
    End If
        
        
    If imSortMinorInx = imSortMajorInx + 1 Then
        'error, cannot have the same sort parameter defined for 2 sort fields
'                cbcSet1.ListIndex = imPrevMajorInx
        cbcSet2.ListIndex = 0             'default to none
        MsgBox "Same sort selection as Sort #1; select another"
'                Exit Sub
'            Else
'                imPrevMinorInx = imSortMinorInx
    End If

    '*************ok up to here
        
    lbcSelection(0).Height = 1500
    lbcSelection(1).Height = 1500
    lbcSelection(2).Height = 1500
    lbcSelection(3).Height = 1500
    lbcSelection(4).Height = 1500
    lbcSelection(5).Height = 1500
    lbcSelection(6).Height = 1500

    ckcAllMinor.Value = vbUnchecked

    If imPrevMinorInx = 1 Then          'advt
        lbcSelection(0).Visible = False
    ElseIf imPrevMinorInx = 2 Then          'agy
        lbcSelection(1).Visible = False
    ElseIf imPrevMinorInx = 3 Then          'ntr or multimedia
        If lbcRptType.ListIndex = CNT_NTRBB Then
            lbcSelection(2).Visible = False
        Else
            lbcSelection(6).Visible = False
        End If
    ElseIf imPrevMinorInx = 4 Then          'owner
        lbcSelection(3).Visible = False
    ElseIf imPrevMinorInx = 5 Then          'slsp
        lbcSelection(4).Visible = False
    ElseIf imPrevMinorInx = 6 Then          'vehicle
        lbcSelection(5).Visible = False
    End If
    
    
    If imSortMinorInx = 1 Then          'advt
         lbcSelection(0).Move 120, 2220
         lbcSelection(0).Visible = True
         ckcAllMinor.Move 120, 1920
         ckcAllMinor.Caption = "All Advertisers"
         ckcAllMinor.Visible = True
         ckcMinorSplit.Visible = False
         imPrevMinorInx = imSortMinorInx
    ElseIf imSortMinorInx = 2 Then     'agency
         lbcSelection(1).Move 120, 2220
         lbcSelection(1).Visible = True
         ckcAllMinor.Move 120, 1920
         ckcAllMinor.Caption = "All Agencies"
         ckcAllMinor.Visible = True
         imPrevMinorInx = imSortMinorInx
        ckcMinorSplit.Visible = False
     ElseIf imSortMinorInx = 3 Then     'ntr or multimedia
        ckcAllMinor.Move 120, 1920
        If lbcRptType.ListIndex = CNT_NTRBB Then
            lbcSelection(2).Move 120, 2220
            lbcSelection(2).Visible = True
            ckcAllMinor.Caption = "All NTR Types"
           Else
            lbcSelection(6).Move 120, 2220
            lbcSelection(6).Visible = True
            ckcAllMinor.Caption = "All Multimedia"
        End If
        ckcAllMinor.Move 120, 1920
        ckcAllMinor.Visible = True
        ckcMinorSplit.Visible = False
       imPrevMinorInx = imSortMinorInx
     ElseIf imSortMinorInx = 4 Then     'Owner
         lbcSelection(3).Move 120, 2220
         lbcSelection(3).Visible = True
         ckcAllMinor.Move 120, 1920
         ckcAllMinor.Caption = "All Owners"
         ckcAllMinor.Visible = True
         imPrevMinorInx = imSortMinorInx
         ckcMinorSplit.Visible = False
         ckcMajorSplit.Visible = False
    ElseIf imSortMinorInx = 5 Then     'slsp
         lbcSelection(4).Move 120, 2220
         lbcSelection(4).Visible = True
         ckcAllMinor.Move 120, 1920
         ckcAllMinor.Caption = "All Salespeople"
         ckcAllMinor.Visible = True
         imPrevMinorInx = imSortMinorInx
         ckcMinorSplit.Visible = True
    ElseIf imSortMinorInx = 6 Then     'vehicle
         lbcSelection(5).Move 120, 2220
         lbcSelection(5).Visible = True
         ckcAllMinor.Move 120, 1920
         ckcAllMinor.Caption = "All Vehicles"
         ckcAllMinor.Visible = True
         imPrevMinorInx = imSortMinorInx
         ckcMinorSplit.Visible = False
   End If
End Sub

Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Dim ilLbcIndex As Integer

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
    If ilIndex = CNT_NTRRECAP Then
        If imSortBy = 0 Then            'advt
            ilLbcIndex = 0
        ElseIf imSortBy = 2 Then        'ntr types
            ilLbcIndex = 2
        ElseIf imSortBy = 4 Then        'slsp
            ilLbcIndex = 4
        ElseIf imSortBy = 5 Then        'vehicle
            ilLbcIndex = 5
        End If
    ElseIf ilIndex = CNT_NTRBB Or ilIndex = CNT_MULTIMBB Then
        ilLbcIndex = cbcSet1.ListIndex  'sort type for NTR B & B
        If ilIndex = CNT_MULTIMBB And ilLbcIndex = 2 Then
            ilLbcIndex = 6
        End If
    End If
    ilValue = Value
    If imSetAll Then

        imAllClicked = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(ilLbcIndex).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcAllMinor_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Dim ilLbcIndex As Integer

    Value = False
    If ckcAllMinor.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    Dim ilIndex As Integer
    ilIndex = lbcRptType.ListIndex          'report index
    If ilIndex = CNT_NTRBB Or ilIndex = CNT_MULTIMBB Then
        ilLbcIndex = cbcSet2.ListIndex  'sort type for NTR B & B
        If ilLbcIndex = 3 Then          'ntr or multimedia
            If ilIndex = CNT_MULTIMBB Then
                ilLbcIndex = 6
            Else
                ilLbcIndex = 2
            End If
        Else
            ilLbcIndex = ilLbcIndex - 1
        End If
        ilValue = Value
        If imSetAllMinor Then
            imAllClickedMinor = True
            llRg = CLng(lbcSelection(ilLbcIndex).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(ilLbcIndex).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            imAllClickedMinor = False
        End If
        mSetCommands
    End If
End Sub

'TTP 10863 - NTR Billed and Booked: Add Political Selectivity
Private Sub ckcInclNonPolit_Click()
    mSetCommands
End Sub

Private Sub ckcInclPolit_Click()
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

    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenReportNT(ilListIndex) Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gCmcGenNT(ilListIndex, imGenShiftKey, smLogUserCode)
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
        ElseIf ilRet = 0 Then           '0 = invalid input data, stay in
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        ElseIf ilRet = 2 Then           'successful from Bridgereport
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
        gCreateNTR ilListIndex, imSortBy
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
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '5-2-02

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
    Resume Next
'cmcGenErr:
'    ilDDFSet = True
'    Resume Next
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
    'cdcSetup.flags = cdlPDPrintSetup
    'cdcSetup.Action = 5    'DLG_PRINT
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub

Private Sub edcContract_GotFocus()
    gCtrlGotFocus edcContract
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

Private Sub edcDate1_Change()
    mSetCommands
End Sub

Private Sub edcDate1_GotFocus()
    gCtrlGotFocus edcDate1
End Sub

Private Sub edcDate2_Change()
    mSetCommands
End Sub

Private Sub edcDate2_GotFocus()
    gCtrlGotFocus edcDate2
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
    If (KeyAscii <= 32) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KeyDown) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcPeriods_Change()
    mSetCommands
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
    mInit
    If imTerminate = -99 Then
        Exit Sub
    End If
    If imTerminate Then 'Used for print only
        'mTerminate
        cmcCancel_Click
        Exit Sub
    End If
    'RptSelNT.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tgAirNameCode
    Erase tgCSVNameCode
    Erase tgSellNameCode
    Erase tgRptSelSalespersonCode
    Erase tgRptSelAgencyCode
    Erase tgRptSelAdvertiserCode
    Erase tgRptSelNameCode
    PECloseEngine
    Set RptSelNT = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcRptType_Click()
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If ilListIndex = CNT_NTRRECAP Then
        smMajorSortCaption = "Sort by"
        lacDate1.Caption = "Dates- Start"
        lacDate1.Move 120, 30, 1440
        CSI_CalFrom.Move 1140, 0
        CSI_CalFrom.Visible = True
        lacDate2.Caption = "End"
        lacDate2.Move 2440, 30, 1440
        edcDate1.Visible = False
        csi_CalTo.Move 2800, 0
        csi_CalTo.Visible = True
        edcDate2.Visible = False
        rbcMajorSort(0).Caption = "Advertiser"
        rbcMajorSort(1).Caption = "NTR Type"
        rbcMajorSort(2).Caption = "Sched Bill Date"
        rbcMajorSort(3).Caption = "Salesperson"
        rbcMajorSort(4).Caption = "Vehicle"
        rbcMajorSort(5).Visible = False
        plcMajorSort.Move 120, edcDate1.Top + edcDate1.Height + 30
        rbcMajorSort(0).Move 720, 0, 1200
        rbcMajorSort(0).Value = True
        rbcMajorSort(1).Move 2040, 0, 1080
        rbcMajorSort(2).Move 240, 210, 1800
        rbcMajorSort(3).Move 1800, 210, 1360
        rbcMajorSort(4).Move 3280, 210, 1200

        plcTotalsBy.Move 120, plcMajorSort.Top + plcMajorSort.Height

        rbcTotalsBy(0).Caption = "Billed"
        rbcTotalsBy(1).Caption = "Unbilled"
        rbcTotalsBy(2).Caption = "Both"
        smTotalsByCaption = "Include"
        rbcTotalsBy(2).Value = True     'include billed & unbilled
        plcTotalsBy.Visible = True
        ckcSkipPage.Move 120, plcTotalsBy.Top + plcTotalsBy.Height
        ckcSkipPage.Value = vbUnchecked
        ckcSkipPage.Visible = True
        ckcShowDescr.Move 120, ckcSkipPage.Top + ckcSkipPage.Height
        ckcShowDescr.Visible = True
        
        lacInclude.Move 120, ckcShowDescr.Top + ckcShowDescr.Height + 10
        'ckcInclHardCost.Move  120, ckcShowDescr.Top + ckcShowDescr.Height
        ckcInclHardCost.Move 840, ckcShowDescr.Top + ckcShowDescr.Height
        ckcInclHardCost.Visible = True
        
        '1-14-10 Use Acq cost, new option
        ckcUseAcqCost.Move 120, ckcInclHardCost.Top + ckcInclHardCost.Height
        ckcUseAcqCost.Visible = True
        
        lacSet1.Visible = False
        cbcSet1.Visible = False
        lacSet2.Visible = False
        cbcSet2.Visible = False
        plcGrossNet.Visible = False
        plcBillBy.Visible = False
        lacPeriods.Visible = False
        edcPeriods.Visible = False
        lacVehGroup.Visible = False
        cbcVehGroup.Visible = False
        lacContract.Visible = False
        edcContract.Visible = False

    ElseIf ilListIndex = CNT_NTRBB Or ilListIndex = CNT_MULTIMBB Then
        smMajorSortCaption = "Major sort"
        lacDate1.Caption = "Year"
        lacDate1.Move 120, 30, 720
        edcDate1.Move 600, 0, 720
        edcDate1.MaxLength = 4
        
        lacDate2.Caption = "Start Month"
        lacDate2.Move 1560, 30, 1440
        edcDate2.Move 2640, 0, 360
        edcDate2.MaxLength = 2

        lacPeriods.Caption = "# Periods"
        lacPeriods.Move 3240, 30, 1080
        edcPeriods.Move 4080, 0, 360
        edcPeriods.MaxLength = 2
        edcPeriods.Text = "12"
        lacPeriods.Visible = True
        edcPeriods.Visible = True

        lacVehGroup.Move 120, edcDate1.Top + edcDate1.Height + 90
        cbcVehGroup.Move 1560, edcDate1.Top + edcDate1.Height + 30
        gPopVehicleGroups RptSelNT!cbcVehGroup, tgVehicleSets1(), True
        lacVehGroup.Visible = True
        cbcVehGroup.Visible = True

        lacSet1.Move 120, cbcVehGroup.Top + cbcVehGroup.Height + 150
        cbcSet1.AddItem "Advertiser"
        cbcSet1.AddItem "Agency"
        If ilListIndex = CNT_MULTIMBB Then
            cbcSet1.AddItem "Multimedia Type"
        Else
            cbcSet1.AddItem "NTR Type"
        End If
        cbcSet1.AddItem "Owner"
        cbcSet1.AddItem "Salesperson"
        cbcSet1.AddItem "Vehicle"
        cbcSet1.Move 1560, cbcVehGroup.Top + cbcVehGroup.Height + 60
        cbcSet1.ListIndex = 0
        cbcSet1.Visible = True
        lacSet1.Visible = True
        ckcMajorSplit.Move cbcSet1.Left + cbcSet1.Width + 240, cbcSet1.Top
        lacSet2.Move 120, cbcSet1.Top + cbcSet1.Height + 120
        cbcSet2.AddItem "None"
        cbcSet2.AddItem "Advertiser"
        cbcSet2.AddItem "Agency"
        If ilListIndex = CNT_MULTIMBB Then
            cbcSet2.AddItem "Multimedia Types"
        Else
            cbcSet2.AddItem "NTR Type"
        End If
        cbcSet2.AddItem "Owner"
        cbcSet2.AddItem "Salesperson"
        cbcSet2.AddItem "Vehicle"
        cbcSet2.Move 1560, cbcSet1.Top + cbcSet1.Height + 60
        cbcSet2.Visible = True
        lacSet2.Visible = True
        cbcSet2.ListIndex = 0
        ckcMinorSplit.Move ckcMajorSplit.Left, cbcSet2.Top

        plcTotalsBy.Move 120, cbcSet2.Top + cbcSet2.Height + 60
        plcTotalsBy.Visible = True
        plcGrossNet.Move 120, plcTotalsBy.Top + plcTotalsBy.Height
        rbcGrossNet(1).Value = True     'default to net
        plcGrossNet.Visible = True
        plcBillBy.Move 120, plcGrossNet.Top + plcGrossNet.Height
        If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
            rbcBillBy(0).Enabled = False
            rbcBillBy(0).Value = False
            rbcBillBy(1).Value = True
        Else
            rbcBillBy(0).Value = True
        End If
        smTotalsByCaption = "Totals by"
        plcBillBy.Visible = True

        ckcSkipPage.Move 120, plcBillBy.Top + plcBillBy.Height
        ckcSkipPage.Value = vbUnchecked
        ckcSkipPage.Caption = "Skip to new page each major sort"
        ckcSkipPage.Visible = True

        plcMajorSort.Visible = False
        ckcShowDescr.Visible = False

        'TTP 10863 - NTR Billed and Booked: Add Political Selectivity
        lacInclude.Move 120, ckcShowDescr.Top + ckcShowDescr.Height
        'ckcInclHardCost.Move  120, ckcShowDescr.Top + ckcShowDescr.Height
        ckcInclHardCost.Move 840, ckcShowDescr.Top + ckcShowDescr.Height
        ckcInclHardCost.Visible = True
        ckcInclPolit.Move 2160, ckcShowDescr.Top + ckcShowDescr.Height
        ckcInclNonPolit.Move 3000, ckcShowDescr.Top + ckcShowDescr.Height
        If ilListIndex = CNT_NTRBB Then
            ckcInclPolit.Visible = True
            ckcInclNonPolit.Visible = True
        End If
        
        ckcInclHardCost.Visible = True
        If ilListIndex = CNT_MULTIMBB Then          'no hard cost question for multimedia
            ckcInclHardCost.Enabled = False
        End If
        lacContract.Move 120, ckcInclHardCost.Top + ckcInclHardCost.Height + 30
        edcContract.Move 1200, ckcInclHardCost.Top + ckcInclHardCost.Height
        lacContract.Visible = True
        edcContract.Visible = True
    End If
    imSortBy = 0                        'default is by advertiser
    pbcSelC.Visible = True
    pbcOption.Visible = True
    pbcOption.Visible = True
    mSetCommands
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    If Not imAllClicked Then
        If Index = 0 Then           'adv list box
            If cbcSet1.ListIndex = 0 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 1 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
           
        ElseIf Index = 1 Then       'agency list box
           If cbcSet1.ListIndex = 1 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 2 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        ElseIf Index = 2 Then       'NTR  list box
            If cbcSet1.ListIndex = 2 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 3 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        ElseIf Index = 3 Then       'Owner list box
            If cbcSet1.ListIndex = 3 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 4 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        ElseIf Index = 4 Then       'Slsp list box
            If cbcSet1.ListIndex = 4 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 5 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        ElseIf Index = 5 Then       'vehicle list box
            If cbcSet1.ListIndex = 5 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 6 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        ElseIf Index = 6 Then       'multimedia list box
            If cbcSet1.ListIndex = 2 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 3 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        End If
        imSetAll = False
        ckcAll.Value = vbUnchecked  '12-11-01 False
        imSetAll = True
    End If
    If Not imAllClickedMinor Then       '9-24-19
        'imSetAll = False
        'ckcAll.Value = False
        'imSetAll = True
        If Index = 0 Then
            If cbcSet1.ListIndex = 0 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 1 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
           
        ElseIf Index = 1 Then       'agency list box
           If cbcSet1.ListIndex = 1 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 2 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        ElseIf Index = 2 Then       'NTR  list box
            If cbcSet1.ListIndex = 2 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 3 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        ElseIf Index = 3 Then       'Owner list box
            If cbcSet1.ListIndex = 3 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 4 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        ElseIf Index = 4 Then       'Slsp list box
            If cbcSet1.ListIndex = 4 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 5 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        ElseIf Index = 5 Then       'vehicle list box
            If cbcSet1.ListIndex = 5 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 6 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        ElseIf Index = 6 Then       'multimedia list box
            If cbcSet1.ListIndex = 2 Then
                gUncheckAll RptSelNT!ckcAll, imSetAll
            ElseIf cbcSet2.ListIndex = 3 Then
                gUncheckAll RptSelNT!ckcAllMinor, imSetAllMinor
            End If
        End If
        imSetAllMinor = False
        ckcAllMinor.Value = vbUnchecked  '12-11-01 False
        imSetAllMinor = True
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
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*                                                     *
'*******************************************************
Private Sub mInit()
    Dim ilRet As Integer
    Dim slStr As String
    Dim illoop As Integer

    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    'Start Crystal report engine
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        Exit Sub
    End If

    ilRet = gRptAdvtPop(RptSelNT, lbcSelection(0))
    If ilRet = True Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainItf(tmItf(), smITFTag)
    If Not ilRet = True Then
        imTerminate = True
        Exit Sub
    End If
    
    ilRet = gRptAgencyPop(RptSelNT, lbcSelection(1))
    If ilRet = True Then
        imTerminate = True
        Exit Sub
    End If
    
    gPopITFBox RptSelNT, lbcSelection(6), tmItf(), tgItfCode()
    ilRet = gRptMnfPop(RptSelNT, "I", lbcSelection(2), tgMnfCodeCT(), sgMNFCodeTagCT)
    If ilRet = True Then
        imTerminate = True
        Exit Sub
    End If
    
'    ilRet = gPopMnfPlusFieldsBox(RptSelNT, lbcSelection(3), tgSalesperson(), sgSalespersonTag, "H1")
    ilRet = gPopMnfPlusFieldsBox(RptSelNT, lbcSelection(3), tgTmpSort(), sgTmpSortTag, "H1")
    If ilRet = True Then
        imTerminate = True
        Exit Sub
    End If
    
    ilRet = gRptSPersonPop(RptSelNT, lbcSelection(4))
    If ilRet = True Then
        imTerminate = True
        Exit Sub
    End If
    
    ilRet = gRptSellConvVehPop(RptSelNT, lbcSelection(5))
    If ilRet = True Then
        imTerminate = True
        Exit Sub
    End If


    RptSelNT.Caption = smSelectedRptName & " Report"

    slStr = Trim$(smSelectedRptName)
    illoop = InStr(slStr, "&")
    If illoop > 0 Then
        slStr = Left$(slStr, illoop - 1) & "&&" & Mid$(slStr, illoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"

    lbcSelection(0).Move 120, ckcAll.Height + 30, 4065, 3390        'Advertiser
    lbcSelection(1).Move 120, ckcAll.Height + 30, 4065, 3390        'agy
    lbcSelection(2).Move 120, ckcAll.Height + 30, 4065, 3390        'ntr
    lbcSelection(3).Move 120, ckcAll.Height + 30, 4065, 3390        'owner
    lbcSelection(4).Move 120, ckcAll.Height + 30, 4065, 3390        'slsp
    lbcSelection(5).Move 120, ckcAll.Height + 30, 4065, 3390        'vehicle
    lbcSelection(6).Move 120, ckcAll.Height + 30, 4065, 3390        '11-14-18 multimedia
    imAllClicked = False
    imSetAll = True
    imAllClickedMinor = False
    imSetAllMinor = True
'    pbcSelC.Move 90, 255, 4515, 3360
    gCenterStdAlone RptSelNT
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
    gPopExportTypes cbcFileType     '5-2-02

    pbcSelC.Visible = True

    frcOption.Enabled = True
    lbcRptType.AddItem "NTR Recap", 0  'CNT_NTRRECAP
    lbcRptType.AddItem "NTR Billed and Booked", 1  'CNT_NTRBB
    lbcRptType.AddItem "Multimedia Billed and Booked", 2   '1-25-08 CNT_MULTIMBB

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
    'If StrComp(slcommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False  'True 'False
    '    imShowHelpmsg = False
    'Else
     '   igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            End
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
        'imShowHelpmsg = True
        'ilRet = gParseItem(slStr, 3, "^", slHelpSystem)
        'If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
        '    imShowHelpmsg = False
        'End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptSelNT, slStr, ilTestSystem
    'ilRet = gParseItem(slCommand, 3, "\", slStr)
    'igRptCallType = Val(slStr)
    If igRptCallType <> GENERICBUTTON Then
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        igRptType = Val(slStr)
    Else
        igRptType = 0
    End If
    'If igStdAloneMode Then
        'smSelectedRptName = "Copy Inventory by Advertiser"
    '    smSelectedrptName = "Producer Earned Distribution"
   '     igRptCallType = -1 'COLLECTIONSJOB 'INVOICESJOB 'NYFEED  'COLLECTIONSJOB 'SLSPCOMMSJOB   'LOGSJOB 'COPYJOB 'COLLECTIONSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB
    '    igRptType = -1  'unused in standalone exe 'Log     '0   'Summary '3 Program  '1  links
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
    Dim ilEnable As Integer
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    ilEnable = False
    If ilListIndex = CNT_NTRRECAP Then
'        If edcDate1.Text <> "" And edcDate2.Text <> "" Then        'need to put something in for dates
        If CSI_CalFrom.Text <> "" And csi_CalTo.Text <> "" Then        'need to put something in for dates
            If ckcAll.Value = vbChecked Or imSortBy = 6 Then     'check to see at least one item selected in list box for all options
                                                                    'except Sched bill date
                ilEnable = True
            Else
                'need at least one item selected
                If lbcSelection(imSortBy).SelCount > 0 Then
                    ilEnable = True
                End If
            End If
        End If
    Else                'CNT_NTRBB or CNT_MULTIMBB
        If edcDate1.Text <> "" And edcDate2.Text <> "" And edcPeriods.Text <> "" Then
            If lbcSelection(imSortBy).SelCount > 0 Then
                ilEnable = True
            End If
        End If
        If ilListIndex = CNT_NTRBB Then
            'TTP 10863 - NTR Billed and Booked: Add Political Selectivity
            If ckcInclNonPolit.Value = vbUnchecked And ckcInclPolit.Value = vbUnchecked Then
                ilEnable = False
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
    Unload RptSelNT
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcMajorSort_Paint()
    plcMajorSort.Cls
    plcMajorSort.CurrentX = 0
    plcMajorSort.CurrentY = 0
    plcMajorSort.Print smMajorSortCaption
End Sub

Private Sub plcTotalsBy_Paint()
    plcTotalsBy.CurrentX = 0
    plcTotalsBy.CurrentY = 0
    plcTotalsBy.Print smTotalsByCaption    '"Totals By"
End Sub

Private Sub rbcMajorSort_Click(Index As Integer)
    Dim ilListIndex As Integer
    Dim ilSortBy As Integer
    Dim illoop As Integer
    Dim ilRet As Integer

    ilListIndex = lbcRptType.ListIndex
    If ilListIndex = CNT_NTRRECAP Then
        ilSortBy = Index
        ckcAll.Value = vbUnchecked
        ckcAll.Visible = True
        For illoop = 0 To 5         'set all the list boxes invisible
            lbcSelection(illoop).Visible = False
        Next illoop
        If ilSortBy = 0 Then
            lbcSelection(0).Visible = True
            ckcAll.Caption = "All Advertisers"
            lbcSelection(0).Visible = True
            imSortBy = 0                    'flag for Advt sort
        ElseIf ilSortBy = 1 Then            'ntr types
            'NTR types
            ilRet = gRptMnfPop(RptSelNT, "I", lbcSelection(2), tgMnfCodeCT(), sgMNFCodeTagCT)
            If ilRet = True Then
                imTerminate = True
                Exit Sub
            End If
            ckcAll.Caption = "All NTR Types"
            lbcSelection(2).Visible = True
            imSortBy = 2                    'flag for NTR sort
        ElseIf ilSortBy = 2 Then            'sched bill dates
            ckcAll.Visible = False
            imSortBy = 6                    'flag for Bill Date sort
        ElseIf ilSortBy = 3 Then            'salesperson
            ilRet = gRptSPersonPop(RptSelNT, lbcSelection(4))
            If ilRet = True Then
                imTerminate = True
                Exit Sub
            End If
            ckcAll.Caption = "All Salespeople"
            lbcSelection(4).Visible = True
            imSortBy = 4                    'flag for Slsp sort
        ElseIf ilSortBy = 4 Then                           'vehicle
            ilRet = gRptSellConvVehPop(RptSelNT, lbcSelection(5))
            If ilRet = True Then
                imTerminate = True
                Exit Sub
            End If
            ckcAll.Caption = "All Vehicles"
            lbcSelection(5).Visible = True
            imSortBy = 5                    'flag for vehicle sort
        End If

    End If
    mSetCommands
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
        'mInitDDE
        imFirstTime = False
        mInitReport
        If imTerminate Then 'Used for print only
            'mTerminate
            cmcCancel_Click
            Exit Sub
        End If
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub

