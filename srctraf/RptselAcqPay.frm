VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelAcqPay 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acquisition Payment"
   ClientHeight    =   5955
   ClientLeft      =   165
   ClientTop       =   1305
   ClientWidth     =   9270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
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
   ScaleHeight     =   5955
   ScaleWidth      =   9270
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8115
      TabIndex        =   24
      Top             =   90
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6615
      TabIndex        =   17
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
      TabIndex        =   22
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   -60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin VB.Frame frcCopies 
      Caption         =   "Printing"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         BackColor       =   &H00C0C0C0&
         Caption         =   "Printer Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
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
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Acquisition Payables"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4155
      Left            =   15
      TabIndex        =   14
      Top             =   1680
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
         Height          =   3720
         Left            =   120
         ScaleHeight     =   3720
         ScaleWidth      =   4455
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4455
         Begin V81TrafficReports.CSI_Calendar CSI_CalEndDate 
            Height          =   315
            Left            =   2520
            TabIndex        =   53
            Top             =   540
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
         Begin VB.TextBox edcEndDate 
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
            Left            =   3600
            MaxLength       =   9
            TabIndex        =   54
            Text            =   "1"
            Top             =   540
            Visible         =   0   'False
            Width           =   720
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalStartDate 
            Height          =   315
            Left            =   600
            TabIndex        =   52
            Top             =   540
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
         Begin VB.ComboBox cbcMonths 
            Height          =   360
            ItemData        =   "RptSelAcqPay.frx":0000
            Left            =   240
            List            =   "RptSelAcqPay.frx":0028
            TabIndex        =   26
            Top             =   2520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox edcMonths 
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
            Left            =   2760
            MaxLength       =   9
            TabIndex        =   32
            Text            =   "1"
            Top             =   2520
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.PictureBox plcPaidOption 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   4290
            TabIndex        =   34
            Top             =   3000
            Visible         =   0   'False
            Width           =   4290
            Begin VB.CheckBox ckcPaidOption 
               Caption         =   "Fully Paid"
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
               Height          =   270
               Index           =   0
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.CheckBox ckcPaidOption 
               Caption         =   "Not Paid"
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
               Height          =   270
               Index           =   1
               Left            =   1440
               TabIndex        =   37
               Top             =   0
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.CheckBox ckcPaidOption 
               Caption         =   "Unposted"
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
               Height          =   270
               Index           =   2
               Left            =   2640
               TabIndex        =   39
               Top             =   0
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1185
            End
         End
         Begin VB.PictureBox plcSortBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   3195
            TabIndex        =   42
            Top             =   1560
            Width           =   3195
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Vehicle"
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
               Height          =   210
               Index           =   1
               Left            =   2040
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Owner"
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
               Height          =   210
               Index           =   0
               Left            =   840
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.CheckBox ckcInclUnposted 
            Caption         =   "Include Unposted"
            Enabled         =   0   'False
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
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Width           =   1905
         End
         Begin VB.PictureBox plcPaid 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4275
            TabIndex        =   27
            Top             =   0
            Width           =   4275
            Begin VB.OptionButton rbcPaid 
               Caption         =   "Not Paid"
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
               Height          =   210
               Index           =   1
               Left            =   2160
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton rbcPaid 
               Caption         =   "Fully Paid"
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
               Height          =   210
               Index           =   0
               Left            =   840
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.PictureBox plcType 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4275
            TabIndex        =   36
            Top             =   1260
            Width           =   4275
            Begin VB.OptionButton rbcType 
               Caption         =   "Air Time"
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
               Height          =   210
               Index           =   0
               Left            =   840
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton rbcType 
               Caption         =   "Both"
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
               Height          =   210
               Index           =   2
               Left            =   2880
               TabIndex        =   41
               Top             =   0
               Width           =   720
            End
            Begin VB.OptionButton rbcType 
               Caption         =   "NTR"
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
               Height          =   210
               Index           =   1
               Left            =   2040
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.CheckBox ckcSummaryOnly 
            Caption         =   "Summary Only"
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
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   1800
            Width           =   1665
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
            MaxLength       =   9
            TabIndex        =   47
            Top             =   2160
            Width           =   1095
         End
         Begin MSComDlg.CommonDialog cdcSetup 
            Left            =   3960
            Top             =   3240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DefaultExt      =   ".Txt"
            Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
            FilterIndex     =   1
            FontSize        =   0
            MaxFileSize     =   256
         End
         Begin VB.Label lacMonths 
            Appearance      =   0  'Flat
            Caption         =   "# Months"
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
            Height          =   210
            Left            =   1800
            TabIndex        =   31
            Top             =   2640
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lacPaidType 
            Appearance      =   0  'Flat
            Caption         =   "Fully paid dates to include-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Width           =   2610
         End
         Begin VB.Label lacEndDate 
            Appearance      =   0  'Flat
            Caption         =   "End"
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
            Height          =   210
            Left            =   2040
            TabIndex        =   30
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lacContract 
            Appearance      =   0  'Flat
            Caption         =   "Contract #"
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
            Height          =   210
            Left            =   0
            TabIndex        =   46
            Top             =   2190
            Width           =   930
         End
         Begin VB.Label lacStartDate 
            Appearance      =   0  'Flat
            Caption         =   "Start"
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
            Height          =   210
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   600
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
         Height          =   3930
         Left            =   4590
         ScaleHeight     =   3930
         ScaleWidth      =   4455
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CheckBox ckcAllOwners 
            Caption         =   "All Owners"
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
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1920
            Width           =   1545
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Index           =   1
            ItemData        =   "RptSelAcqPay.frx":008E
            Left            =   120
            List            =   "RptSelAcqPay.frx":0095
            MultiSelect     =   2  'Extended
            TabIndex        =   51
            Top             =   2280
            Width           =   4260
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Index           =   0
            ItemData        =   "RptSelAcqPay.frx":009C
            Left            =   120
            List            =   "RptSelAcqPay.frx":009E
            MultiSelect     =   2  'Extended
            TabIndex        =   49
            Top             =   360
            Width           =   4260
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
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
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   90
            Width           =   2025
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6615
      TabIndex        =   18
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   16
      Top             =   105
      Width           =   2805
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   1890
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Save to File"
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
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8850
      Top             =   1035
      Width           =   360
   End
End
Attribute VB_Name = "RptSelAcqPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptSelAcqPay.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelAcqPay.Frm - Network Program Schedule (Radar Worksheet)
'         5-13-03
'
' Release: 5.1
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
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection(0))
Dim imSetAllOwners As Integer 'True=Set list box; False= don't change list box
Dim imAllClickedOwners As Integer  'True=All box clicked (don't call ckcAll within lbcSelection(1))
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
Dim smLogUserCode As String

Dim imTerminate As Integer

Dim imUseCodes() As Integer
Dim imIncludeCodes As Integer

Dim imUseCodesOwners() As Integer
Dim imIncludeCodesOwners As Integer

Dim bmIncludeFullyPaid As Boolean
Dim bmIncludeNotPaid As Boolean
Dim bmIncludeUnposted As Boolean
Dim bmIncludeAirTime As Boolean
Dim bmIncludeNTR As Boolean

Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GRF record length

Dim tmApf As APF
Dim hmApf As Integer
Dim imApfRecLen As Integer        'Acquisition payable record length
Dim tmApfSrchKey3 As APFKEY3
Dim tmApfSrchKey0 As LONGKEY0

Dim tmChf As CHF
Dim hmCHF As Integer
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChfSrchKey1 As CHFKEY1

Dim hmClf As Integer        'Contract line file handle
Dim tmClfSrchKey0 As CLFKEY0 'CLF key record image
Dim tmClfSrchKey1 As CLFKEY1 'CLF key record image
Dim tmClfSrchKey2 As LONGKEY0 'CLF key record image
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image

Dim hmRvf As Integer        'Receivables file handle
Dim tmRvf As RVF
Dim tmRvfSrchKey5 As LONGKEY0
Dim imRvfRecLen As Integer        'Rvf record length

Dim hmSdf As Integer    'Demo Book Name file handle
Dim tmSdf As SDF
Dim tmSdfSrchKey0 As SDFKEY0
Dim tmSdfSrchKey1 As SDFKEY1
Dim tmSdfSrchKey3 As LONGKEY0
Dim tmSdfSrchKey7 As SDFKEY7
Dim imSdfRecLen As Integer        'Sdf record length

Dim hmSmf As Integer    'Demo Book Name file handle
Dim tmSmf As SMF
Dim tmSmfSrchKey2 As LONGKEY0
Dim tmSmfSrchKey5 As LONGKEY0
Dim imSmfRecLen As Integer        'Sdf record length


Dim tmVef As VEF
Dim hmVef As Integer
Dim imVefRecLen As Integer        'VEF record length

Dim tmIihf As IIHF
Dim hmIihf As Integer
Dim imIihfRecLen As Integer        'IIHF record length
Dim tmIihfSrchKey2 As IIHFKEY2      'contr, veh, date

Private Type STATION_INVNO
    lApfCode As Long
    sStationInvNo As String * 20
    iAiredSpotCount As Integer
End Type
Private lmAPFChanged() As STATION_INVNO


'*******************************************************
'*                                                     *
'*      Procedure Name:mConvAirVeh                         *
'*                                                     *
'*             Created:9-3-02       By:D. Hosaka       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*            box with conventional& airing vehicles   *
'*                                                     *
'*******************************************************
Private Sub mConvAirVeh()
    Dim ilRet As Integer
    
    'ilRet = gPopUserVehicleBox(RptSelAcqPay, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHNTR + ACTIVEVEH + VEHONINSERTION, lbcSelection(0), tgCSVNameCode(), sgCSVNameCodeTag)        'lbcCSVNameCode)
    'user wants to see ALL vehicles on the report, not just the OnInsertions
    ilRet = gPopUserVehicleBox(RptSelAcqPay, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHNTR + ACTIVEVEH, lbcSelection(0), tgCSVNameCode(), sgCSVNameCodeTag)         'lbcCSVNameCode)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mConvAirVehERr
        gCPErrorMsg ilRet, "mConvAirVeh (gPopUserVehicleBox: Vehicle)", RptSelAcqPay
        On Error GoTo 0
    End If
    Exit Sub
mConvAirVehERr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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


Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added

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

Private Sub ckcAllOwners_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllOwners.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added

Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long


    ilValue = Value
    If imSetAllOwners Then
        imAllClickedOwners = True
        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClickedOwners = False
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
    Dim ilListIndex As Integer      '8-29-02

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

         If Not mGenReportAcqPay(ilListIndex) Then      '7-1-16 option to run Acq Monitor report with common routines
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = mCmcGenAcqPay(ilListIndex)              '7-1-16 option to run Acq Monitor report with common routines
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

        mCreateAcquisitionPayable ilListIndex

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
            'ilret = gOutputToFile(slFileName, imFTSelectedIndex)
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '2-21-02
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
    'cdcSetup.flags = cdlPDPrintSetup
    'cdcSetup.Action = 5    'DLG_PRINT
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub

Private Sub CSI_CalEndDate_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalEndDate_GotFocus()
    gCtrlGotFocus CSI_CalEndDate
End Sub

Private Sub CSI_CalStartDate_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalStartDate_GotFocus()
    gCtrlGotFocus CSI_CalStartDate
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

Private Sub edcEndDate_Change()
    mSetCommands
End Sub

Private Sub edcEndDate_GotFocus()
    gCtrlGotFocus edcEndDate
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
    If (KeyAscii <= 32) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KEYDOWN) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub Form_Load()
    imTerminate = False
    igGenRpt = False
    mParseCmmdLine
    mInit
    If imTerminate Then 'Used for print only
        mTerminate True
        Exit Sub
    End If
    'RptSelAcqPay.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    PECloseEngine
    Set RptSelAcqPay = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcRptType_Click()
Dim llStdLastInvDate As Long
Dim slDate As String
Dim ilMonth As Integer
Dim ilYear As Integer

        Select Case lbcRptType.ListIndex
            Case ACQ_PAY
            Case ACQ_MONITOR
                plcPaid.Visible = False
                rbcPaid(0).Visible = False
                rbcPaid(1).Visible = False
                lacPaidType.Visible = False
                ckcInclUnposted.Visible = False
'                plcType.Visible = False
'                rbcType(0).Visible = False
'                rbcType(1).Visible = False
'                rbcType(2).Visible = False
'                edcStartDate.Visible = False
                CSI_CalStartDate.Visible = False
                CSI_CalEndDate.Visible = False
                
                'obtain site, and get last month billed.  Default to current month to be billed
                gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llStdLastInvDate
                llStdLastInvDate = llStdLastInvDate
                'get the month & Year of next invoice period
                slDate = gObtainEndStd(Format$(llStdLastInvDate, "m/d/yy"))
                gObtainMonthYear 0, slDate, ilMonth, ilYear
                edcEndDate.Text = Trim$(str(ilYear))
                edcEndDate.Visible = True
                lacStartDate.Move 120, 30, 720
                lacStartDate.Caption = "Month"
                cbcMonths.Move 840, 0
                cbcMonths.ListIndex = 0
                cbcMonths.Visible = True
                lacEndDate.Move 2295, 30, 600
                lacEndDate.Caption = "Year"
                edcEndDate.Move 2895, 0, 600
                edcEndDate.MaxLength = 4
                lacMonths.Move 120, cbcMonths.Top + cbcMonths.Height + 90
                edcMonths.Move 120 + lacMonths.Width + 120, lacMonths.Top - 30
                lacMonths.Visible = True
                edcMonths.Visible = True
                plcPaidOption.Move 120, edcMonths.Top + edcMonths.Height + 60
                plcPaidOption.Visible = True
                ckcPaidOption(0).Visible = True
                ckcPaidOption(1).Visible = True
                ckcPaidOption(2).Visible = True
                plcType.Move 120, plcPaidOption.Top + plcPaidOption.Height + 60
                plcType.Visible = True
                rbcType(0).Value = True
                plcSortBy.Move 120, plcType.Top + plcType.Height + 60
                ckcSummaryOnly.Move 120, plcSortBy.Top + plcSortBy.Height + 30, 2040
                ckcSummaryOnly.Caption = "Discrepancy Only"
                lacContract.Move 120, ckcSummaryOnly.Top + ckcSummaryOnly.Height + 60
                edcContract.Move 120 + lacContract.Width, lacContract.Top - 30
        End Select

End Sub

Private Sub lbcSelection_Click(Index As Integer)

     If Index = 0 Then
         If Not imAllClicked Then
            imSetAll = False
            ckcAll.Value = vbUnchecked  '9-12-02 False
            imSetAll = True
        End If
    Else
        If Not imAllClickedOwners Then
            imSetAllOwners = False
            ckcAllOwners.Value = vbUnchecked  '9-12-02 False
            imSetAllOwners = True
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

'    lbcRptType.AddItem "Acquisition Payable Fees", ACQ_PAY
'    lbcRptType.AddItem "Station Acquisition Monitor", ACQ_MONITOR
'
'    If lbcRptType.ListCount > 0 Then
'        gFindMatch smSelectedRptName, 0, lbcRptType
'        If gLastFound(lbcRptType) < 0 Then
'            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
'            imTerminate = True
'            Exit Sub
'        End If
'        lbcRptType.ListIndex = gLastFound(lbcRptType)
'    End If
'
'    RptSelAcqPay.Caption = smSelectedRptName & " Report"
'    slStr = Trim$(smSelectedRptName)
'    'Handle the apersand in the option box
'    ilLoop = InStr(slStr, "&")
'    If ilLoop > 0 Then
'        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
'    End If
'    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    imAllClickedOwners = False
    imSetAllOwners = True
    'ckcAll.Move 30, 60
    'lbcSelection(0).Move 15, ckcAll.Height + 90, 4380
    'pbcSelC.Move 90, 255, 4515, 3360

    gCenterStdAlone RptSelAcqPay
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
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    
    gPopExportTypes cbcFileType     '2-21-02
    pbcSelC.Visible = False
    lbcRptType.Clear
    lbcRptType.AddItem "Acquisition Payable Fees", ACQ_PAY
    lbcRptType.AddItem "Station Acquisition Monitor", ACQ_MONITOR

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True

    ckcAll.Visible = True
    ckcAll.Value = vbUnchecked                 'default to no vehicles selected
    mConvAirVeh
    smVehGp5CodeTag = ""                         'init to reread
    lbcSelection(1).Clear
    lbcSelection(1).Tag = ""
    ilRet = gPopMnfPlusFieldsBox(RptSelAcqPay, lbcSelection(1), tgMultiCntrCode(), smVehGp5CodeTag, "H1")

    pbcOption.Visible = True
    pbcSelC.Visible = True
    frcOption.Enabled = True


    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
    End If
        
    RptSelAcqPay.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"

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

    slCommand = sgCommandStr
    ilRet = gParseItem(slCommand, 1, "||", smCommand)
    If (ilRet <> CP_MSG_NONE) Then
        smCommand = slCommand
    End If
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False  'True 'False
    '    imShowHelpMsg = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
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
    '    imShowHelpMsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptSelAcqPay, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Remote Invoice Worksheet"
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


    ilEnable = False
    If lbcRptType.ListIndex = ACQ_PAY Then
'        If (edcStartDate.Text <> "" And edcEndDate.Text <> "") Then
        If (CSI_CalStartDate.Text <> "" And CSI_CalEndDate.Text <> "") Then         '8-27-19 use csi cal control vs edit boxes
            ilEnable = True
    
            If ilEnable Then
                ilEnable = False
                If (ckcAll.Value = vbChecked Or lbcSelection(0).SelCount > 0) And (ckcAllOwners.Value = vbChecked Or lbcSelection(1).SelCount > 0) Then
                    ilEnable = True
               End If
            End If
        End If
    Else                    'Acq_Monitor
'        If edcEndDate.Text <> "" And edcMonths.Text <> "" Then
        If edcEndDate.Text <> "" And edcMonths.Text <> "" Then
            ilEnable = True
        End If
        If ilEnable Then
            ilEnable = False
            If (ckcAll.Value = vbChecked Or lbcSelection(0).SelCount > 0) And (ckcAllOwners.Value = vbChecked Or lbcSelection(1).SelCount > 0) Then
                ilEnable = True
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
    Unload RptSelAcqPay
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
    '    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    '    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    '    Traffic!cdcSetup.Action = 6
    End If
End Sub



Private Sub plcPaid_Paint()
    plcPaid.CurrentX = 0
    plcPaid.CurrentY = 0
    plcPaid.Print "Include"
End Sub

Private Sub plcSortBy_Paint()
    plcSortBy.CurrentX = 0
    plcSortBy.CurrentY = 0
    plcSortBy.Print "Sort By"
End Sub
Private Sub plcType_Paint()
    plcType.CurrentX = 0
    plcType.CurrentY = 0
    plcType.Print "Include"
End Sub

Private Sub rbcOutput_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcOutput(Index).Value
    'End of Coded added
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
        mInitReport
        If imTerminate Then 'Used for print only
            mTerminate True
            Exit Sub
        End If
        imFirstTime = False
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub



Private Sub rbcPaid_Click(Index As Integer)
        If Index = 0 Then
            lacPaidType.Caption = "Fully paid dates to include-"
            ckcInclUnposted.Value = vbUnchecked
            ckcInclUnposted.Enabled = False
        Else
            lacPaidType.Caption = "Invoice Dates to include-"
            ckcInclUnposted.Enabled = True
            ckcInclUnposted.Value = vbChecked
        End If
        mSetCommands
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReport                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function mCmcGenAcqPay(ilListIndex) As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slTime As String
    Dim slStr As String
    Dim slSelection As String
    Dim slInclude As String
    Dim slExclude As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilMonthSelected As Integer
    Dim ilHowManyMonths As Integer
    Dim ilLoop As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilYear As Integer
    
        mCmcGenAcqPay = 0
   
        slExclude = ""
        slInclude = ""
       
        If ilListIndex = ACQ_PAY Then           '7-1-16
        
'            slStartDate = RptSelAcqPay!edcStartDate.Text   'Active From Date
            slStartDate = RptSelAcqPay!CSI_CalStartDate.Text   'Active From Date
            If slStartDate = "" Or Not gValidDate(slStartDate) Then
                mReset
                RptSelAcqPay!CSI_CalStartDate.SetFocus
                Exit Function
            End If
            
'            slEndDate = RptSelAcqPay!edcEndDate.Text   'Active From Date
            slEndDate = RptSelAcqPay!CSI_CalEndDate.Text   'Active From Date
            If slEndDate = "" Or Not gValidDate(slEndDate) Then
                mReset
                RptSelAcqPay!CSI_CalEndDate.SetFocus
                Exit Function
            End If
            
'            'send dates requested
'            slSelection = gFormatRequestDatesForCrystal(slStartDate, slEndDate)
'            If Not gSetFormula("DatesRequested", "'" & slSelection & "'") Then
'                mCmcGenAcqPay = -1
'                Exit Function
'            End If
            
            If RptSelAcqPay!rbcPaid(0).Value Then           'fully paid
                If Not gSetFormula("PaidOrUnpaid", "'P'") Then
                    mCmcGenAcqPay = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("PaidOrUnpaid", "'U'") Then      'use invoice dates to filter
                    mCmcGenAcqPay = -1
                    Exit Function
                End If
            End If
    
    
            If RptSelAcqPay!ckcInclUnposted.Value = vbChecked Then           'include unposted
                If Not gSetFormula("IncludeUnposted", "'Y'") Then
                    mCmcGenAcqPay = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("IncludeUnposted", "'N'") Then
                    mCmcGenAcqPay = -1
                    Exit Function
                End If
            End If
    
'            If RptSelAcqPay!rbcSortBy(0).Value = True Then           'sort by owner
'                If Not gSetFormula("SortBy", "'O'") Then
'                    mCmcGenAcqPay = -1
'                    Exit Function
'                End If
'            Else
'                If Not gSetFormula("SortBy", "'V'") Then
'                    mCmcGenAcqPay = -1
'                    Exit Function
'                End If
'            End If
    
            
'            slSelection = gGRFSelectionForCrystal()     'record selection by date and time genned
'            If Not gSetSelection(slSelection) Then
'                mCmcGenAcqPay = -1
'                Exit Function
'            End If
        ElseIf ilListIndex = ACQ_MONITOR Then           '7-1-16
            slStr = Trim$(edcMonths.Text)
            ilRet = gVerifyInt(slStr, 1, 12)        'if month number came back 0, its invalid
            If ilRet = -1 Then
                mReset
                mCmcGenAcqPay = -1
                RptSelAcqPay!edcMonths.SetFocus
                Exit Function
            End If
            
            slStr = edcEndDate.Text                 'entered year
            ilYear = gVerifyYear(slStr)
            If ilYear = 0 Then
                mReset
                mCmcGenAcqPay = -1
                RptSelAcqPay!edcEndDate.SetFocus                 'invalid year
                Exit Function
            End If
            
            ilMonthSelected = cbcMonths.ListIndex
            ilMonthSelected = ilMonthSelected + 1
            ilHowManyMonths = Val(edcMonths.Text)
            'determine end date of the number of months selected
            slStr = edcEndDate.Text   'Year
            slStartDate = Trim$(str(ilMonthSelected)) & "/15/" & Trim$(slStr)
            slStartDate = gObtainStartStd(slStartDate)
            llStartDate = gDateValue(slStartDate)
            slEndDate = gObtainEndStd(slStartDate)
            llEndDate = gDateValue(slEndDate)
            For ilLoop = 1 To ilHowManyMonths - 1
                llEndDate = llEndDate + 1           'start date of next period
                slEndDate = Format$(llEndDate, "m/d/yy")
                slEndDate = gObtainEndStd(slEndDate)        'get end date of new onth
                llEndDate = gDateValue(slEndDate)
            Next ilLoop
            
            slStr = ""
            If ckcPaidOption(0).Value = vbChecked Then
                slStr = "Fully Paid"
            End If
            If ckcPaidOption(1).Value = vbChecked Then
                If slStr = "" Then
                    slStr = "Not Paid"
                Else
                    slStr = slStr & ", Not Paid"
                End If
            End If
            If ckcPaidOption(2).Value = vbChecked Then
                If slStr = "" Then
                    slStr = "Unposted"
                Else
                    slStr = slStr & ", Unposted"
                End If
            End If
            
            If Not gSetFormula("PaidOptions", "'" & slStr & "'") Then
                mCmcGenAcqPay = -1
                Exit Function
            End If
            
            If RptSelAcqPay!ckcSummaryOnly.Value = vbChecked Then           'Discrepancy only
                If Not gSetFormula("DiscrepancyOnly", "'Y'") Then
                    mCmcGenAcqPay = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("DiscrepancyOnly", "'N'") Then
                    mCmcGenAcqPay = -1
                    Exit Function
                End If
            End If
 
        End If
        
        'common formulas to send to Acqpayables & Acq Monitor reports
        'send dates requested
        slSelection = gFormatRequestDatesForCrystal(slStartDate, slEndDate)
        If Not gSetFormula("DatesRequested", "'" & slSelection & "'") Then
            mCmcGenAcqPay = -1
            Exit Function
        End If
        
        If RptSelAcqPay!rbcSortBy(0).Value = True Then           'sort by owner
            If Not gSetFormula("SortBy", "'O'") Then
                mCmcGenAcqPay = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SortBy", "'V'") Then
                mCmcGenAcqPay = -1
                Exit Function
            End If
        End If
       

        slSelection = gGRFSelectionForCrystal()     'record selection by date and time genned
        If Not gSetSelection(slSelection) Then
            mCmcGenAcqPay = -1
            Exit Function
        End If

    mCmcGenAcqPay = 1         'ok
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReset                   *
'*                                                     *
'*             Created:1/31/96       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Reset controls                 *
'*                                                     *
'*******************************************************
Sub mReset()
    igGenRpt = False
    RptSelAcqPay!frcOutput.Enabled = igOutput
    RptSelAcqPay!frcCopies.Enabled = igCopies
    'RptSelAcqPay!frcWhen.Enabled = igWhen
    RptSelAcqPay!frcFile.Enabled = igFile
    RptSelAcqPay!frcOption.Enabled = igOption
    'RptSelAcqPay!frcRptType.Enabled = igReportType
    Beep
End Sub
Public Function mGenReportAcqPay(ilListIndex As Integer) As Integer

        mGenReportAcqPay = True
       If ilListIndex = ACQ_PAY Then
            If ckcSummaryOnly.Value = vbChecked Then                'summary only
                If Not gOpenPrtJob("AcqPayableSum.rpt") Then
                    mGenReportAcqPay = False
                    Exit Function
                End If
            Else
                If Not gOpenPrtJob("AcqPayable.rpt") Then
                    mGenReportAcqPay = False
                    Exit Function
                End If
            End If
        ElseIf ilListIndex = ACQ_MONITOR Then               '7-1-16
            If Not gOpenPrtJob("AcqMonitor.rpt") Then
                mGenReportAcqPay = False
                Exit Function
            End If
        End If
            
    Exit Function
End Function
'
'       Prepass to produce the Acquisition payment fees to stations
'
'       mCreateAcquisitionPayable
'
'           <input>  ilListIndex - report selected
Public Sub mCreateAcquisitionPayable(ilListIndex As Integer)
Dim ilError As Integer
Dim slStartDate As String
Dim slEndDate As String
Dim llStartDate As Long
Dim llEndDate As Long
Dim ilStartDate(0 To 1) As Integer
Dim ilEndDate(0 To 1) As Integer
Dim ilExtLen As Integer
Dim llNoRec As Long
Dim llRecPos As Long
Dim ilRet As Integer
Dim tlDateTypeBuff As POPDATETYPE
Dim tlLongTypeBuff As POPLCODE
Dim ilOffSet As Integer
Dim llSingleCntr As Long
Dim ilVefIndex As Integer
Dim blIncludeUnposted As Boolean
Dim llFullyPaidDate As Long
Dim llTempDate As Long
Dim llStdEndDate As Long
Dim slStdStartDate As String
Dim ilTempStartDate(0 To 1) As Integer
Dim ilLoop As Integer
Dim ilAiredSpotCount As Integer
Dim llStdStartDate As Long
Dim slPostLogSource As String
Dim ilVffIndex As Integer
Dim ilLoopOnInclusions As Integer   '1-3 reference fully paid, not paid, & unposted
Dim ilLoopOnPaidOption As Integer
Dim ilLo As Integer
Dim ilHi As Integer
Dim blPaymentExists As Boolean
Dim llLatestPaymentDate As Long
Dim blInsertGrf As Boolean
Dim llInvoiceDate As Long

            ilError = mOpenAcqPayable(ilListIndex, slStartDate, slEndDate)
            If ilError Then
                Exit Sub            'at least 1 open error
            End If
            
'            slStartDate = RptSelAcqPay!edcStartDate.Text
'            slEndDate = RptSelAcqPay!edcEndDate.Text
            llStartDate = gDateValue(slStartDate)
            llEndDate = gDateValue(slEndDate)
        
            ReDim imUseCodes(0 To 0) As Integer
            ReDim imUseCodesOwners(0 To 0) As Integer
            gObtainCodesForMultipleLists 0, tgCSVNameCode(), imIncludeCodes, imUseCodes(), RptSelAcqPay
            gObtainCodesForMultipleLists 1, tgMultiCntrCode(), imIncludeCodesOwners, imUseCodesOwners(), RptSelAcqPay
            llSingleCntr = Val(RptSelAcqPay!edcContract.Text)
            
            ReDim lmAPFChanged(0 To 0) As STATION_INVNO
            blIncludeUnposted = False
            If RptSelAcqPay!ckcInclUnposted.Value = vbChecked Then
                blIncludeUnposted = True
            End If
            
            ilLo = 1
            ilHi = 2
            If Not bmIncludeFullyPaid Then
                ilLo = 2
            End If
            If ((Not bmIncludeNotPaid) And (Not bmIncludeUnposted)) Then
                ilHi = 1
            End If
            
            For ilLoopOnPaidOption = ilLo To ilHi
                'obtain the fully paid items
                btrExtClear hmApf   'Clear any previous extend operation
                ilExtLen = Len(tmApf)  'Extract operation record size
            
                ilRet = btrGetFirst(hmApf, tmApf, imApfRecLen, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_END_OF_FILE Then
                    llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
                    Call btrExtSetBounds(hmApf, llNoRec, -1, "UC", "Apf", "") '"EG") 'Set extract limits (all records)
            
                    If llSingleCntr > 0 Then
                        tlLongTypeBuff.lCode = llSingleCntr
                        ilOffSet = gFieldOffset("Apf", "ApfCntrNo")
                        ilRet = btrExtAddLogicConst(hmApf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
                    End If
                    
                    'If RptSelAcqPay.rbcPaid(0).Value Then           'use fully paid dates
                    If bmIncludeFullyPaid And ilLoopOnPaidOption = 1 Then                       '6-30-16 use fully paid dates, test new variables
                        If ilListIndex = ACQ_PAY Then
                            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                            ilOffSet = gFieldOffset("Apf", "ApfFullyPaidDate")
                            ilRet = btrExtAddLogicConst(hmApf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                    
                            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                            ilOffSet = gFieldOffset("Apf", "ApfFullyPaidDate")
                            ilRet = btrExtAddLogicConst(hmApf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                        Else                                        '7-19-16 include paid option on Acq Monitor, find the matching invoice dates with a fully paid date other than 1-1-1970
                            gPackDate "1/1/1970", tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                            ilOffSet = gFieldOffset("Apf", "ApfFullyPaidDate")
                            ilRet = btrExtAddLogicConst(hmApf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlDateTypeBuff, 4)
        
                            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                            ilOffSet = gFieldOffset("Apf", "apfInvDate")
                            ilRet = btrExtAddLogicConst(hmApf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                    
                            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                            ilOffSet = gFieldOffset("Apf", "ApfInvDate")
                            ilRet = btrExtAddLogicConst(hmApf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                        End If
                    Else
                        If ((bmIncludeNotPaid = True Or bmIncludeUnposted = True) And (ilLoopOnPaidOption = 2)) Then
                            gPackDate "1/1/1970", tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                            ilOffSet = gFieldOffset("Apf", "ApfFullyPaidDate")
                            ilRet = btrExtAddLogicConst(hmApf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlDateTypeBuff, 4)
        
                            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                            ilOffSet = gFieldOffset("Apf", "apfInvDate")
                            ilRet = btrExtAddLogicConst(hmApf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                    
                            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                            ilOffSet = gFieldOffset("Apf", "ApfInvDate")
                            ilRet = btrExtAddLogicConst(hmApf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                        End If
                    End If
            
                    ilRet = btrExtAddField(hmApf, 0, ilExtLen) 'Extract the whole record
                    On Error GoTo mObtainApfErr
                    gBtrvErrorMsg ilRet, "mCreateAcquisitionPayable (btrExtAddField):" & "Apf.Btr", RptSelAcqPay
                    On Error GoTo 0
                    ilRet = btrExtGetNext(hmApf, tmApf, ilExtLen, llRecPos)
                    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                        On Error GoTo mObtainApfErr
                        gBtrvErrorMsg ilRet, "mCreateAcquisitionPayable (btrExtGetNextExt):" & "Apf.Btr", RptSelAcqPay
                        On Error GoTo 0
                        ilExtLen = Len(tmApf)  'Extract operation record size
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hmApf, tmApf, ilExtLen, llRecPos)
                        Loop
                        Do While ilRet = BTRV_ERR_NONE
                            tmGrf.sDateType = ""                'I indicates incomplete posting of cash (not fully paid)
                            gPackDate "1/1/1970", tmGrf.iDate(0), tmGrf.iDate(1)
                            'process record
                            'test for filtering of air time, ntr or both
                            ilVefIndex = gBinarySearchVef(tmApf.iVefCode)
                            ilVffIndex = gBinarySearchVff(tmApf.iVefCode)
    '                        If (((tmApf.lSbfCode = 0) And (RptSelAcqPay!rbcType(1).Value = False)) Or ((tmApf.lSbfCode > 0) And (RptSelAcqPay!rbcType(0).Value = False))) And (ilVefIndex > 0) And (ilVffIndex > 0) Then   'test for air time
                            If (((tmApf.lSbfCode = 0) And (bmIncludeAirTime = True)) Or ((tmApf.lSbfCode > 0) And (bmIncludeNTR = True))) And (ilVefIndex >= 0) And (ilVffIndex >= 0) Then   '6-30-16 test new variables for air time
                                If gFilterLists(tmApf.iVefCode, imIncludeCodes, imUseCodes()) Then
                                    If gFilterLists(tgMVef(ilVefIndex).iOwnerMnfCode, imIncludeCodesOwners, imUseCodesOwners()) Then
                                        slPostLogSource = tgVff(ilVffIndex).sPostLogSource
                                If tmApf.iVefCode = 107 Then
                                ilRet = ilRet
                                End If
                                        gUnpackDateLong tmApf.iFullyPaidDate(0), tmApf.iFullyPaidDate(1), llFullyPaidDate
                                        gUnpackDateLong tmApf.iInvDate(0), tmApf.iInvDate(1), llInvoiceDate
                                        tmChfSrchKey1.lCntrNo = tmApf.lCntrNo
                                        tmChfSrchKey1.iCntRevNo = 32000
                                        tmChfSrchKey1.iPropVer = 32000
                                        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmApf.lCntrNo) And (tmChf.sSchStatus <> "F")
                                            ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                        Loop
                                        gUnpackDateLong tmApf.iInvDate(0), tmApf.iInvDate(1), llStdEndDate  'convert the apf std bdcst end date to a start date, which is stored in iihf
                                        slStdStartDate = gObtainStartStd(Format$(llStdEndDate, "m/d/yy"))
                                        llStdStartDate = gDateValue(slStdStartDate)
                                        ilAiredSpotCount = 0
                                        If tmChf.lCntrNo = tmApf.lCntrNo And tmApf.lSbfCode = 0 Then       'matching contract, get the spot counts from sdf, 6-14-16 ignore spot counts for NTR item
                                            ilAiredSpotCount = mGetAiredCount(tmChf.lCode, tmApf.iVefCode, llStdStartDate, llStdEndDate)
                                        End If
                                        'If tmApf.iAiredSpotCount > 0 Or llFullyPaidDate > gDateValue("1/1/1970") Then          'test aired spot count has value or fully paid date is greater than 1/1/70
                                        If ilAiredSpotCount > 0 Or llFullyPaidDate > gDateValue("1/1/1970") Then          'test aired spot count has value (something posted) or fully paid date is greater than 1/1/70
                                            tmGrf.sBktType = "P"            'posted flag
    
                                            If tmChf.lCntrNo <> tmApf.lCntrNo Then      'no matching active contract
                                                ilRet = 1           'flag any kind of error, just to ignore the record
                                            End If
                                            If (ilRet = BTRV_ERR_NONE) Then     'found the contract to access iihf to see if posted since the aired spot count = 0
                                                '5-4-16 retrieve the station inv# from iihf due to wrong # updated into apf from manualposting bug
    '                                            gUnpackDateLong tmApf.iInvDate(0), tmApf.iInvDate(1), llStdEndDate  'convert the apf std bdcst end date to a start date, which is stored in iihf
    '                                            slStdStartDate = gObtainStartStd(Format$(llStdEndDate, "m/d/yy"))
                                                
                                                'if no station posting, still need to show on report (no IIHF exists) for the wired networks (those are the regular sched spots)
                                                If slPostLogSource = "S" Then       'station invoices posted
                                                    'retrieve iihf to get the Station Inv #; this is workaround for bug where the APF is updated with the incorrect station inv # while manual posting
                                                    gPackDate slStdStartDate, ilTempStartDate(0), ilTempStartDate(1)
                                                    tmIihfSrchKey2.iVefCode = tmApf.iVefCode
                                                    tmIihfSrchKey2.lChfCode = tmChf.lCode
                                                    tmIihfSrchKey2.iInvStartDate(0) = ilTempStartDate(0)
                                                    tmIihfSrchKey2.iInvStartDate(1) = ilTempStartDate(1)
                                                    ilRet = btrGetGreaterOrEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                                    If ilRet = BTRV_ERR_NONE Then               '6-9-16; avoid invalid procedure call in gunpackdatelong with old csiIO rtn, doesnt give the last record; it is garbage
                                                        gUnpackDateLong tmIihf.iInvStartDate(0), tmIihf.iInvStartDate(1), llTempDate
                                                        'If (ilRet = BTRV_ERR_NONE) And (tmIihf.iVefCode = tmApf.iVefCode) And (tmIihf.lChfCode = tmChf.lCode) And (llTempDate >= llStartDate And llTempDate <= llEndDate) Then
                                                        If (ilRet = BTRV_ERR_NONE) And (tmIihf.iVefCode = tmApf.iVefCode) And (tmIihf.lChfCode = tmChf.lCode) Then
                                                            If Trim$(tmApf.sStationInvNo) <> Trim$(tmIihf.sStnInvoiceNo) Or tmApf.iAiredSpotCount <> ilAiredSpotCount Then
                                                                lmAPFChanged(UBound(lmAPFChanged)).lApfCode = tmApf.lCode
                                                                lmAPFChanged(UBound(lmAPFChanged)).iAiredSpotCount = ilAiredSpotCount
            
                                                                lmAPFChanged(UBound(lmAPFChanged)).sStationInvNo = Trim$(tmIihf.sStnInvoiceNo)  '5-5-16 fix bug in manual posting that was setting station inv # incorrectly                                                   ReDim Preserve llAPFChanged(0 To UBound(llAPFChanged) + 1) As Long
                                                                ReDim Preserve lmAPFChanged(0 To UBound(lmAPFChanged) + 1) As STATION_INVNO
                                                            End If
                                                        End If
                                                        
                                                        '7-19-16 Acq_Pay matches on fully paid date , Acq Monitor tests with Invoice Dates
                                                        If (ilRet = BTRV_ERR_NONE And tmIihf.iVefCode = tmApf.iVefCode And tmIihf.lChfCode = tmChf.lCode) And (((llFullyPaidDate >= llStartDate And llFullyPaidDate <= llEndDate) And (ilListIndex = ACQ_PAY)) Or ((llInvoiceDate >= llStartDate And llInvoiceDate <= llEndDate) And (llFullyPaidDate > gDateValue("1/1/1970")) And (ilListIndex = ACQ_MONITOR))) Then
                                                            tmGrf.sGenDesc = tmIihf.sStnInvoiceNo
                                                            'tmGrf.sBktType = "P"            'posted unwired
                                                            tmGrf.sBktType = "F"             'mark fully paid
     '                                                       If RptSelAcqPay.rbcPaid(1).Value = True Then            'not paid option
                                                            If (bmIncludeNotPaid = True) And ilLoopOnPaidOption = 2 Then             '6-30-16, use new varaible to test not paid option
                                                                ilRet = 1
                                                            End If
        '                                                    If Trim$(tmApf.sStationInvNo) <> Trim$(tmIihf.sStnInvoiceNo) Or tmApf.iAiredSpotCount <> ilAiredSpotCount Then
        '                                                        lmAPFChanged(UBound(lmAPFChanged)).lApfCode = tmApf.lCode
        '                                                        lmAPFChanged(UBound(lmAPFChanged)).iAiredSpotCount = ilAiredSpotCount
        '
        '                                                        lmAPFChanged(UBound(lmAPFChanged)).sStationInvNo = Trim$(tmIihf.sStnInvoiceNo)  '5-5-16 fix bug in manual posting that was setting station inv # incorrectly                                                   ReDim Preserve llAPFChanged(0 To UBound(llAPFChanged) + 1) As Long
        '                                                        ReDim Preserve lmAPFChanged(0 To UBound(lmAPFChanged) + 1) As STATION_INVNO
        '                                                    End If
                                                        Else            'no IIHF, not posted yet
                                                            'ilret returned found, but not fully posted
                                                            tmGrf.sGenDesc = ""
     '                                                       If RptSelAcqPay.rbcPaid(1).Value = True Then            'not paid option
                                                             If (bmIncludeNotPaid = True Or bmIncludeUnposted = True) And ilLoopOnPaidOption = 2 Then         '6-30-16 use new varaible to test not paid option
    
                                                                If (ilRet = BTRV_ERR_NONE) And (tmIihf.iVefCode = tmApf.iVefCode) And (tmIihf.lChfCode = tmChf.lCode) Then
                                                                    If bmIncludeNotPaid Then
                                                                        blPaymentExists = mTestInvForPayment(ilListIndex, llLatestPaymentDate)
                                                                        If blPaymentExists Then
                                                                            gPackDateLong llLatestPaymentDate, tmGrf.iDate(0), tmGrf.iDate(1)
                                                                            tmGrf.sDateType = "I"               'incomplete cash posting
                                                                        End If
                                                                            
                                                                        'posted but not fully paid, ok to include
                                                                        tmGrf.sBktType = "P"
                                                                        tmGrf.sGenDesc = tmIihf.sStnInvoiceNo
                                                                    Else
                                                                        ilRet = 1               'flag any kind of error , just to ignore
                                                                    End If
                                                                Else
                                                                    'not posted, not fully paid
                                                                    'If blIncludeUnposted Then
                                                                    If bmIncludeUnposted Then               '7-4-16 use new common variable
                                                                        'iihf doesnt exist, means not posted
                                                                        tmGrf.sBktType = "U"        'unposted
                                                                        tmGrf.sGenDesc = ""         'IIHF Station Inv #
                                                                        ilRet = BTRV_ERR_NONE       'allow to show on report
                                                                    Else
                                                                        ilRet = 1       'flag any kind of error, just to ignore the record
                                                                        
                                                                    End If
                                                                End If
                                                            Else            'no iihf
                                                                ilRet = 1   'exclude not paid or unposted
                                                            End If
                                                        End If
                                                    End If              'ilret = btrv_err_none
                                                Else                    'no station posting (wired), no iihf
                                                    'slPostLogSource <> "S"
                                                    If tmApf.iAiredSpotCount <> ilAiredSpotCount Then
                                                         lmAPFChanged(UBound(lmAPFChanged)).lApfCode = tmApf.lCode
                                                         lmAPFChanged(UBound(lmAPFChanged)).iAiredSpotCount = ilAiredSpotCount
                                                         lmAPFChanged(UBound(lmAPFChanged)).sStationInvNo = ""
                                                         ReDim Preserve lmAPFChanged(0 To UBound(lmAPFChanged) + 1) As STATION_INVNO
                                                    End If
                                                    tmGrf.sBktType = "P"        'posted for wired
                                                    '7-10-16 for acq monitor, match on invoice dates requested and verify its fully paid by a date other than 1/1/1970
                                                    If ((llFullyPaidDate >= llStartDate And llFullyPaidDate <= llEndDate) And (bmIncludeFullyPaid = True And ilLoopOnPaidOption = 1) And (ilListIndex = ACQ_PAY)) Or ((llInvoiceDate >= llStartDate And llInvoiceDate <= llEndDate) And (llFullyPaidDate > gDateValue("1/1/1970")) And (bmIncludeFullyPaid = True And ilLoopOnPaidOption = 1) And (ilListIndex = ACQ_MONITOR)) Then
                                                        tmGrf.sBktType = "F"
                                                    Else
                                                        If bmIncludeNotPaid Then
                                                            tmGrf.sBktType = "P"
                                                            blPaymentExists = mTestInvForPayment(ilListIndex, llLatestPaymentDate)
                                                            If blPaymentExists Then
                                                                gPackDateLong llLatestPaymentDate, tmGrf.iDate(0), tmGrf.iDate(1)
                                                                tmGrf.sDateType = "I"               'incomplete cash posting
                                                            End If
                                                        Else
                                                            ilRet = 1                       'exclude station posted not paid
                                                        End If
                                                    End If
                                                    tmGrf.sGenDesc = ""         'station invoice #
                                                End If                  'No station posting exists, continue to show for the wired networks
                                            End If
    
                                        Else                'aired count is 0 AND hasnt been fully paid
                                            'test IIHF to see if posted, 0 aired spots could indicate 0 spots aired
                                            'need to read the contract header to get its internal code (its not stored in apf)
    '                                        tmChfSrchKey1.lCntrNo = tmApf.lCntrNo
    '                                        tmChfSrchKey1.iCntRevNo = 32000
    '                                        tmChfSrchKey1.iPropVer = 32000
    '                                        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    '                                        Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmApf.lCntrNo) And (tmChf.sSchStatus <> "F")
    '                                            ilRet = btrGetNext(hmCHF, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    '                                        Loop
                                            If tmChf.lCntrNo <> tmApf.lCntrNo Then      'no matching active contract
                                                ilRet = 1           'flag any kind of error, just to ignore the record
                                            End If
                                            If (ilRet = BTRV_ERR_NONE) Then     'found the contract to access iihf to see if posted since the aired spot count = 0
                                                '5-4-16 retrieve the station inv# from iihf due to wrong # updated into apf from manualposting bug
    '                                            gUnpackDateLong tmApf.iInvDate(0), tmApf.iInvDate(1), llStdEndDate  'convert the apf std bdcst end date to a start date, which is stored in iihf
    '                                            slStdStartDate = gObtainStartStd(Format$(llStdEndDate, "m/d/yy"))
     '                                           If RptSelAcqPay.rbcPaid(0).Value = True Then            'fully paid option, all spots posted (all missed) & not fully paid
                                                 If bmIncludeFullyPaid = True And ilLoopOnPaidOption = 1 Then            '6-30-16 test new variable for fully paid option, all spots posted (all missed) & not fully paid
    
                                                    ilRet = 1
                                                 Else
                                                
     '                                           If slPostLogSource = "S" Then           'if station posting, continue to see if posted occurred.  If no station posting, show on report anyway as there is no IIHF for wired networks
                                                    gPackDate slStdStartDate, ilTempStartDate(0), ilTempStartDate(1)
                                                    tmIihfSrchKey2.iVefCode = tmApf.iVefCode
                                                    tmIihfSrchKey2.lChfCode = tmChf.lCode
                                                    tmIihfSrchKey2.iInvStartDate(0) = ilTempStartDate(0)
                                                    tmIihfSrchKey2.iInvStartDate(1) = ilTempStartDate(1)
                                                    ilRet = btrGetGreaterOrEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                                    If ilRet <> BTRV_ERR_NONE Then          '6-9-16 IIHF doesnt exist, init values for test below
                                                        tmIihf.lChfCode = 0
                                                        tmIihf.iVefCode = 0
                                                        llTempDate = gDateValue("1/1/1970")
                                                    Else
                                                        gUnpackDateLong tmIihf.iInvStartDate(0), tmIihf.iInvStartDate(1), llTempDate
                                                    End If
                                                    'ilret returns not found if not posted (no IIHF)
                                                    If (ilRet <> BTRV_ERR_NONE) Or (tmIihf.iVefCode <> tmApf.iVefCode) Or (tmIihf.lChfCode <> tmChf.lCode) Or (llTempDate < llStartDate Or llTempDate > llEndDate) Then
'                                                        If blIncludeUnposted Or slPostLogSource <> "S" Then     'include the unwired (no station posting; they are not posted )
                                                        If bmIncludeUnposted Or slPostLogSource <> "S" Then     '7-4-16 use new common variable; include the unwired (no station posting; they are not posted )
                                                            'iihf doesnt exist, means not posted
                                                            If tmApf.iAiredSpotCount <> ilAiredSpotCount Then
                                                                 lmAPFChanged(UBound(lmAPFChanged)).lApfCode = tmApf.lCode
                                                                 lmAPFChanged(UBound(lmAPFChanged)).iAiredSpotCount = ilAiredSpotCount
                                                                 lmAPFChanged(UBound(lmAPFChanged)).sStationInvNo = ""
                                                                 ReDim Preserve lmAPFChanged(0 To UBound(lmAPFChanged) + 1) As STATION_INVNO
                                                            End If
                                                            tmGrf.sBktType = "U"        'unposted
                                                            If slPostLogSource <> "S" Then                   'wired, consider it posted
                                                                tmGrf.sBktType = "P"
                                                                blPaymentExists = mTestInvForPayment(ilListIndex, llLatestPaymentDate)
                                                                If blPaymentExists Then
                                                                    gPackDateLong llLatestPaymentDate, tmGrf.iDate(0), tmGrf.iDate(1)
                                                                    tmGrf.sDateType = "I"               'incomplete cash posting
                    
                                                                End If
                                                                If Not bmIncludeNotPaid Then
                                                                    ilRet = 1                   'flag any kind of error to ignore
                                                                End If
                                                            End If
                                                            tmGrf.sGenDesc = ""         'IIHF Station Inv #
                                                            ilRet = BTRV_ERR_NONE       'allow to show on report
                                                        Else
                                                            ilRet = 1       'flag any kind of error, just to ignore the record
                                                            
                                                        End If
                                                    Else
                                                        If bmIncludeNotPaid = True And ilLoopOnPaidOption = 2 Then
                                                             tmGrf.sBktType = "P"        'posted, nothing aired
                                                             tmGrf.sGenDesc = tmIihf.sStnInvoiceNo       'workaround because wrong station inv stored in apf due to manual posting
                                                             blPaymentExists = mTestInvForPayment(ilListIndex, llLatestPaymentDate)
                                                             If blPaymentExists Then
                                                                 gPackDateLong llLatestPaymentDate, tmGrf.iDate(0), tmGrf.iDate(1)
                                                                 tmGrf.sDateType = "I"               'incomplete cash posting
                                                             End If
                                                             If Trim$(tmApf.sStationInvNo) <> Trim$(tmIihf.sStnInvoiceNo) Or tmApf.iAiredSpotCount <> ilAiredSpotCount Then
                                                                 lmAPFChanged(UBound(lmAPFChanged)).lApfCode = tmApf.lCode
                                                                 lmAPFChanged(UBound(lmAPFChanged)).iAiredSpotCount = ilAiredSpotCount
                                                                 lmAPFChanged(UBound(lmAPFChanged)).sStationInvNo = Trim$(tmIihf.sStnInvoiceNo)  '5-5-16 fix bug in manual posting that was setting station inv # incorrectly                                                   ReDim Preserve llAPFChanged(0 To UBound(llAPFChanged) + 1) As Long
                                                                 ReDim Preserve lmAPFChanged(0 To UBound(lmAPFChanged) + 1) As STATION_INVNO
                                                            End If
                                                        Else
                                                            ilRet = 1                   'flag any error to ignore this invoice
                                                        End If
                                                    
                                                    End If
                                                    
                                                  End If
    '                                            End If                      'continue to show on report for wired and unwired unposted option
                                            End If
                                        End If
                                        
                                        'write record if no errors , ilret set to 1 (any value other than 0) to ignore when it didnt pass some filters.
                                        'Write record if not discrepancy only, or discrepany only and the aired count doesnt match the ordered count
                                        'Acq Payables report always writes as long as no errors with ilRET
                                        'Acq Monitor has a discrepancy only option, where all invoices are printed or only those with spot count variances
                                        'if including Unposted, include whether discrep or not.  the apf aired and found spot counts will always show 0, but they are discrepant since nothing has been posted yet
                                        blInsertGrf = False
                                        If ilListIndex = ACQ_PAY Then
                                            blInsertGrf = True
                                        Else            'station acq monitor
                                            If ckcSummaryOnly.Value = vbUnchecked Then      'show all
                                                blInsertGrf = True
                                            Else
                                                If ckcSummaryOnly.Value = vbChecked Then      'discrep only
                                                    If tmApf.iOrderSpotCount <> ilAiredSpotCount Then
                                                        blInsertGrf = True
                                                    End If
                                                End If
                                            End If
                                            If bmIncludeUnposted And tmGrf.sBktType = "U" Then      'if this invoice hasnt been posted and type has been selected, include it for discreps  as well.
                                                blInsertGrf = True                                  'the apf aired count and sdf aired count are both 0
                                            End If
                                        End If
                                        If (ilRet = BTRV_ERR_NONE) And blInsertGrf Then
                                            tmGrf.lGenTime = lgNowTime
                                            tmGrf.iGenDate(0) = igNowDate(0)
                                            tmGrf.iGenDate(1) = igNowDate(1)
                                            tmGrf.lCode4 = tmApf.lCode          'reference to get all data from apf
                                            'tmgrf.sgendescr = the IIHF Station Inv #
                                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                        End If
                                    End If
                                End If
                            End If
                            ilRet = btrExtGetNext(hmApf, tmApf, ilExtLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hmApf, tmApf, ilExtLen, llRecPos)
                            Loop
                        Loop
                    End If
                End If
            Next ilLoopOnPaidOption
            
            For ilLoop = 0 To UBound(lmAPFChanged) - 1
                tmApfSrchKey0.lCode = lmAPFChanged(ilLoop).lApfCode
                ilRet = btrGetEqual(hmApf, tmApf, Len(tmApf), tmApfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    tmApf.sStationInvNo = lmAPFChanged(ilLoop).sStationInvNo
' *** 5-31-16 remove updating of aired spot count; only update the station inv # if different
'                    tmApf.iAiredSpotCount = lmAPFChanged(ilLoop).iAiredSpotCount
                    ilRet = btrUpdate(hmApf, tmApf, imApfRecLen)
                End If
            Next ilLoop
            
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmVef)
            ilRet = btrClose(hmApf)
            ilRet = btrClose(hmIihf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmSdf)
            ilRet = btrClose(hmSmf)
            
            btrDestroy hmGrf
            btrDestroy hmVef
            btrDestroy hmApf
            btrDestroy hmIihf
            btrDestroy hmCHF
            btrDestroy hmClf
            btrDestroy hmSdf
            btrDestroy hmSmf
            
            If ilListIndex = ACQ_MONITOR Then
                ilRet = btrClose(hmRvf)
                btrDestroy hmRvf
            End If
            
        Exit Sub
            
mObtainApfErr:
            On Error GoTo 0
            MsgBox "RptSelAcqPay: gObtainApf error", vbCritical + vbOKOnly, "Apf I/O Error"
        
            Exit Sub
End Sub
'
'           mOpenAcqPayable - open all applicables files for Acquisition Payables
'           <input>  ilListIndex = report selected
'           <output> slSTartDate - start date of std bdcst month
'                    slEndDate - end date of std bdcst month
'           <return>  true if some kind of I/o error
'
Public Function mOpenAcqPayable(ilListIndex, slStartDate As String, slEndDate As String) As Integer

Dim ilRet As Integer
Dim slTable As String * 3
Dim ilError As Integer
Dim ilMonthSelected As Integer
Dim slStr As String
Dim slDate As String
Dim ilHowManyMonths As Integer
Dim ilLoop As Integer
Dim llStartDate As Long
Dim llEndDate As Long

    ilError = False
    On Error GoTo mOpenAcqPayableErr

    slTable = "Grf"
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imGrfRecLen = Len(tmGrf)

    slTable = "Vef"
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imVefRecLen = Len(tmVef)

    slTable = "Apf"
    hmApf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmApf, "", sgDBPath & "Apf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imApfRecLen = Len(tmApf)

    slTable = "Iihf"
    hmIihf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmIihf, "", sgDBPath & "Iihf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imIihfRecLen = Len(tmIihf)

    slTable = "Chf"
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCHFRecLen = Len(tmChf)

    slTable = "Clf"
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imClfRecLen = Len(tmClf)

    slTable = "Sdf"
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSdfRecLen = Len(tmSdf)
    
    slTable = "Smf"
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSmfRecLen = Len(tmSmf)
    
    If ilListIndex = ACQ_MONITOR Then           'need to access receivables for partial payments
        slTable = "Rvf"
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        imRvfRecLen = Len(tmRvf)
        
    End If
    
    bmIncludeAirTime = False
    bmIncludeNTR = False


    If ilListIndex = ACQ_PAY Then           '7-1-16
        '6-30-16
        bmIncludeFullyPaid = False
        bmIncludeNotPaid = False
        bmIncludeUnposted = False
        If rbcPaid(0).Value Then
            bmIncludeFullyPaid = True
        Else
            If rbcPaid(1).Value Then
                bmIncludeNotPaid = True
            End If
            If ckcInclUnposted.Value = vbChecked Then
                bmIncludeUnposted = True
            End If
        End If
        
        bmIncludeAirTime = False
        bmIncludeNTR = False
        If rbcType(0).Value = True Or rbcType(2).Value = True Then
            bmIncludeAirTime = True
        End If
        If rbcType(1).Value = True Or rbcType(2).Value = True Then
            bmIncludeNTR = True
        End If
        
'        slStartDate = RptSelAcqPay!edcStartDate.Text
'        slEndDate = RptSelAcqPay!edcEndDate.Text
        slStartDate = RptSelAcqPay!CSI_CalStartDate.Text        '8-27-19 use csi cal control vs edit boxes
        slEndDate = RptSelAcqPay!CSI_CalEndDate.Text

        llStartDate = gDateValue(slStartDate)
        slStartDate = Format$(llStartDate, "m/d/yy")
        llEndDate = gDateValue(slEndDate)
        slEndDate = Format$(llEndDate, "m/d/yy")
    End If
    
    If ilListIndex = ACQ_MONITOR Then               '7-1-16
        '7-1-16
        bmIncludeFullyPaid = False
        bmIncludeNotPaid = False
        bmIncludeUnposted = False
        If ckcPaidOption(0).Value Then
            bmIncludeFullyPaid = True
        End If
        If ckcPaidOption(1).Value Then
            bmIncludeNotPaid = True
        End If
        If ckcPaidOption(2).Value = vbChecked Then
            bmIncludeUnposted = True
        End If
        
        '8/9/16 add option for a/t, ntr or both vs always including everything
'        bmIncludeAirTime = True
'        bmIncludeNTR = True
        If rbcType(0).Value = True Or rbcType(2).Value = True Then
            bmIncludeAirTime = True
        End If
        If rbcType(1).Value = True Or rbcType(2).Value = True Then
            bmIncludeNTR = True
        End If
        ilMonthSelected = cbcMonths.ListIndex
        ilMonthSelected = ilMonthSelected + 1
        ilHowManyMonths = Val(edcMonths.Text)
        'determine end date of the number of months selected
        slStr = edcEndDate.Text   'Year
        slStartDate = Trim$(str(ilMonthSelected)) & "/15/" & Trim$(slStr)
        slStartDate = gObtainStartStd(slStartDate)
        llStartDate = gDateValue(slStartDate)
        slEndDate = gObtainEndStd(slStartDate)
        llEndDate = gDateValue(slEndDate)
        For ilLoop = 1 To ilHowManyMonths - 1
            llEndDate = llEndDate + 1           'start date of next period
            slEndDate = Format$(llEndDate, "m/d/yy")
            slEndDate = gObtainEndStd(slEndDate)        'get end date of new onth
            llEndDate = gDateValue(slEndDate)
        Next ilLoop
        
    End If
    

    If ilError Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmApf)
        ilRet = btrClose(hmIihf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        
        btrDestroy hmGrf
        btrDestroy hmVef
        btrDestroy hmApf
        btrDestroy hmIihf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmSdf
        btrDestroy hmSmf
       
        Screen.MousePointer = vbDefault

    End If
    mOpenAcqPayable = ilError
    Exit Function

mOpenAcqPayableErr:
    ilError = True
    gBtrvErrorMsg ilRet, "mOpenAcqPayable (OpenError) #" & str(ilRet) & ": " & slTable, RptSelAcqPay
    Resume Next

End Function
Private Function mIncludeSpot(llMonthStart As Long, llMonthEnd As Long) As Boolean
    Dim blIncludeSpot As Boolean
    Dim ilRet As Integer
    Dim llMissedDate As Long
    
    blIncludeSpot = True
    If tmSdf.sSpotType = "X" Then
        blIncludeSpot = False
    ElseIf tmSdf.sSchStatus = "M" Then
        blIncludeSpot = False
    Else
        If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
            If tmSdf.lSmfCode > 0 Then
                tmSmfSrchKey2.lCode = tmSdf.lCode
                ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llMissedDate
                    If (llMissedDate < llMonthStart) Or (llMissedDate > llMonthEnd) Then
                        blIncludeSpot = False
                    End If
                Else
                    blIncludeSpot = False
                End If
            Else
                blIncludeSpot = False
            End If
        End If
    End If
    mIncludeSpot = blIncludeSpot
End Function

Private Function mGetAiredCount(llChfCode As Long, ilVefCode As Integer, llInvStartdate As Long, llInvEndDate As Long) As Integer
    Dim ilRet As Integer
    Dim llSdfDate As Long
    Dim ilAiredSpotCount As Integer
    
    'Required Files be opened: chf, clf, sdf, smf
    'Find lines
    ilAiredSpotCount = 0
    tmClfSrchKey1.lChfCode = llChfCode
    tmClfSrchKey1.iVefCode = ilVefCode
    ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iVefCode = ilVefCode)
        If tmApf.lAcquisitionCost = tmClf.lAcquisitionCost Then
            'Update aired count
            tmSdfSrchKey0.iVefCode = ilVefCode
            tmSdfSrchKey0.lChfCode = llChfCode
            tmSdfSrchKey0.iLineNo = tmClf.iLine
            tmSdfSrchKey0.lFsfCode = 0
            gPackDateLong llInvStartdate, tmSdfSrchKey0.iDate(0), tmSdfSrchKey0.iDate(1)
            tmSdfSrchKey0.sSchStatus = ""
            gPackTime "12AM", tmSdfSrchKey0.iTime(0), tmSdfSrchKey0.iTime(1)
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.lChfCode = llChfCode) And (tmSdf.iLineNo = tmClf.iLine)
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
                If llSdfDate > llInvEndDate Then
                    Exit Do
                End If
                If mIncludeSpot(llInvStartdate, llInvEndDate) Then
                    ilAiredSpotCount = ilAiredSpotCount + 1
                End If
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mGetAiredCount = ilAiredSpotCount

End Function
'
'               mTestInvForPayment - for posted invoices, see if its a partial payment
'               Search RVF for partial payment; Do not need to search PHF otherwise
'               the fully paid date would be set
'               <input> ilListIndex = ACQ_PAY or ACQ_MONITOR
'               <output> lllatestPaymentDate = 1/1/1970 or if partial payment, the latest payment made
'               <return> true if partial payment exists
'
Public Function mTestInvForPayment(ilListIndex As Integer, llLatestPaymentDate As Long) As Boolean
Dim tlLongTypeBuff As POPLCODE
Dim tlIntTypeBuff As POPICODE
Dim tlStrTypeBuff As SORTCODE
Dim ilExtLen As Integer
Dim llNoRec As Integer
Dim ilRet As Integer
Dim ilOffSet As Integer
Dim llRecPos As Long
Dim llDate As Long

                mTestInvForPayment = False
                If ilListIndex = ACQ_PAY Then           'acq payables report doesnt test for partial payments
                    Exit Function
                End If
                
                llLatestPaymentDate = gDateValue("1/1/1970")
                btrExtClear hmRvf   'Clear any previous extend operation
                ilExtLen = Len(tmRvf)  'Extract operation record size
            
                ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY5, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_END_OF_FILE Then
                    llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
                    Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "rvf", "") 'Set extract limits (all records)
                    tlLongTypeBuff.lCode = tmApf.lInvNo
                    ilOffSet = gFieldOffset("Rvf", "RvfInvNo")
                    ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
                    
                    tlIntTypeBuff.iCode = tmApf.iVefCode
                    ilOffSet = gFieldOffset("Rvf", "rvfAirVefCode")
                    ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                    
                    tlStrTypeBuff.sKey = "PI"    'Extract all matching records
                    ilOffSet = gFieldOffset("Rvf", "RvfTranType")
                    ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_STRING, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlStrTypeBuff, 2)
    
                End If
                
                ilRet = btrExtAddField(hmRvf, 0, ilExtLen) 'Extract the whole record
                On Error GoTo mTestInvForPaymentfErr
                gBtrvErrorMsg ilRet, "mCreateAcquisitionPayable (btrExtAddField):" & "Rvf.Btr", RptSelAcqPay
                On Error GoTo 0
                ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                    On Error GoTo mTestInvForPaymentfErr
                    gBtrvErrorMsg ilRet, "mTestInvForPayment (btrExtGetNextExt):" & "Rvf.Btr", RptSelAcqPay
                    On Error GoTo 0
                    ilExtLen = Len(tmRvf)  'Extract operation record size
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Loop
                    Do While ilRet = BTRV_ERR_NONE
                        If tmRvf.lInvNo <> tmApf.lInvNo Or tmRvf.iAirVefCode <> tmApf.iVefCode Or tmRvf.sTranType <> "PI" Then      'RVF transaction must match the inv #, vehicle code & must be a payment PI
                            Exit Do
                        End If
                        
                        'at least one payment found, mark partially paid
                        mTestInvForPayment = True
                        gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
                        If llDate > llLatestPaymentDate Then
                            llLatestPaymentDate = llDate
                        End If
                        
                        ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                        Loop
                    Loop
                End If
                
                Exit Function

mTestInvForPaymentfErr:
            On Error GoTo 0
            MsgBox "RptSelAcqPay: mTestInvForPayment error ", vbCritical + vbOKOnly, "RVF I/O Error"
        
            Exit Function
End Function
