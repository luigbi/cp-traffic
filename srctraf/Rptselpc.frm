VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelPC 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avails Pressure Selection"
   ClientHeight    =   5535
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
   ScaleHeight     =   5535
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
      TabIndex        =   24
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
      TabIndex        =   25
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4020
      Top             =   4875
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
      Caption         =   "Avails Pressure Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3960
      Left            =   45
      TabIndex        =   14
      Top             =   1515
      Width           =   9165
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
         Height          =   3660
         Left            =   15
         ScaleHeight     =   3660
         ScaleWidth      =   4500
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   4500
         Begin VB.PictureBox plcProposal 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   165
            ScaleHeight     =   420
            ScaleWidth      =   4230
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1395
            Width           =   4230
            Begin VB.CheckBox ckcPropType 
               Caption         =   "Working"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   735
               TabIndex        =   66
               Top             =   0
               Width           =   1410
            End
            Begin VB.CheckBox ckcPropType 
               Caption         =   "Complete"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1905
               TabIndex        =   65
               Top             =   -15
               Value           =   1  'Checked
               Width           =   1425
            End
            Begin VB.CheckBox ckcPropType 
               Caption         =   "Unapproved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   735
               TabIndex        =   64
               Top             =   225
               Value           =   1  'Checked
               Width           =   1815
            End
         End
         Begin VB.PictureBox plcSelC6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   3195
            Width           =   4380
            Begin VB.OptionButton rbcSelC6 
               Caption         =   "Exclude"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3165
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   0
               Width           =   1035
            End
            Begin VB.OptionButton rbcSelC6 
               Caption         =   "Hide"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   840
               TabIndex        =   57
               Top             =   0
               Value           =   -1  'True
               Width           =   750
            End
            Begin VB.OptionButton rbcSelC6 
               Caption         =   "Show separately"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1425
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   0
               Width           =   1740
            End
         End
         Begin VB.PictureBox plcSelC5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   105
            ScaleHeight     =   240
            ScaleWidth      =   4140
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   2955
            Width           =   4140
            Begin VB.OptionButton rbcSelC5 
               Caption         =   "Summary"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1515
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   0
               Width           =   1200
            End
            Begin VB.OptionButton rbcSelC5 
               Caption         =   "Detail"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   600
               TabIndex        =   53
               Top             =   0
               Value           =   -1  'True
               Width           =   915
            End
         End
         Begin VB.PictureBox plcSelC4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4140
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   2475
            Width           =   4140
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Standard Quarter"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   600
               TabIndex        =   51
               Top             =   0
               Value           =   -1  'True
               Width           =   1755
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Start Date"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2400
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   0
               Width           =   1695
            End
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
            Left            =   165
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   "Major Set #"
            Top             =   2220
            Visible         =   0   'False
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
            Left            =   1680
            TabIndex        =   47
            Top             =   2190
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.PictureBox plcSelC1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   4290
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   390
            Width           =   4290
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Feed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   13
               Left            =   3000
               TabIndex        =   68
               Top             =   0
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Fill"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   12
               Left            =   3300
               TabIndex        =   41
               Top             =   720
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "N/C"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   11
               Left            =   2460
               TabIndex        =   40
               Top             =   720
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   1560
               TabIndex        =   39
               Top             =   690
               Value           =   1  'Checked
               Width           =   1035
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Trade"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   765
               TabIndex        =   38
               Top             =   705
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   3240
               TabIndex        =   37
               Top             =   495
               Width           =   1065
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   2400
               TabIndex        =   36
               Top             =   495
               Width           =   1050
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Per Inquiry"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   1575
               TabIndex        =   35
               Top             =   465
               Value           =   1  'Checked
               Width           =   1305
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   795
               TabIndex        =   34
               Top             =   480
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Remnant"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   3060
               TabIndex        =   33
               Top             =   225
               Value           =   1  'Checked
               Width           =   1170
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   1890
               TabIndex        =   32
               Top             =   240
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   810
               TabIndex        =   31
               Top             =   240
               Value           =   1  'Checked
               Width           =   1125
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1800
               TabIndex        =   30
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1155
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   825
               TabIndex        =   29
               Top             =   -45
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.PictureBox plcSelC3 
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   135
            ScaleHeight     =   240
            ScaleWidth      =   4290
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   2685
            Width           =   4290
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Show Other Missed Spots Separately"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   0
               Left            =   75
               TabIndex        =   45
               Top             =   45
               Visible         =   0   'False
               Width           =   3540
            End
         End
         Begin VB.PictureBox plcSelC2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   135
            ScaleHeight     =   255
            ScaleWidth      =   4140
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1935
            Width           =   4140
            Begin VB.OptionButton rbcSelC2 
               Caption         =   "30/60"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1860
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   45
               Width           =   1035
            End
            Begin VB.OptionButton rbcSelC2 
               Caption         =   "Units"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   960
               TabIndex        =   42
               Top             =   15
               Value           =   -1  'True
               Width           =   855
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
            Left            =   3630
            MaxLength       =   3
            TabIndex        =   23
            Top             =   0
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox edcSelCFrom 
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
            MaxLength       =   8
            TabIndex        =   21
            Top             =   45
            Width           =   1170
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Active Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   135
            TabIndex        =   26
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "# of Days"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2400
            TabIndex        =   22
            Top             =   75
            Visible         =   0   'False
            Width           =   810
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
         Height          =   3705
         Left            =   4605
         ScaleHeight     =   3705
         ScaleWidth      =   4455
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   150
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CheckBox ckcAllProps 
            Caption         =   "All Proposals"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   195
            TabIndex        =   67
            Top             =   1905
            Width           =   2040
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   2
            Left            =   210
            MultiSelect     =   2  'Extended
            TabIndex        =   62
            Top             =   2235
            Width           =   4245
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   1
            Left            =   2430
            TabIndex        =   61
            Top             =   315
            Width           =   1950
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   0
            Left            =   225
            MultiSelect     =   2  'Extended
            TabIndex        =   60
            Top             =   330
            Width           =   2100
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   315
            TabIndex        =   59
            Top             =   45
            Width           =   2040
         End
         Begin VB.Label laclbcName 
            Appearance      =   0  'Flat
            Caption         =   "Rate Cards"
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
            Left            =   2835
            TabIndex        =   27
            Top             =   30
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
         Top             =   225
         Value           =   -1  'True
         Width           =   1065
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
Attribute VB_Name = "RptSelPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselpc.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelPC.Frm - Avails Pressure report
'                           Avails Clearance report
'
' Release: 4.5 5/20/99   (Summary)
'          4.5 8/10/99   (Detail)
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
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSetAllProps As Integer    'true = set list box; false = dont change list box
Dim imAllPropsClicked As Integer    'true=all box clicked (dont call ckcAllProps with lbcselection)

Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
'Log
Dim imCodes() As Integer
Dim smLogUserCode As String
Dim imTerminate As Integer
'Rate Card
Dim smRateCardTag As String
Dim tmPropCode() As SORTCODE
Dim smPropTagCode As String
Dim imLowLimit As Integer
Dim tmVsf As VSF


'Comment record-Header/Line
Dim hmCxf As Integer        'CXF Handle
'*
'*      Procedure Name:mTestChfAdvtExt                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if Contract is OK to be    *
'*                     viewed by the user              *
'*                                                     *
'*******************************************************
Private Function mTestChfAdvtExt(frm As Form, ilInSlfCode As Integer, tlChfAdvtExt As CHFADVTEXT, hlVsf As Integer, ilCurrent As Integer) As Integer

    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slStartDate As String
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim tlSrchKey As LONGKEY0
    Dim llTodayDate As Long
    Dim ilUser As Integer
    Dim ilRet As Integer
    Dim ilSlf As Integer
    Dim llLkVsfCode As Long
    Dim ilSlfCode As Integer

    llTodayDate = gDateValue(gNow())
    ilSlfCode = ilInSlfCode
    ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
    If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Then
        ilFound = True
    Else
        ilFound = False
        For ilLoop = LBound(tgUrf) To UBound(tgUrf) Step 1
            If (tgUrf(ilLoop).iVefCode = 0) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
    End If
    If Not ilFound Then
        If tlChfAdvtExt.lVefCode > 0 Then
            For ilLoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                If (tgUrf(ilLoop).iVefCode = tlChfAdvtExt.lVefCode) Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
        ElseIf tlChfAdvtExt.lVefCode < 0 Then
            llLkVsfCode = -tlChfAdvtExt.lVefCode
            Do While llLkVsfCode > 0
                tlSrchKey.lCode = llLkVsfCode
                ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                On Error GoTo mTestChfAdvtExtErr
                gBtrvErrorMsg ilRet, "mPopCntrBoxRec (btrGetEqual): Vsf.Btr", frm
                On Error GoTo 0
                For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                    If tmVsf.iFSCode(ilLoop) > 0 Then
                        For ilUser = LBound(tgUrf) To UBound(tgUrf) Step 1
                            If (tgUrf(ilUser).iVefCode = tmVsf.iFSCode(ilLoop)) Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilUser
                        If ilFound Then
                            Exit For
                        End If
                    End If
                Next ilLoop
                If ilFound Then
                    Exit Do
                End If
                llLkVsfCode = tmVsf.lLkVsfCode
            Loop
            'For ilLoop = LBound(lgVehComboCode) To UBound(lgVehComboCode) - 1 Step 1
            '    If lgVehComboCode(ilLoop) = -tlChfExt.lVefCode Then
            '        ilFound = True
            '        Exit For
            '    End If
            'Next ilLoop
        Else    'All vehicles
            ilFound = True
            If igUserByVeh Then 'Test lines as user defined by vehicle after contract added
            End If
        End If
    End If
    If ilFound Then
        'If ilSlfCode > 0 Then
        '    ilFound = False
        '    For ilSlf = LBound(tlChfAdvtExt.iSlfCode) To UBound(tlChfAdvtExt.iSlfCode) Step 1
        '        If tlChfAdvtExt.iSlfCode(ilSlf) <> 0 Then
        '            If ilSlfCode = tlChfAdvtExt.iSlfCode(ilSlf) Then
        '                ilFound = True
        '                Exit For
        '            End If
        '        End If
        '    Next ilSlf
        'End If
        If (ilSlfCode > 0) And (tgUrf(0).iGroupNo > 0) Then
            For ilSlf = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
                If tgMSlf(ilSlf).iCode = ilSlfCode Then
                    If StrComp(tgMSlf(ilSlf).sJobTitle, "S", 1) <> 0 Then
                        ilSlfCode = 0
                    End If
                    Exit For
                End If
            Next ilSlf
        End If
        ilFound = False
        If ilSlfCode > 0 Then
            For ilSlf = LBound(tlChfAdvtExt.iSlfCode) To UBound(tlChfAdvtExt.iSlfCode) Step 1
                If tlChfAdvtExt.iSlfCode(ilSlf) <> 0 Then
                    If ilSlfCode = tlChfAdvtExt.iSlfCode(ilSlf) Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilSlf
        Else
            'If tgUrf(0).iGroupNo > 0 Then
            '    For ilLoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
            '        If tgPopUrf(ilLoop).iCode = tlRUChf.iUrfCode Then
            '            If tgPopUrf(ilLoop).iGroupNo = tgUrf(0).iGroupNo Then
            '                ilFound = True
            '            End If
             '           Exit For
            '        End If
            '    Next ilLoop
            'End If
            If Not ilFound Then
                If tgUrf(0).iGroupNo > 0 Then
                    For ilSlf = LBound(tlChfAdvtExt.iSlfCode) To UBound(tlChfAdvtExt.iSlfCode) Step 1
                        If tlChfAdvtExt.iSlfCode(ilSlf) <> 0 Then
                            For ilLoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
                                If tgPopUrf(ilLoop).iSlfCode = tlChfAdvtExt.iSlfCode(ilSlf) Then
                                    If tgPopUrf(ilLoop).iGroupNo = tgUrf(0).iGroupNo Then
                                        ilFound = True
                                    End If
                                    Exit For
                                End If
                            Next ilLoop
                            If ilFound Then
                                Exit For
                            End If
                        End If
                    Next ilSlf
                Else
                    ilFound = True
                End If
            End If
        End If
    End If
    If ilFound Then
        If ilCurrent = 0 Then   'Current
            gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slDate
            If gDateValue(slDate) < llTodayDate Then
                ilFound = False
            End If
        ElseIf ilCurrent = 2 Then
            If (tlChfAdvtExt.iStartDate(0) <> 0) Or (tlChfAdvtExt.iStartDate(1) <> 0) Or (tlChfAdvtExt.iEndDate(0) <> 0) Or (tlChfAdvtExt.iEndDate(1) <> 0) Then
                gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slDate
                If gDateValue(slStartDate) <= gDateValue(slDate) Then
                    If gDateValue(slDate) < llTodayDate Then
                        ilFound = False
                    End If
                End If
            End If
        End If
    End If
    mTestChfAdvtExt = ilFound
    Exit Function
mTestChfAdvtExtErr:
    mTestChfAdvtExt = False
    Exit Function
End Function




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
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcAllProps_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllProps.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllProps Then
        imAllPropsClicked = True
        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(2).hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllPropsClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcPropType_Click(Index As Integer)
Dim slCntrStatus As String
Dim ilAAS As Integer
Dim ilAASCode As Integer
Dim ilCurrent As Integer
Dim ilState  As Integer
Dim ilShow As Integer
Dim slCntrType As String
Dim ilRet As Integer


    slCntrStatus = ""
    ReDim tmPropCode(0 To 0) As SORTCODE
    If ckcPropType(0).Value = vbChecked Then
        slCntrStatus = "W"
    End If
    If ckcPropType(1).Value = vbChecked Then
        slCntrStatus = slCntrStatus & "C"
    End If
    If ckcPropType(2).Value = vbChecked Then
        slCntrStatus = slCntrStatus & "I"
    End If
    If slCntrStatus = "" Then
        lbcSelection(2).Clear
    Else
        ilAAS = -1  'no advt, agy, slsp selection
        ilAASCode = -1      'include all advt
        ilCurrent = 0
        slCntrType = ""
        ilState = 6 'for WCI only (proposals)
        ilShow = 1      'show product, version

        ilRet = mPopCntrForAASBox(RptSelPC, ilAAS, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcSelection(2), tmPropCode(), smPropTagCode)
    End If
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
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slCode As String

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
        If rbcSelC5(1).Value Then           'summary (like qtrly avails)
            If Not gOpenPrtJob("Pressure.Rpt") Then
                igGenRpt = False
                frcOutput.Enabled = igOutput
                frcCopies.Enabled = igCopies
                frcFile.Enabled = igFile
                frcOption.Enabled = igOption
                Exit Sub
            End If
        Else
            If Not gOpenPrtJob("PresDetl.Rpt") Then      'detail version
                igGenRpt = False
                frcOutput.Enabled = igOutput
                frcCopies.Enabled = igCopies
                frcFile.Enabled = igFile
                frcOption.Enabled = igOption
                Exit Sub
            End If
        End If
        ilRet = gCmcGenPC(imGenShiftKey, smLogUserCode)
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
        'build array of the valid props to process
        ReDim tgProcessProp(0 To 0) As SELECTPROP
        For ilVehicle = 0 To lbcSelection(2).ListCount - 1 Step 1
            If (lbcSelection(2).Selected(ilVehicle)) Then

                slNameCode = tmPropCode(ilVehicle).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "|", slCode)
                tgProcessProp(UBound(tgProcessProp)).lCntrNo = 99999999 - CLng(slCode)

                slNameCode = tmPropCode(ilVehicle).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'proposal code
                tgProcessProp(UBound(tgProcessProp)).lCode = Val(slCode)
                ReDim Preserve tgProcessProp(0 To UBound(tgProcessProp) + 1) As SELECTPROP
            End If
        Next ilVehicle
        gCRAvailsProposal
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
    imGenShiftKey = 0

    Screen.MousePointer = vbHourglass
    gCRAvrClear
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
Private Sub edcSelCFrom_Change()
Dim slDate As String
Dim llDate As Long
Dim ilRet As Integer
Dim ilLen As Integer
    ilLen = Len(edcSelCFrom)
    If ilLen >= 4 Then
        slDate = edcSelCFrom           'retrieve jan thru dec year
        slDate = gObtainStartStd(slDate)
        llDate = gDateValue(slDate)

        'populate Rate Cards and bring in Rcf, Rif, and Rdf
        ilRet = gPopRateCardBox(RptSelPC, llDate, RptSelPC!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
    End If
    mSetCommands
End Sub
Private Sub edcSelCFrom_GotFocus()
    gCtrlGotFocus edcSelCFrom
End Sub
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
    RptSelPC.Refresh
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
    'RptSelPC.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgClfPC
    Erase tgCffPC
    Erase imCodes
    PECloseEngine
    
    Set RptSelPC = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    If Index < 2 Then
        If Not imAllClicked Then
            If Index = 0 Then           'vehicle list box
                imSetAll = False
                ckcAll.Value = vbUnchecked    'False
                imSetAll = True
            End If
        End If
    Else
        If Not imAllPropsClicked Then
            If Index = 2 Then           'vehicle list box
                imSetAllProps = False
                ckcAllProps.Value = vbUnchecked    'False
                imSetAllProps = True
            End If
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
'*            Place focus before populating all lists  *
'*          7-29-04 Implement inclusion/exclusion of

'*******************************************************
Private Sub mInit()
Dim ilRet As Integer
Dim ilLoop As Integer
Dim slStr As String
Dim ilOnly10 As Integer
Dim ilFound As Integer
Dim ilLoop2 As Integer
Dim ilAAS As Integer
Dim ilAASCode As Integer
Dim ilShow As Integer
Dim ilCurrent As Integer
Dim slCntrType As String
Dim slCntrStatus As String
Dim ilState As Integer


    smPropTagCode = ""
    ReDim tmPropCode(0 To 0) As SORTCODE
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

    RptSelPC.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imAllPropsClicked = False
    imSetAll = True
    imSetAllProps = True
    'cbcSel.Move 120, 30
    lacSelCFrom.Move 120, 75
    lacSelCFrom.Caption = "Start Date"  '"Effective Date"
    edcSelCFrom.Move 1020, edcSelCFrom.Top, 1080
    edcSelCFrom.MaxLength = 8
    lacSelCFrom1.Move 2400, lacSelCFrom.Top, 1200
    lacSelCFrom1.Caption = "# Quarters"
    edcSelCFrom1.Move 3360, edcSelCFrom1.Top, 360
    plcSelC1.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 60, 4380, 900
    ckcSelC1(0).Move 660, -30, 840
    ckcSelC1(1).Move 1500, -30, 960
    ckcSelC1(2).Move 660, 195, 1080
    ckcSelC1(3).Move 1800, 195, 1200
    ckcSelC1(4).Move 3000, 195, 1080
    ckcSelC1(5).Move 660, 420, 600
    ckcSelC1(6).Move 1260, 420, 1320
    ckcSelC1(7).Move 2580, 420, 720
    ckcSelC1(8).Move 3300, 420, 900
    ckcSelC1(9).Move 660, 645, 840
    ckcSelC1(10).Move 1500, 645, 960
    ckcSelC1(11).Move 2460, 645, 600
    ckcSelC1(12).Move 3060, 645, 840
    If tgSpf.sSystemType = "R" Then             'allow exclusion of feed spots if station system
        ckcSelC1(13).Move 2460, -30, 840
        ckcSelC1(13).Visible = True
    Else
        ckcSelC1(13).Visible = False
    End If

    'Proposal types
    plcProposal.Move 120, plcSelC1.Top + plcSelC1.Height
    ckcPropType(0).Move 990, -30, 1020
    ckcPropType(1).Move 2130, -30, 1920
    ckcPropType(2).Move 990, 195, 1920

    plcSelC2.Move 120, plcProposal.Top + plcProposal.Height + 30
    rbcSelC2(0).Move 990, 0, 780
    rbcSelC2(1).Move 1770, 0, 1920

    gPopVehicleGroups RptSelPC!cbcSet1, tgVehicleSets1(), True

    ilOnly10 = True
    For ilLoop = LBound(tgMVef) To UBound(tgMVef) Step 1
    If (tgMVef(ilLoop).sState <> "D") And (tgMVef(ilLoop).sType = "S" Or tgMVef(ilLoop).sType = "C" Or tgMVef(ilLoop).sType = "V") Then
        ilFound = gVpfFindIndex(tgMVef(ilLoop).iCode)
        If ilFound < 0 Then
            ilOnly10 = False
        Else
            For ilLoop2 = 0 To 9
                If tgVpf(ilFound).iSLen(ilLoop2) <> 10 And tgVpf(ilFound).iSLen(ilLoop2) <> 0 Then
                    ilOnly10 = False
                    ilLoop = UBound(tgMVef)
                    Exit For
                End If
            Next ilLoop2
        End If
    End If
    Next ilLoop
    If ilOnly10 Then
        rbcSelC2(0).Value = True            'only 10", default to show units
    Else
        rbcSelC2(1).Value = True            'combo of spot lengths, dfault to 30/60
    End If
    cbcSet1.ListIndex = 0
    edcSet1.Move 120, plcSelC2.Top + plcSelC2.Height + 60
    cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 60
    edcSet1.Visible = True
    cbcSet1.Visible = True
    plcSelC4.Move 120, edcSet1.Top + edcSet1.Height
    'plcSelC3.Move 120, plcSelC4.Top + plcSelC4.Height
    plcSelC5.Move 120, plcSelC4.Top + plcSelC4.Height       'detail or summary

    If rbcSelC5(0).Value Then
        rbcSelC5_Click 0
    Else
        rbcSelC5(0).Value = vbChecked
    End If

    ilAAS = -1  'no advt, agy, slsp selection
    ilAASCode = -1      'include all advt
    ilCurrent = 0
    slCntrStatus = "CI"
    slCntrType = ""
    ilState = 6     'get WCI only
    ilShow = 1      'show product, version
    ilRet = mPopCntrForAASBox(RptSelPC, ilAAS, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcSelection(2), tmPropCode(), smPropTagCode)

    ckcSelC3(0).Move 0, 0
    ckcAll.Move 15, 0
    lbcSelection(0).Move 15, ckcAll.Height + 30, 2135, 1500

    lbcSelection(1).Move lbcSelection(0).Width + 120, lbcSelection(0).Top, 2135, 1500
    laclbcName(0).Move lbcSelection(1).Left, 0


    ckcAllProps.Move 15, lbcSelection(0).Top + lbcSelection(0).Height + 30
    lbcSelection(2).Move lbcSelection(0).Left, ckcAllProps.Top + ckcAllProps.Height + 30, 4380, 1500
    pbcSelC.Move 90, 255, 4515, 3360
    gCenterStdAlone RptSelPC
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
    gPopExportTypes cbcFileType     '10-20-01
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    pbcSelC.Visible = False
    'lbcRptType.Clear
    'lbcRptType.AddItem smSelectedRptName

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    mSellConvVirtVehPop 0, False
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True
    lacSelCFrom.Visible = True
    edcSelCFrom.Visible = True
    ckcAll.Visible = True
    'edcSelCFrom.Move 1350, edcSelCFrom.Top, 945
    lacSelCFrom.Visible = True
    edcSelCFrom.Visible = True
    lacSelCFrom1.Visible = True
    edcSelCFrom1.Visible = True
    lbcSelection(0).Visible = True                  'show budget name list box (base budget)
    lbcSelection(1).Visible = True                 'split budgets
    laclbcName(0).Visible = True
    pbcOption.Visible = True
    pbcOption.Enabled = True

    'If lbcRptType.ListCount > 0 Then
    '    gFindMatch smSelectedRptName, 0, lbcRptType
    '    If gLastFound(lbcRptType) < 0 Then
    '        MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
    '        imTerminate = True
    '        Exit Sub
    '    End If
     '   lbcRptType.ListIndex = gLastFound(lbcRptType)
    'End If
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
    'gInitStdAlone RptSelPC, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Avails Pressure"
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
'*      Procedure Name:mSellConvVirtVehPop             *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSellConvVirtVehPop(ilIndex As Integer, ilUselbcVehicle As Integer)
    Dim ilRet As Integer
    If ilUselbcVehicle Then
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        ilRet = gPopUserVehicleBox(RptSelPC, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        ilRet = gPopUserVehicleBox(RptSelPC, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)    'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelPC
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
    Dim ilLoop As Integer

    ilEnable = False
    If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") Then
        ilEnable = True
        'atleast one must be selected
        If ilEnable Then
            ilEnable = False
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1      'budget entry must be selected
                If lbcSelection(0).Selected(ilLoop) Then
                    igBSelectedIndex = ilLoop
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
            If ilEnable Then    'continue checking other selections if OK so far
                ilEnable = False
                'Check rate card selectionn
                For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1      'vehicle entry must be selected
                    If lbcSelection(1).Selected(ilLoop) Then
                        igRCSelectedIndex = ilLoop
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
                'ok to not select any props
                'If ilEnable Then
                '    ilEnable = False
                '    'Check proposal selectionn
                '    For ilLoop = 0 To lbcSelection(2).ListCount - 1 Step 1      'vehicle entry must be selected
                '        If lbcSelection(2).Selected(ilLoop) Then
                '            ilEnable = True
                '            Exit For
                '        End If
                '    Next ilLoop
                'End If
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
    Unload RptSelPC
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcProposal_Paint()
    plcProposal.Cls
    plcProposal.CurrentX = 0
    plcProposal.CurrentY = 0
    plcProposal.Print "Proposals"
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
Private Sub rbcSelC2_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC2(Index).Value
    'End of coded added
    If rbcSelC2(0).Value Then            'Hide reservations
        ckcSelC1(3).Value = vbChecked   'True           'disallow to be selected
        ckcSelC1(3).Enabled = False
    ElseIf rbcSelC2(1).Value Then               'show separately
        ckcSelC1(3).Value = vbChecked   'True
        ckcSelC1(3).Enabled = True
    Else                                 'exclude
        ckcSelC1(3).Value = vbUnchecked 'False
        ckcSelC1(3).Enabled = False
    End If
End Sub
Private Sub rbcSelC5_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC5(Index).Value
    'End of coded added
    If Index = 0 Then           'detail
        If tgUrf(0).iSlfCode = 0 Then     'guide or counterpoint password, allow anyway way answered
            ckcSelC1(3).Enabled = False
            ckcSelC1(3).Value = vbChecked   'True
            'plcSelC6.Caption = "Reserved"
            plcSelC6.Move 120, plcSelC5.Top + plcSelC5.Height
            plcSelC6.Height = 435
            rbcSelC6(0).Caption = "Hide"
            rbcSelC6(0).Move 990, 0, 780
            rbcSelC6(0).Visible = True
            rbcSelC6(0).Value = True            'default to hide reservations
            rbcSelC6(1).Caption = "Show separately"
            rbcSelC6(1).Move 1770, 0, 1920
            rbcSelC6(1).Visible = True
            rbcSelC6(2).Caption = "Exclude"
            rbcSelC6(2).Move 990, 195, 1080
            rbcSelC6(2).Visible = True
            plcSelC6.Visible = True
        Else                                    'its a slsp
            'If tgUrf(0).iSlfCode > 0 Then       'disallow slsp from seeing reserves, force to bury them within sold
                plcSelC6.Visible = False
            'Else
            '    plcSelC6.Visible = True
            'End If
        'Else
            rbcSelC6(0).Value = True        'force to include in sold (Hide)
        End If
    Else
        ckcSelC1(3).Enabled = True
        ckcSelC1(3).Value = vbChecked   'True
        plcSelC6.Visible = False
    End If
End Sub
Private Sub rbcSelC6_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC6(Index).Value
    'End of coded added
    If Index = 2 Then       'exclude reserves, check it off
        ckcSelC1(3).Value = vbUnchecked 'False
    Else                    'hide or show reserves separately
        ckcSelC1(3).Value = vbChecked   'True
    End If
End Sub
Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcSelC6_Paint()
    plcSelC6.CurrentX = 0
    plcSelC6.CurrentY = 0
    plcSelC6.Print "Reserved"
End Sub
Private Sub plcSelC5_Paint()
    plcSelC5.CurrentX = 0
    plcSelC5.CurrentY = 0
    plcSelC5.Print "Show"
End Sub
Private Sub plcSelC4_Paint()
    plcSelC4.CurrentX = 0
    plcSelC4.CurrentY = 0
    plcSelC4.Print "Use"
End Sub
Private Sub plcSelC1_Paint()
    plcSelC1.CurrentX = 0
    plcSelC1.CurrentY = 0
    plcSelC1.Print "Select"
End Sub
Private Sub plcSelC2_Paint()
    plcSelC2.CurrentX = 0
    plcSelC2.CurrentY = 0
    plcSelC2.Print "Counts by"
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPopCntrForAASBox               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     contract number given adfcode,  *
'*                     or agency or user               *
'*
'*     4-24-02 This is a copy of gPopCntrForAASBox which
'*     has been changed to retrieve only proposals whose
'*     active dates are after the requested report period
'*                                                     *
'*******************************************************
Private Function mPopCntrForAASBox(frm As Form, ilAAS As Integer, ilAASCode As Integer, slStatus As String, slCntrType As String, ilCurrent As Integer, ilHOType As Integer, ilShow As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = mPopCntrForAASBox (MainForm, ilAAS, ilAASCode, slStatus, slCntrType, ilCurrent, ilHOType, ilShow, lbcLocal, tlSortCode(), slSortCodeTag)    MainForm (I)- Name of Form to unload if error exist
'       ilAAS(I)=0=Obtain Contracts for Specified Advertiser Code (ilAASCode) and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'                1=Obtain Contracts for Specified Agency Code (ilAASCode) and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'                2=Obtain Contracts for Specified Salesperson Code (ilAASCode) that matches one of the salespersons defined for the contract
'                3=Obtain Contracts for Specified Vehicle Code (ilAASCode)  and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'               -1=No selection by advertiser or agency or salesperson
'                Note if tgUrf(0).iSlfCode not specified, then salesperson test is bypassed
'       ilAASCode(I)- Advertiser or Agency code to obtain contracts for (-1 for all advertiser or agency)
'       slStatus (I)- chfStatus value or blank
'                         W=Working; D=Dead; C=Completed; I=Incomplete; H=Hold; O=Order; G=Unschd Hold; N=Unschd Order
'                         Multiple status can be specified (WDI)
'                         If H or O (ilHOType indicates which H or O to show)
'       slCntrType (I)- chfType value or blank
'                       C=Standard; V=Reservation; T=Remnant; R=DR; Q=PI; S=PSA; M=Promo
'       ilCurrent (I)- 0=Current (Active) (chfDelete <> y); 1=Past and Current (chfDelete <> y); 2=Current(Active) plus all cancel before start (chfDelete <> y); 3=All plus history (any value for chfDelete)
'       ilHOType (I)-  1=H or O only; 2=H or O or G or N (if G or N exists show it over H or O);
'                      3=H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
'                        Note: G or N can't exist at the same time as W or C or I for an order
'                              G or N or W or C or I CntrRev > 0
'                      4=H or O only and No W or C or I or N or G
'                      5=W or I only and No C or N or G (set slStatus = WC)
        'Old way
        '       ilHOType (I)- 1=Order only (ignore revision); 2=Order if no revision exist or revision only if it exist;
        '                      3=Include Revisision and Order; 4=Only Revision of orders
'       ilShow(I)-  0=Only show numbers,
'                   1=Show Number and advertiser  and product version, status start date & # weeks (4-26-02)
'                   2=Show Number, Dates, Product and vehicle
'                   3=Show Number, Advertiser, Dates
'                   4=Show Number, Dates
'                   5=Show Number, Advertiser
'                   6=Show Number, Status, Advertiser, Dates
'       lbcLocal (O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'CHF date/time stamp
    Dim hlChf As Integer        'CHF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlChf As CHF
    Dim llCntrNo As Long
    Dim ilRevNo As Integer
    Dim ilVerNo As Integer
    Dim ilExtRevNo As Integer
    Dim slShow As String
    Dim slName As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilLoop1 As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim llTodayDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim slCode As String    'Sales source code number
    Dim hlVef As Integer        'Vef handle
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer     'Record length
    Dim tlVefSrchKey As INTKEY0
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim hlVsf As Integer        'Vsf handle
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim tlSrchKey As LONGKEY0
    Dim hlAdf As Integer        'Adf handle
    Dim tlAdf As ADF
    Dim ilAdfRecLen As Integer     'Record length
    Dim tlAdfSrchKey As INTKEY0
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim ilOffSet As Integer
    Dim tlVsf As VSF
    Dim llLen As Long
    Dim ilOper As Integer
    Dim slStr As String
    Dim slExtStr As String
    Dim ilSlfCode As Integer
    Dim ilVefCode As Integer
    Dim ilTestCntrNo As Integer
    Dim tlChfAdvtExt As CHFADVTEXT
    Dim slCntrStatus As String
    Dim slHOStatus As String
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim slToday As String
    Dim slAdvt As String * 10   'when showing advt &prod for ilshow = 1, max 20 char for advt & prod
    Dim slProd As String * 10

    If slStatus = "" Then
        slCntrStatus = "WCIDHO"
    Else
        slCntrStatus = slStatus
    End If
    slHOStatus = ""
    If (ilHOType = 1) Or (ilHOType = 4) Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "H"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "O"
        End If
    ElseIf ilHOType = 2 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
    ElseIf ilHOType = 3 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
        If (InStr(1, slCntrStatus, "H", 1) <> 0) Or (InStr(1, slCntrStatus, "O", 1) <> 0) Then
            slHOStatus = slHOStatus & "WCI"
        End If
        If InStr(1, slCntrStatus, "H", 1) = 0 Then
            slHOStatus = slHOStatus & "G"
        End If
        If InStr(1, slCntrStatus, "O", 1) = 0 Then
            slHOStatus = slHOStatus & "N"
        End If
    ElseIf ilHOType = 5 Then
        slHOStatus = ""
        slCntrStatus = "WI"
    End If
    ilPop = True
    llLen = 0
    'ilRet = 0
    'On Error GoTo mPopCntrForAASBoxErr2
    'ilFound = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilFound = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilFound = 0
    End If
    
    slStamp = gFileDateTime(sgDBPath & "Chf.Btr") & Trim$(str$(ilAASCode)) & Trim$(slCntrStatus) & Trim$(slCntrType) & Trim$(str$(ilCurrent)) & Trim$(str$(ilHOType)) & Trim$(str$(ilShow))

    'On Error GoTo mPopCntrForAASBoxErr2
    'ilRet = 0
    'imLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        imLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        imLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                mPopCntrForAASBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    mPopCntrForAASBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        llTodayDate = gDateValue(gNow())
        slToday = Format$(llTodayDate, "m/d/yy")
        'gObtainVehComboList
        hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrOpen):" & "Chf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlChf) 'btrRecordLength(hlChf)  'Get and save record length
        hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrOpen):" & "Vsf.Btr", frm
        On Error GoTo 0
        ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
        hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrOpen):" & "Vef.Btr", frm
        On Error GoTo 0
        ilVefRecLen = Len(tlVef) 'btrRecordLength(hlSlf)  'Get and save record length
        hlAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrOpen):" & "Adf.Btr", frm
        On Error GoTo 0
        hmCxf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrOpen):" & "Adf.Btr", frm
        On Error GoTo 0
        ilAdfRecLen = Len(tlAdf) 'btrRecordLength(hlSlf)  'Get and save record length
        tlAdf.iCode = 0
        ilSortCode = 0
        ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlChfAdvtExt)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlChf) 'Obtain number of records
        btrExtClear hlChf   'Clear any previous extend operation
        If (ilAAS = 0) And (ilAASCode > 0) Then
            tlIntTypeBuff.iType = ilAASCode
            ilRet = btrGetEqual(hlChf, tlChf, ilRecLen, tlIntTypeBuff, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Else
            ilRet = btrGetFirst(hlChf, tlChf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        End If
        If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
            ilRet = btrClose(hmCxf)
            btrDestroy hmCxf
            ilRet = btrClose(hlAdf)
            btrDestroy hlAdf
            ilRet = btrClose(hlVsf)
            btrDestroy hlVsf
            ilRet = btrClose(hlVef)
            btrDestroy hlVef
            ilRet = btrClose(hlChf)
            btrDestroy hlChf
            Exit Function
        Else
            On Error GoTo mPopCntrForAASBoxErr
            gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrGetFirst):" & "Chf.Btr", frm
            On Error GoTo 0
        End If
        Call btrExtSetBounds(hlChf, llNoRec, -1, "UC", "CHFADVTEXTPK", CHFADVTEXTPK) 'Set extract limits (all records)
        ilSlfCode = tgUrf(0).iSlfCode
        If (tgUrf(0).iGroupNo > 0) Then 'And (tgUrf(0).iSlfCode <= 0) Then
            ilRet = gObtainUrf()
            ilRet = gObtainSalesperson()
        End If



        gPackDate slToday, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Chf", "ChfStartDate")
        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_OR, tlDateTypeBuff, 4)

        ilOffSet = gFieldOffset("Chf", "ChfOHDDate")
        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

        If ilAAS = 1 Then
            If ilAASCode > 0 Then
                tlIntTypeBuff.iType = ilAASCode
                ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
                If (slCntrType = "") And (ilCurrent = 3) Then
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                Else
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                End If
            End If
        ElseIf ilAAS = 2 Then
            If ilAASCode > 0 Then
                ilSlfCode = ilAASCode
            End If
        ElseIf ilAAS = 3 Then
            If ilAASCode > 0 Then
                ilVefCode = ilAASCode
            End If
        ElseIf ilAAS = 0 Then
            If ilAASCode > 0 Then
                tlIntTypeBuff.iType = ilAASCode
                ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
                If ilCurrent <> 3 Then
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                Else
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                End If
            End If
        End If

        tlCharTypeBuff.sType = "Y"
        ilOffSet = gFieldOffset("Chf", "ChfDelete")
        'If selecting by advertiser- bypass slCntrType and slCntrStatus Test until get contract for speed
        'If ((slCntrStatus = "") And (slCntrType = "")) Or ((ilAAS = 0) And (ilAASCode > 0)) Then
        If (slCntrType = "") Or ((ilAAS = 0) And (ilAASCode > 0)) Then
            If ilCurrent <> 3 Then
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
            End If
        Else
            If ilCurrent <> 3 Then
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            End If

            If slCntrType <> "" Then
                ilOper = BTRV_EXT_OR
                slStr = slCntrType
                Do While slStr <> ""
                    If Len(slStr) = 1 Then
                        ilOper = BTRV_EXT_LAST_TERM
                    End If
                    tlCharTypeBuff.sType = Left$(slStr, 1)
                    ilOffSet = gFieldOffset("Chf", "ChfType")
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
                    slStr = Mid$(slStr, 2)
                Loop
            End If
        End If
        ilOffSet = gFieldOffset("Chf", "ChfCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract iCode field
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfCntrNo")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract Contract number
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfExtRevNo")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfCntRevNo")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfType")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfProduct")
        ilRet = btrExtAddField(hlChf, ilOffSet, 35) 'Extract Product
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfSlfCode1")
        ilRet = btrExtAddField(hlChf, ilOffSet, 20) 'Extract salesperson code
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfMnfDemo1")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract salesperson code
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfCxfInt")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfPropVer")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract end date
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfStatus")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfMnfPotnType")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract SellNet
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfStartDate")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfEndDate")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract end date
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfVefCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfSifCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

         '8-21-05 add pct of trade to array
        ilOffSet = gFieldOffset("Chf", "ChfPctTrade")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'pct trade
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

        '7/12/10
        ilOffSet = gFieldOffset("Chf", "ChfCBSOrder")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

        '2/24/12
        ilOffSet = gFieldOffset("Chf", "ChfBillCycle")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo mPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0


        ReDim lmWCINGCntrNo(0 To 0) As Long     'Used to filter out Holds that have W or C or I or G or N
        'ilRet = btrExtGetNextExt(hlChf)    'Extract record
        ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mPopCntrForAASBoxErr
            gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrExtGetNextExt):" & "Chf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlChfAdvtExt)  'Extract operation record size
            'ilRet = btrExtGetFirst(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilFound = True
                If (ilAAS = 3) And (ilVefCode <> -1) Then   '-1 = All Vehicles
                    ilFound = False
                    If tlChfAdvtExt.lVefCode > 0 Then
                        If tlChfAdvtExt.lVefCode = ilVefCode Then
                            ilFound = True
                        End If
                    ElseIf tlChfAdvtExt.lVefCode < 0 Then
                        tlSrchKey.lCode = -tlChfAdvtExt.lVefCode
                        ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        On Error GoTo mPopCntrForAASBoxErr
                        gBtrvErrorMsg ilRet, "mPopCntrForAASBox (btrGetEqual): Vsf.Btr", frm
                        On Error GoTo 0
                        Do While ilRet = BTRV_ERR_NONE
                            For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                If tmVsf.iFSCode(ilLoop) > 0 Then
                                    If tmVsf.iFSCode(ilLoop) = ilVefCode Then
                                        ilFound = True
                                        Exit For
                                    End If
                                    If tmVsf.iFSCode(ilLoop) <> tlVef.iCode Then
                                        tlVefSrchKey.iCode = tmVsf.iFSCode(ilLoop)
                                        ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        On Error GoTo mPopCntrForAASBoxErr
                                        gCPErrorMsg ilRet, "mPopCntrForAASBox (btrGetEqual: Vef)", frm
                                        On Error GoTo 0
                                    End If
                                    If tlVef.sType = "V" Then
                                        If tlVef.iCode <> ilVefCode Then
                                            tlSrchKey.lCode = tlVef.lVsfCode
                                            ilRet = btrGetEqual(hlVsf, tlVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            For ilLoop1 = LBound(tlVsf.iFSCode) To UBound(tlVsf.iFSCode) Step 1
                                                If tlVsf.iFSCode(ilLoop1) > 0 Then
                                                    If tlVsf.iFSCode(ilLoop1) = ilVefCode Then
                                                        ilFound = True
                                                        Exit For
                                                    End If
                                                End If
                                            Next ilLoop1
                                        Else
                                            ilFound = True
                                        End If
                                    End If
                                End If
                            Next ilLoop
                            If ilFound Then
                                Exit Do
                            End If
                            If tmVsf.lLkVsfCode <= 0 Then
                                Exit Do
                            End If
                            tlSrchKey.lCode = tmVsf.lLkVsfCode
                            ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    Else
                        ilFound = True  'all vehicles
                    End If
                End If
                If ilHOType = 5 Then
                    ilTestCntrNo = True
                Else
                    ilTestCntrNo = False
                End If
                If tlChfAdvtExt.lCntrNo = 1668 Then
                    ilRet = ilRet
                End If
                'For Proposal CntRevNo = 0; For Orders CntRevNo >= 0 (for W, C, I CntRevNo > 0)
                If (tlChfAdvtExt.iCntRevNo = 0) And ((tlChfAdvtExt.sStatus <> "H") And (tlChfAdvtExt.sStatus <> "O") And (tlChfAdvtExt.sStatus <> "G") And (tlChfAdvtExt.sStatus <> "N")) Then  'Proposal
                    'Proposal only
                    If (InStr(1, slCntrStatus, tlChfAdvtExt.sStatus) = 0) Then
                        ilFound = False
                    End If
                Else    'Order only
                    If ilHOType = 5 Then
                        If (InStr(1, slCntrStatus, tlChfAdvtExt.sStatus) = 0) Then
                            ilFound = False
                        End If
                    ElseIf ilHOType = 6 Then
                        If (InStr(1, slCntrStatus, tlChfAdvtExt.sStatus) = 0) Then
                            ilFound = False
                        End If
                        'do nothing
                    Else
                        If (InStr(1, slHOStatus, tlChfAdvtExt.sStatus) <> 0) Then
                            If (ilHOType = 2) Or (ilHOType = 3) Then
                                ilTestCntrNo = True
                            End If
                        Else
                            ilFound = False
                        End If
                    End If
                End If
                If slCntrType <> "" Then
                    If InStr(1, slCntrType, tlChfAdvtExt.sType) = 0 Then
                        ilFound = False
                    End If
                End If
                If ilFound Then
                    ilFound = mTestChfAdvtExt(frm, ilSlfCode, tlChfAdvtExt, hlVsf, ilCurrent)
                End If
                If (Not ilFound) And (ilHOType = 4) And (InStr(1, "WCING", tlChfAdvtExt.sStatus) <> 0) Then
                    lmWCINGCntrNo(UBound(lmWCINGCntrNo)) = tlChfAdvtExt.lCntrNo
                    ReDim Preserve lmWCINGCntrNo(0 To UBound(lmWCINGCntrNo) + 1) As Long   'Used to filter out Holds that have W or C or I or G or N
                End If
                If (Not ilFound) And (ilHOType = 5) And (InStr(1, "CNG", tlChfAdvtExt.sStatus) <> 0) Then
                    lmWCINGCntrNo(UBound(lmWCINGCntrNo)) = tlChfAdvtExt.lCntrNo
                    ReDim Preserve lmWCINGCntrNo(0 To UBound(lmWCINGCntrNo) + 1) As Long   'Used to filter out Holds that have W or C or I or G or N
                End If
                If ilFound Then
                    slStr = Trim$(str$(99999999 - tlChfAdvtExt.lCntrNo))
                    Do While Len(slStr) < 8
                        slStr = "0" & slStr
                    Loop
                    slName = slStr
                    slStr = Trim$(str$(999 - tlChfAdvtExt.iCntRevNo))
                    Do While Len(slStr) < 3
                        slStr = "0" & slStr
                    Loop
                    slExtStr = Trim$(str$(999 - tlChfAdvtExt.iExtRevNo))
                    Do While Len(slExtStr) < 3
                        slExtStr = "0" & slExtStr
                    Loop
                    slName = slName & "|" & slStr & "-" & slExtStr & "|"
                    If (tlChfAdvtExt.sStatus = "W") Or (tlChfAdvtExt.sStatus = "C") Or (tlChfAdvtExt.sStatus = "I") Then
                        'Add Potential
                        If tlChfAdvtExt.iMnfPotnType > 0 Then
                            slStr = " "
                        Else
                            slStr = "~"
                        End If
                        slName = slName & slStr & "|"
                    Else
                        slName = slName & " |"
                    End If
                    slStr = Trim$(str$(999 - tlChfAdvtExt.iPropVer))
                    Do While Len(slStr) < 3
                        slStr = "0" & slStr
                    Loop
                    slName = slName & slStr & "|"
                    slName = slName & tlChfAdvtExt.sStatus & "|"
                    If ilShow = 0 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo)) & " V" & Trim$(Str$(tlChfAdvtExt.iPropVer))
                    ElseIf ilShow = 2 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                        slName = slName & ": " & slStartDate & "-" & slEndDate
                        slName = slName & " " & Trim$(tlChfAdvtExt.sProduct)
                        If tlChfAdvtExt.lVefCode > 0 Then
                            If tlChfAdvtExt.lVefCode <> tlVef.iCode Then
                                tlVefSrchKey.iCode = tlChfAdvtExt.lVefCode
                                ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mPopCntrForAASBoxErr
                                gCPErrorMsg ilRet, "mPopCntrForAASBox (btrGetEqual: Vef)", frm
                                On Error GoTo 0
                            End If
                            slName = slName & " " & Trim$(tlVef.sName)
                        End If
                    ElseIf ilShow = 3 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                            tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                            ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo mPopCntrForAASBoxErr
                            gCPErrorMsg ilRet, "mPopCntrForAASBox (btrGetEqual: Adf)", frm
                            On Error GoTo 0
                        End If
                        slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                        slName = slName & " " & slStartDate & "-" & slEndDate
                    ElseIf ilShow = 4 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                        slName = slName & " " & slStartDate & "-" & slEndDate
                    ElseIf ilShow = 5 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                            tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                            ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo mPopCntrForAASBoxErr
                            gCPErrorMsg ilRet, "mPopCntrForAASBox (btrGetEqual: Adf)", frm
                            On Error GoTo 0
                        End If
                        slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                    Else    '1 or 6
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo)) & " V" & Trim$(Str$(tlChfAdvtExt.iPropVer))
                        Select Case tlChfAdvtExt.sStatus
                            Case "W"
                                If tlChfAdvtExt.iCntRevNo > 0 Then
                                    slStr = "Rev Working"
                                Else
                                    slStr = "Working"
                                End If
                            Case "D"
                                slStr = "Rejected"          '4-29-09 chged from dead
                            Case "C"
                                If tlChfAdvtExt.iCntRevNo > 0 Then
                                    slStr = "Rev Completed"
                                Else
                                    slStr = "Completed"
                                End If
                            Case "I"
                                If tlChfAdvtExt.iCntRevNo > 0 Then
                                    slStr = "Rev Unapproved"    '4-29-09 chged from incomplete
                                Else
                                    slStr = "Unapproved"        '4-29-09 chged from incomplete
                                End If
                            Case "G"
                                slStr = "Unappr Hold"           '4-29-09 chged from unsch hold
                            Case "N"
                                slStr = "Unappr Order"          '4-29-09 chged from unsch order
                            Case "H"
                                slStr = "Hold"
                            Case "O"
                                slStr = "Order"
                        End Select
                        slName = slName & " " & slStr
                        If ilShow = 6 Then
                            If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                                tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                                ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mPopCntrForAASBoxErr
                                gCPErrorMsg ilRet, "mPopCntrForAASBox (btrGetEqual: Adf)", frm
                                On Error GoTo 0
                            End If
                            slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                        Else    'show =1
                            If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                                 tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                                 ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                 On Error GoTo mPopCntrForAASBoxErr
                                 gCPErrorMsg ilRet, "mPopCntrForAASBox (btrGetEqual: Adf)", frm
                                 On Error GoTo 0
                             End If
                             slAdvt = Trim$(tlAdf.sName)        '10 max char
                             slProd = Trim$(tlChfAdvtExt.sProduct)
                             slName = slName & " " & Trim$(slAdvt) & "/" & Trim$(slProd)
                        End If
                        'Start Date Plus # weeks
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDateLong tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), llEndDate
                        If ilShow = 6 Then
                            slName = slName & " " & slStartDate & "-" & Format$(llEndDate, "m/d/yy")
                        Else    'show = 1
                            slStr = str$((llEndDate - gDateValue(slStartDate)) \ 7 + 1)
                            slName = slName & " " & slStartDate & slStr
                            'no comments for this special version of PopCntrForAASBox
                            'tmCxfSrchKey.lCode = tlChfAdvtExt.lCxfInt
                            'If tmCxfSrchKey.lCode <> 0 Then
                            '    tmCxf.sComment = ""
                            '    imCxfRecLen = Len(tmCxf) '5027
                            '    ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            '    If ilRet = BTRV_ERR_NONE Then
                            '        If tmCxf.iStrLen > 0 Then
                            '            If tmCxf.iStrLen < 40 Then
                            '                slName = slName & " " & Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
                            '            Else
                            '                slName = slName & " " & Trim$(Left$(tmCxf.sComment, 40))
                            '            End If
                            '        End If
                            '    End If
                            'End If
                        End If
                    End If
                    ilFound = False
                    If ilTestCntrNo Then
                        For ilLoop = 0 To ilSortCode - 1 Step 1
                            slNameCode = tlSortCode(ilLoop).sKey    'lbcMster.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 1, "\", slNameCode)
                            ilRet = gParseItem(slNameCode, 1, "|", slCode)
                            llCntrNo = 99999999 - CLng(slCode)
                            If llCntrNo = tlChfAdvtExt.lCntrNo Then
                                ilRet = gParseItem(slNameCode, 2, "|", slCode)
                                ilRet = gParseItem(slCode, 1, "-", slExtStr)
                                ilRevNo = 999 - CLng(slExtStr)
                                If tlChfAdvtExt.iCntRevNo > ilRevNo Then
                                    'Replace
                                    'lbcMster.RemoveItem ilLoop
                                    'llLen = llLen - Len(slNameCode)
                                    ilFound = True
                                    slName = slName & "\" & Trim$(str$(tlChfAdvtExt.lCode))
                                    tlSortCode(ilLoop).sKey = slName
                                Else
                                    'Leave
                                    ilFound = True
                                End If
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilFound Then
                        slName = slName & "\" & Trim$(str$(tlChfAdvtExt.lCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit Do
                        'End If
                        'lbcMster.AddItem slName
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                End If
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
                Loop

            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
            If (ilHOType = 4) Or (ilHOType = 5) Then
                For ilLoop1 = 0 To UBound(lmWCINGCntrNo) - 1 Step 1
                    For ilLoop = UBound(tlSortCode) - 1 To 0 Step -1
                        slNameCode = tlSortCode(ilLoop).sKey    'lbcMster.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 1, "\", slName)
                        ilRet = gParseItem(slName, 1, "|", slCode)
                        llCntrNo = 99999999 - CLng(slCode)
                        If lmWCINGCntrNo(ilLoop1) = llCntrNo Then
                            For ilIndex = ilLoop To UBound(tlSortCode) - 1 Step 1
                                tlSortCode(ilIndex) = tlSortCode(ilIndex + 1)
                            Next ilIndex
                            ReDim Preserve tlSortCode(0 To UBound(tlSortCode) - 1) As SORTCODE
                        End If
                    Next ilLoop
                Next ilLoop1
            End If
            Erase lmWCINGCntrNo
            If UBound(tlSortCode) - 1 > 0 Then
                ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
            End If
        End If
        ilRet = btrClose(hmCxf)
        btrDestroy hmCxf
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        ilRet = btrClose(hlVsf)
        btrDestroy hlVsf
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        ilRet = btrClose(hlChf)
        btrDestroy hlChf
    End If
    llLen = 0
    For ilLoop = 0 To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 1, "|", slCode)
        llCntrNo = 99999999 - CLng(slCode)
        slShow = Trim$(str$(llCntrNo))
        ilRet = gParseItem(slName, 2, "|", slCode)
        ilRet = gParseItem(slCode, 1, "-", slExtStr)
        ilRevNo = 999 - CLng(slExtStr)
        ilRet = gParseItem(slCode, 2, "-", slExtStr)
        ilExtRevNo = 999 - CLng(slExtStr)
        ilRet = gParseItem(slName, 4, "|", slCode)
        ilVerNo = 999 - CLng(slCode)
        ilRet = gParseItem(slName, 5, "|", slCode)
        If (slCode = "W") Or (slCode = "C") Or (slCode = "I") Or (slCode = "D") Then
            If ilRevNo > 0 Then
                slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
            Else
                slShow = slShow & " V" & Trim$(str$(ilVerNo))
            End If
        Else
            slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
        End If
        If ilShow = 0 Then      'Number only
        ElseIf ilShow = 2 Then  'Number, Dates, Product, Vehicle
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 3 Then  'Number, Advertiser, Dates
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 4 Then  'Number, Dates
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 5 Then  'Number, Advertiser
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        Else                    'Number, Product, Internal comment
            'Potential
            ilRet = gParseItem(slName, 3, "|", slCode)
            If (Trim$(slCode) <> "") And (slCode <> "~") Then
                slShow = slShow & " " & slCode
            End If
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        End If
        If Not gOkAddStrToListBox(slShow, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slShow  'Add ID to list box
    Next ilLoop
    Exit Function
mPopCntrForAASBoxErr:
    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf
    ilRet = btrClose(hlAdf)
    btrDestroy hlAdf
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    gDbg_HandleError "RptSelPC: mPopCntrForAASBox"
'    mPopCntrForAASBox = CP_MSG_NOSHOW
'    Exit Function
mPopCntrForAASBoxErr2:
    ilRet = 1
    Resume Next
End Function
