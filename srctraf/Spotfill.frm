VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SpotFill 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6210
   ClientLeft      =   2250
   ClientTop       =   1470
   ClientWidth     =   9345
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   9345
   Begin VB.CheckBox ckcAdjBreaks 
      Caption         =   "Treat Adjacent Breaks as One Break"
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
      Height          =   195
      Left            =   120
      TabIndex        =   57
      Top             =   5925
      Width           =   3510
   End
   Begin VB.CheckBox ckcEmptyAvails 
      Caption         =   "Ignore Empty Avails"
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
      Height          =   195
      Left            =   6015
      TabIndex        =   53
      Top             =   5460
      Width           =   2040
   End
   Begin VB.CheckBox ckcAvailNames 
      Caption         =   "Honor Avail Names"
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
      Height          =   195
      Left            =   6015
      TabIndex        =   48
      Top             =   5220
      Width           =   2010
   End
   Begin VB.CheckBox ckc10To15 
      Caption         =   "Allow 10"" spots into 15"" Avails"
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
      Height          =   195
      Left            =   4755
      TabIndex        =   56
      Top             =   5670
      Value           =   1  'Checked
      Width           =   3150
   End
   Begin VB.CheckBox ckcLock 
      Caption         =   "Skip Locked Avails"
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
      Height          =   195
      Left            =   2535
      TabIndex        =   55
      Top             =   5670
      Value           =   1  'Checked
      Width           =   1980
   End
   Begin VB.PictureBox plcCalendar 
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
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1905
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Spotfill.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   104
            Top             =   405
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   99
         Top             =   30
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcAC 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   5760
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   5445
      Width           =   5760
      Begin VB.OptionButton rbcAC 
         Caption         =   "Not Same Break"
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
         Height          =   195
         Index           =   1
         Left            =   3150
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   0
         Width           =   1710
      End
      Begin VB.OptionButton rbcAC 
         Caption         =   "Vehicle Rules"
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
         Height          =   195
         Index           =   0
         Left            =   1665
         TabIndex        =   50
         Top             =   0
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.OptionButton rbcAC 
         Caption         =   "None"
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
         Height          =   195
         Index           =   2
         Left            =   4860
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.CheckBox ckcDaysTimes 
      Caption         =   "Use Line Days/Times"
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
      Height          =   195
      Left            =   120
      TabIndex        =   54
      Top             =   5670
      Value           =   1  'Checked
      Width           =   2145
   End
   Begin VB.ListBox lbcMissedSort 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "Spotfill.frx":2E1A
      Left            =   9225
      List            =   "Spotfill.frx":2E1C
      Sorted          =   -1  'True
      TabIndex        =   79
      Top             =   5445
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   5580
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3045
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "Spotfill.frx":2E1E
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "Spotfill.frx":3ADC
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox plcFillInv 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   4020
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5220
      Width           =   4020
      Begin VB.OptionButton rbcFillInv 
         Caption         =   "as Advt Set"
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
         Height          =   195
         Index           =   2
         Left            =   1260
         TabIndex        =   45
         Top             =   0
         Value           =   -1  'True
         Width           =   1350
      End
      Begin VB.OptionButton rbcFillInv 
         Caption         =   "No"
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
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   47
         Top             =   0
         Width           =   600
      End
      Begin VB.OptionButton rbcFillInv 
         Caption         =   "Yes"
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
         Height          =   195
         Index           =   0
         Left            =   2655
         TabIndex        =   46
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9255
      Top             =   5010
   End
   Begin VB.Timer tmcLine 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   9285
      Top             =   4515
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
      ScaleWidth      =   120
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   5220
      Width           =   120
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   150
      ScaleHeight     =   270
      ScaleWidth      =   6945
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6945
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Fill"
      Height          =   285
      Left            =   8175
      TabIndex        =   58
      Top             =   5505
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   8175
      TabIndex        =   59
      Top             =   5850
      Width           =   1050
   End
   Begin VB.PictureBox plcInv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4845
      Left            =   90
      ScaleHeight     =   4785
      ScaleWidth      =   9075
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   315
      Width           =   9135
      Begin VB.CommandButton cmcTVeh 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7425
         Picture         =   "Spotfill.frx":3DE6
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   45
         Width           =   195
      End
      Begin VB.PictureBox pbcFillVehType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
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
         FontTransparent =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   5250
         ScaleHeight     =   210
         ScaleWidth      =   2175
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   45
         Width           =   2175
      End
      Begin VB.Frame frcWith 
         Caption         =   "With"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2850
         Left            =   120
         TabIndex        =   34
         Top             =   1875
         Width           =   8865
         Begin VB.PictureBox pbcLbcLines 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   2085
            Left            =   75
            ScaleHeight     =   2085
            ScaleWidth      =   8700
            TabIndex        =   105
            Top             =   630
            Width           =   8700
         End
         Begin VB.CheckBox ckcSpotType 
            Caption         =   "DRs"
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
            Height          =   225
            Index           =   6
            Left            =   4755
            TabIndex        =   42
            Top             =   360
            Width           =   705
         End
         Begin VB.CheckBox ckcSpotType 
            Caption         =   "Remnants"
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
            Height          =   225
            Index           =   5
            Left            =   2955
            TabIndex        =   41
            Top             =   360
            Width           =   1170
         End
         Begin VB.PictureBox pbcDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
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
            FontTransparent =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   75
            ScaleHeight     =   210
            ScaleWidth      =   1035
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   270
            Width           =   1035
         End
         Begin VB.CheckBox ckcSpotType 
            Caption         =   "PIs"
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
            Height          =   225
            Index           =   4
            Left            =   4155
            TabIndex        =   40
            Top             =   360
            Width           =   600
         End
         Begin VB.CheckBox ckcSpotType 
            Caption         =   "Promos"
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
            Height          =   225
            Index           =   3
            Left            =   1995
            TabIndex        =   39
            Top             =   360
            Width           =   960
         End
         Begin VB.ListBox lbcLines 
            Appearance      =   0  'Flat
            Height          =   2130
            Left            =   60
            MultiSelect     =   2  'Extended
            TabIndex        =   43
            Top             =   615
            Width           =   8745
         End
         Begin VB.CheckBox ckcSpotType 
            Caption         =   "$ Spots"
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
            Height          =   225
            Index           =   0
            Left            =   1200
            TabIndex        =   36
            Top             =   135
            Value           =   1  'Checked
            Width           =   1050
         End
         Begin VB.CheckBox ckcSpotType 
            Caption         =   "N/C, .00, Bonus Spots"
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
            Height          =   225
            Index           =   1
            Left            =   2250
            TabIndex        =   37
            Top             =   135
            Width           =   2265
         End
         Begin VB.CheckBox ckcSpotType 
            Caption         =   "PSAs"
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
            Height          =   225
            Index           =   2
            Left            =   1200
            TabIndex        =   38
            Top             =   360
            Width           =   795
         End
         Begin VB.Label lacNote 
            Appearance      =   0  'Flat
            Caption         =   "(Trades, MG's Excluded from List)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   6375
            TabIndex        =   80
            Top             =   120
            Width           =   2445
         End
         Begin VB.Label lacTitle 
            Appearance      =   0  'Flat
            Caption         =   "Rate  Sch.  Extra"
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
            Left            =   7320
            TabIndex        =   78
            Top             =   360
            Width           =   1485
         End
      End
      Begin VB.Frame frcFromGame 
         Caption         =   "From"
         Height          =   1860
         Left            =   135
         TabIndex        =   106
         Top             =   0
         Visible         =   0   'False
         Width           =   4500
         Begin ComctlLib.ListView lbcFromGame 
            Height          =   1155
            Left            =   180
            TabIndex        =   110
            Top             =   285
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   3
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Event #"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Teams"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CheckBox ckcFromGame 
            Caption         =   "All Events"
            Height          =   210
            Left            =   195
            TabIndex        =   107
            Top             =   1575
            Width           =   1785
         End
      End
      Begin VB.Frame frcFrom 
         Caption         =   "From"
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
         Height          =   1830
         Left            =   120
         TabIndex        =   2
         Top             =   30
         Width           =   4530
         Begin VB.CommandButton cmcFDropDown 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   4185
            Picture         =   "Spotfill.frx":3EE0
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   660
            Width           =   195
         End
         Begin VB.TextBox edcFDropDown 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   3
            Left            =   3165
            MaxLength       =   20
            TabIndex        =   8
            Top             =   660
            Width           =   1020
         End
         Begin VB.CommandButton cmcFDropDown 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   4185
            Picture         =   "Spotfill.frx":3FDA
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   375
            Width           =   195
         End
         Begin VB.TextBox edcFDropDown 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   2
            Left            =   3165
            MaxLength       =   20
            TabIndex        =   6
            Top             =   375
            Width           =   1020
         End
         Begin VB.CheckBox ckcFAll 
            Caption         =   "All Vehicles"
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
            Height          =   195
            Left            =   165
            TabIndex        =   4
            Top             =   1545
            Width           =   1350
         End
         Begin VB.TextBox edcFDropDown 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   1
            Left            =   3165
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1470
            Width           =   1020
         End
         Begin VB.CommandButton cmcFDropDown 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   4185
            Picture         =   "Spotfill.frx":40D4
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1470
            Width           =   195
         End
         Begin VB.TextBox edcFDropDown 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   0
            Left            =   3165
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1185
            Width           =   1020
         End
         Begin VB.CommandButton cmcFDropDown 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   4185
            Picture         =   "Spotfill.frx":41CE
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1185
            Width           =   195
         End
         Begin VB.ListBox lbcVehicle 
            Appearance      =   0  'Flat
            Height          =   1080
            Left            =   165
            MultiSelect     =   2  'Extended
            TabIndex        =   3
            Top             =   255
            Width           =   2865
         End
         Begin VB.Label lacFDate 
            Appearance      =   0  'Flat
            Caption         =   "Start/End Date"
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
            Left            =   3120
            TabIndex        =   5
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label lacFSTime 
            Appearance      =   0  'Flat
            Caption         =   "Start/End Time"
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
            Left            =   3120
            TabIndex        =   10
            Top             =   960
            Width           =   1380
         End
      End
      Begin VB.Frame frcTo 
         Caption         =   "To        "
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
         Height          =   1830
         Left            =   4710
         TabIndex        =   15
         Top             =   30
         Width           =   4260
         Begin VB.PictureBox plcTDays 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   825
            ScaleHeight     =   135
            ScaleWidth      =   3315
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   3375
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "S"
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
               Height          =   195
               Index           =   13
               Left            =   2775
               TabIndex        =   97
               Top             =   0
               Value           =   1  'Checked
               Width           =   465
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "S"
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
               Height          =   195
               Index           =   12
               Left            =   2265
               TabIndex        =   96
               Top             =   0
               Value           =   1  'Checked
               Width           =   480
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "F"
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
               Height          =   195
               Index           =   11
               Left            =   1845
               TabIndex        =   95
               Top             =   0
               Value           =   1  'Checked
               Width           =   390
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "T"
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
               Height          =   195
               Index           =   10
               Left            =   1365
               TabIndex        =   94
               Top             =   0
               Value           =   1  'Checked
               Width           =   495
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "W"
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
               Height          =   195
               Index           =   9
               Left            =   930
               TabIndex        =   93
               Top             =   0
               Value           =   1  'Checked
               Width           =   465
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "T"
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
               Height          =   195
               Index           =   8
               Left            =   405
               TabIndex        =   92
               Top             =   0
               Value           =   1  'Checked
               Width           =   495
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "M"
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
               Height          =   195
               Index           =   7
               Left            =   0
               TabIndex        =   91
               Top             =   0
               Value           =   1  'Checked
               Width           =   405
            End
         End
         Begin VB.CheckBox ckcTAll 
            Caption         =   "All "
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
            Height          =   195
            Left            =   3045
            TabIndex        =   89
            Top             =   1380
            Width           =   630
         End
         Begin VB.PictureBox plcSVeh 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1155
            Left            =   90
            ScaleHeight     =   1095
            ScaleWidth      =   4050
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   270
            Width           =   4110
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "M"
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
               Height          =   195
               Index           =   0
               Left            =   795
               TabIndex        =   20
               Top             =   105
               Value           =   1  'Checked
               Width           =   405
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "T"
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
               Height          =   195
               Index           =   1
               Left            =   1200
               TabIndex        =   21
               Top             =   105
               Value           =   1  'Checked
               Width           =   540
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "W"
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
               Height          =   195
               Index           =   2
               Left            =   1725
               TabIndex        =   22
               Top             =   105
               Value           =   1  'Checked
               Width           =   465
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "T"
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
               Height          =   195
               Index           =   3
               Left            =   2160
               TabIndex        =   23
               Top             =   105
               Value           =   1  'Checked
               Width           =   495
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "F"
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
               Height          =   195
               Index           =   4
               Left            =   2640
               TabIndex        =   24
               Top             =   105
               Value           =   1  'Checked
               Width           =   390
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "S"
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
               Height          =   195
               Index           =   5
               Left            =   3060
               TabIndex        =   25
               Top             =   105
               Value           =   1  'Checked
               Width           =   480
            End
            Begin VB.CheckBox ckcDay 
               Alignment       =   1  'Right Justify
               Caption         =   "S"
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
               Height          =   195
               Index           =   6
               Left            =   3570
               TabIndex        =   26
               Top             =   105
               Value           =   1  'Checked
               Width           =   465
            End
            Begin VB.ComboBox cbcLen 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   30
               TabIndex        =   27
               Top             =   660
               Width           =   720
            End
            Begin VB.Label lacN60 
               Appearance      =   0  'Flat
               Caption         =   "#60"
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
               Height          =   225
               Left            =   390
               TabIndex        =   63
               Top             =   480
               Width           =   345
            End
            Begin VB.Label lac30 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   825
               TabIndex        =   64
               Top             =   315
               Width           =   360
            End
            Begin VB.Label lac30 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1275
               TabIndex        =   65
               Top             =   315
               Width           =   360
            End
            Begin VB.Label lac30 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   1800
               TabIndex        =   66
               Top             =   315
               Width           =   360
            End
            Begin VB.Label lac30 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2250
               TabIndex        =   67
               Top             =   315
               Width           =   360
            End
            Begin VB.Label lac30 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   2670
               TabIndex        =   68
               Top             =   315
               Width           =   360
            End
            Begin VB.Label lac30 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   5
               Left            =   3135
               TabIndex        =   69
               Top             =   315
               Width           =   360
            End
            Begin VB.Label lac30 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   6
               Left            =   3615
               TabIndex        =   70
               Top             =   315
               Width           =   360
            End
            Begin VB.Label lacN30 
               Appearance      =   0  'Flat
               Caption         =   "#30"
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
               Height          =   225
               Left            =   390
               TabIndex        =   71
               Top             =   315
               Width           =   345
            End
            Begin VB.Label lac60 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   6
               Left            =   3615
               TabIndex        =   72
               Top             =   495
               Width           =   360
            End
            Begin VB.Label lac60 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   5
               Left            =   3135
               TabIndex        =   73
               Top             =   495
               Width           =   360
            End
            Begin VB.Label lac60 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   2670
               TabIndex        =   74
               Top             =   495
               Width           =   360
            End
            Begin VB.Label lac60 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2250
               TabIndex        =   75
               Top             =   495
               Width           =   360
            End
            Begin VB.Label lac60 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   1800
               TabIndex        =   76
               Top             =   495
               Width           =   360
            End
            Begin VB.Label lac60 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1275
               TabIndex        =   77
               Top             =   495
               Width           =   360
            End
            Begin VB.Label lac60 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   825
               TabIndex        =   81
               Top             =   495
               Width           =   360
            End
            Begin VB.Label lacLen 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   825
               TabIndex        =   82
               Top             =   675
               Width           =   360
            End
            Begin VB.Label lacLen 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1275
               TabIndex        =   83
               Top             =   675
               Width           =   360
            End
            Begin VB.Label lacLen 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   1800
               TabIndex        =   84
               Top             =   675
               Width           =   360
            End
            Begin VB.Label lacLen 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2250
               TabIndex        =   85
               Top             =   675
               Width           =   360
            End
            Begin VB.Label lacLen 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   2670
               TabIndex        =   86
               Top             =   675
               Width           =   360
            End
            Begin VB.Label lacLen 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   5
               Left            =   3135
               TabIndex        =   87
               Top             =   675
               Width           =   360
            End
            Begin VB.Label lacLen 
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
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   6
               Left            =   3615
               TabIndex        =   88
               Top             =   675
               Width           =   360
            End
         End
         Begin VB.ListBox lbcTVehicle 
            Appearance      =   0  'Flat
            Height          =   1080
            Index           =   0
            Left            =   165
            TabIndex        =   17
            Top             =   255
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.ListBox lbcTVehicle 
            Appearance      =   0  'Flat
            Height          =   1080
            Index           =   1
            Left            =   165
            MultiSelect     =   2  'Extended
            TabIndex        =   18
            Top             =   255
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.TextBox edcTDropDown 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   0
            Left            =   975
            MaxLength       =   20
            TabIndex        =   29
            Top             =   1485
            Width           =   885
         End
         Begin VB.CommandButton cmcTDropDown 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1860
            Picture         =   "Spotfill.frx":42C8
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1485
            Width           =   195
         End
         Begin VB.TextBox edcTDropDown 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   1
            Left            =   3090
            MaxLength       =   20
            TabIndex        =   32
            Top             =   1485
            Width           =   885
         End
         Begin VB.CommandButton cmcTDropDown 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   3975
            Picture         =   "Spotfill.frx":43C2
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1485
            Width           =   195
         End
         Begin VB.Label lacTSTime 
            Appearance      =   0  'Flat
            Caption         =   "Start Time"
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
            Left            =   75
            TabIndex        =   28
            Top             =   1485
            Width           =   885
         End
         Begin VB.Label lacTETime 
            Appearance      =   0  'Flat
            Caption         =   "End Time"
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
            Left            =   2145
            TabIndex        =   31
            Top             =   1485
            Width           =   885
         End
      End
      Begin VB.Frame frcToGame 
         Caption         =   "To"
         Height          =   1800
         Left            =   4710
         TabIndex        =   108
         Top             =   30
         Visible         =   0   'False
         Width           =   4245
         Begin ComctlLib.ListView lbcToGame 
            Height          =   1080
            Left            =   240
            TabIndex        =   111
            Top             =   330
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   1905
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Event #"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "# 30s"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "# 60s"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CheckBox ckcToGame 
            Caption         =   "All Events"
            Height          =   210
            Left            =   165
            TabIndex        =   109
            Top             =   1530
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "SpotFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Spotfill.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmGhfSrchKey0                 tmGhfSrchKey1                 tmGsfSrchKey0             *
'*  tmGsfSrchKey1                 tmCgfCff                                                *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mGhfGsfReadRec                                                                        *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SpotFill.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract terminate screen code
Option Explicit
Option Compare Text
'Spot detail record information
Dim hmSdf As Integer        'Spot detail file handle
Dim tmSdf As SDF            'SDF record image
Dim tmSdfSrchKey1 As SDFKEY1 'SDF key record image
Dim tmSdfSrchKey3 As LONGKEY0 'SDF key record image
Dim tmSdfSrchKey6 As SDFKEY6
Dim imSdfRecLen As Integer  'SDF record length
Dim tmSdfMdExt() As SDFMDEXT 'Spot summary
Dim smSdfMdExtTag As String
Dim imLBSdfMdExt As Integer
Dim imGetMissed As Integer
Dim hmSsf As Integer        'Spot summary file handle
Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
Dim tmSsfSrchKey1 As SSFKEY1 'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2 'SSF key record image
Dim imSsfRecLen As Integer  'SSF record length
Dim tmSsf As SSF         'Spot summary for one week
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

'Advertiser
Dim hmAdf As Integer            'Advertiser file handle
Dim tmAdf As ADF               'ADF record image
Dim tmAdfSrchKey As INTKEY0     'ADF key record image
Dim imAdfRecLen As Integer         'ADF record length
'Contract line
Dim hmCHF As Integer        'Contract line file handle
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer     'CHF record length
'Contract line
Dim hmClf As Integer        'Contract line file handle
Dim tmClf As CLF            'CLF record image
Dim tmClfSrchKey As CLFKEY0 'CLF key record image
Dim imClfRecLen As Integer     'CLF record length
' Contract Flight File
Dim hmCff As Integer        'Contract Flight file handle
Dim tmCff As CFF            'CFF record image (array required for compatiblity with sch routines)
Dim tmICff(0 To 2) As CFF
Dim tmFCff() As CFF
Dim tmGCff() As CFF
Dim tmCffSrchKey As CFFKEY0 'CFF key record image
Dim imCffRecLen As Integer     'CFF record length

Dim hmGhf As Integer
Dim tmGhf As GHF        'GHF record image
Dim imGhfRecLen As Integer        'GHF record length
Dim tmGhfSrchKey0 As LONGKEY0 'GHF key record image

Dim hmGsf As Integer
Dim tmGsf() As GSF        'GSF record image
Dim tmGsfSrchKey3 As GSFKEY3
Dim imGsfRecLen As Integer        'GSF record length

Dim hmCgf As Integer
Dim tmCgf As CGF
Dim imCgfRecLen As Integer
Dim tmCgfSrchKey1 As CGFKEY1    'CntrNo; CntRevNo; PropVer

Dim imBkQH As Integer
Dim imPriceLevel As Integer
Dim lmSepLength As Long 'Separation length for advertiser
Dim lmCompTime As Long  'Competitive time for vehicle
' Rate Card Programs/Times File
Dim hmRdf As Integer        'Rate Card Programs/Times file handle
Dim tmRdf As RDF            'RDF record image
Dim tmRdfSrchKey As INTKEY0 'RDF key record image
Dim imRdfRecLen As Integer     'RDF record length
' Spot Tracking File (only only if spots can be moved from Todays date+1 to Last log date)
Dim hmStf As Integer        'Spot tracking file handle
Dim tmStf As STF            'STF record image
Dim imStfRecLen As Integer  'STF record length
'Spot MG record
Dim hmSmf As Integer        'Spot MG file handle
Dim tmSmf As SMF            'SMF record image
Dim imSmfRecLen As Integer  'SMF record length

'Record Lock
Dim hmRlf As Integer

'Feed
Dim hmFsf As Integer
Dim tmFsf As FSF            'FSF record image
Dim tmFSFSrchKey As LONGKEY0 'FSF key record image
Dim imFsfRecLen As Integer     'FSF record length

'Feed Name
Dim hmFnf As Integer

'Product
Dim hmPrf As Integer

Dim hmLcf As Integer

Dim hmSxf As Integer

Dim tmGsfInfo() As GSFINFO
Dim tmTeam() As MNF
Dim smTeamTag As String

'Copy Rotation
Dim hmCrf As Integer
Dim lmSDate As Long     'Original date viewed on spot screen
Dim lmEDate As Long     'Original date viewed on spot screen
Dim lmPSTime As Long    'Previous start time
Dim lmPETime As Long    'Previous End Time
Dim imFDateIndex As Integer 'From date Index
Dim imFTimeIndex As Integer 'From time Index
Dim imTTimeIndex As Integer 'To Time Index
Dim imSelLen As Integer
Dim imIgnoreChg As Integer
Dim imDetail As Integer 'Show detail: 0=Yes; 1=No (Contract summary)
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBSMode As Integer
Dim imBypassFocus As Integer
Dim imVpfIndex As Integer
Dim lmLastLogDate As Long
Dim lmDatesClearedResv() As Long    'Dates cleared of reservation spots
Dim imFillVehType As Integer    '0=Single; 1=Multi-Vehicle
Dim imFromVefCode As Integer
Dim imToVefCode As Integer      'Single mode
Dim imBypassAll As Integer
Dim imUpdateAllowed As Integer
Dim smNowDate As String
Dim lmNowDate As Long

Dim imGameVehicle As Integer
Dim imMixtureOfVehicles As Integer

Dim smSingleName As String
Dim imDays(0 To 6) As Integer   'Day status
Dim tmChfAdvtExt() As CHFADVTEXT
'Drag
Dim imLbcHeight As Integer
Dim imDragIndexSrce As Integer  '
Dim imDragIndexDest As Integer  '
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragButton As Integer 'Value 1= Left button; 2=Right button; 4=Middle button
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer  'Shift state when mouse down event occurrs
Dim imDragSrce As Integer 'Values defined below
Dim imDragDest As Integer 'Values defined below
Const DRAGLINE = 1
'Dim tmSpotMove(1 To 2) As SPOTMOVE
Dim tmSpotMove(0 To 1) As SPOTMOVE
Dim tmVcf0() As VCF
Dim tmVcf6() As VCF
Dim tmVcf7() As VCF
'Required to be compatible with general schedule routines
'The array are not used by spots except for compatiblity
'Dim lmTBStartTime(1 To 49) As Long  'Allowed times if time buy (-1 indicates end of times) and library times if library buy
'Dim lmTBEndTime(1 To 49) As Long
Dim lmTBStartTime(0 To 48) As Long  'Allowed times if time buy (-1 indicates end of times) and library times if library buy
Dim lmTBEndTime(0 To 48) As Long
'Dim imHour(1 To 24) As Integer     'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
'Dim imDay(1 To 7) As Integer       'Day count (Index 1 for monday; Index 2 for tuesday;..)
'Dim imQH(1 To 4) As Integer        'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
''Actual for the day or week be processed- this will be a subset from
''imC---- or imP----
'Dim imAHour(1 To 24) As Integer    'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
'Dim imADay(1 To 7) As Integer      'Day count (Index 1 for monday; Index 2 for tuesday;..)
'Dim imAQH(1 To 4) As Integer       'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
'Dim imSkip(1 To 24, 1 To 4, 0 To 6) As Integer  '-1=Skip all test;0=All test;
'                                    'Bit 0=Skip insert;
'                                    'Bit 1=Skip move;
'                                    'Bit 2=Skip competitive pack;
'                                    'Bit 3=Skip Preempt
Dim imHour(0 To 23) As Integer     'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
Dim imDay(0 To 6) As Integer       'Day count (Index 1 for monday; Index 2 for tuesday;..)
Dim imQH(0 To 3) As Integer        'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
'Actual for the day or week be processed- this will be a subset from
'imC---- or imP----
Dim imAHour(0 To 23) As Integer    'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
Dim imADay(0 To 6) As Integer      'Day count (Index 1 for monday; Index 2 for tuesday;..)
Dim imAQH(0 To 3) As Integer       'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
Dim imSkip(0 To 23, 0 To 3, 0 To 6) As Integer  '-1=Skip all test;0=All test;
                                    'Bit 0=Skip insert;
                                    'Bit 1=Skip move;
                                    'Bit 2=Skip competitive pack;
                                    'Bit 3=Skip Preempt


'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls  As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
'Dim imListField(1 To 10) As Integer
Dim imListField(0 To 10) As Integer
Dim imLBCtrls As Integer

Private bmFirstCallToVpfFind As Boolean

Const MOVENONTOSPORT = &H1
Const MOVESPORTTOSPORT = &H2
Const MOVESPORTTONON = &H4

'Dim imListFieldChar(1 To 9) As Integer

Private Sub cbcLen_Change()
    If (imIgnoreChg = True) Then
        imIgnoreChg = False
        Exit Sub
    End If
    If imSelLen <> Val(cbcLen.List(cbcLen.ListIndex)) Then
        lmPETime = -1
        imSelLen = Val(cbcLen.List(cbcLen.ListIndex))
        mGet30Count 0, False
    End If
End Sub
Private Sub cbcLen_Click()
    cbcLen_Change
End Sub
Private Sub cbcLen_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcAdjBreaks_Click()
    If ckcAdjBreaks.Value = vbChecked Then
        ckcEmptyAvails.Value = vbChecked
    End If
End Sub

Private Sub ckcDay_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcDay(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    tmcLine.Enabled = False
    tmcLine.Enabled = True
End Sub
Private Sub ckcDay_DragDrop(Index As Integer, Source As control, X As Single, Y As Single)
    Dim slSTime As String
    Dim slETime As String
    Dim llSTime As Long
    Dim llETime As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCode As String
    Dim ilLoop As Integer
    Dim slLength As String
    If (imDragIndexSrce >= 0) And (imDragIndexSrce <= lbcLines.ListCount - 1) Then
        If (imDragDest >= 0) And (imDragDest <= 6) Then
            If ckcDay(imDragDest).Value = vbChecked Then
                Screen.MousePointer = vbHourglass
                slSTime = edcTDropDown(0).Text
                If (Not gValidTime(slSTime)) Or (slSTime = "") Then
                    Screen.MousePointer = vbDefault
                    Beep
                    edcTDropDown(0).SetFocus
                    Exit Sub
                End If
                slETime = edcTDropDown(1).Text
                If (Not gValidTime(slETime)) Or (slETime = "") Then
                    Screen.MousePointer = vbDefault
                    Beep
                    edcTDropDown(1).SetFocus
                    Exit Sub
                End If
                llSTime = CLng(gTimeToCurrency(slSTime, False))
                llETime = CLng(gTimeToCurrency(slETime, True))
                If llETime < llSTime Then
                    Screen.MousePointer = vbDefault
                    Beep
                    edcTDropDown(1).SetFocus
                    Exit Sub
                End If
                llSDate = lmSDate + imDragDest
                llEDate = llSDate
                If lbcTVehicle(0).ListIndex < 0 Then
                    Screen.MousePointer = vbDefault
                    Beep
                    lbcTVehicle(0).SetFocus
                    Exit Sub
                End If
                slNameCode = tgUserVehicle(lbcTVehicle(0).ListIndex).sKey 'Traffic!lbcUserVehicle.List(ilVef)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                imToVefCode = Val(slCode)
                If bmFirstCallToVpfFind Then
                    imVpfIndex = gVpfFind(SpotFill, imToVefCode)
                    bmFirstCallToVpfFind = False
                Else
                    imVpfIndex = gVpfFindIndex(imToVefCode)
                End If
                If (tgVpf(imVpfIndex).iLLD(0) <> 0) Or (tgVpf(imVpfIndex).iLLD(1) <> 0) Then
                    gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLastLogDate
                Else
                    lmLastLogDate = -1
                End If
                If tgVpf(imVpfIndex).sSCompType = "T" Then
                    gUnpackLength tgVpf(imVpfIndex).iSCompLen(0), tgVpf(imVpfIndex).iSCompLen(1), "3", False, slLength
                    lmCompTime = CLng(gLengthToCurrency(slLength))
                Else
                    lmCompTime = 0&
                End If
                For ilLoop = 0 To 6 Step 1
                    If ckcDay(ilLoop).Value = vbChecked Then
                        imDays(ilLoop) = True
                    Else
                        imDays(ilLoop) = False
                    End If
                Next ilLoop
                mBookSpot llSDate, llEDate, llSTime, llETime, imDragIndexSrce, 0
                lmPSTime = -1
                mGet30Count 0, False
                lbcLines.Selected(imDragIndexSrce) = False
                imDragIndexSrce = -1
                imDragDest = -1
                pbcLbcLines_Paint
                Screen.MousePointer = vbDefault
            End If
        End If
    End If
    lbcLines.DragIcon = IconTraf!imcIconDrag.DragIcon
    lac30(Index).BackColor = lacTSTime.BackColor
    lac30(Index).ForeColor = lacTSTime.ForeColor
    lac60(Index).BackColor = lacTSTime.BackColor
    lac60(Index).ForeColor = lacTSTime.ForeColor
    lacLen(Index).BackColor = lacTSTime.BackColor
    lacLen(Index).ForeColor = lacTSTime.ForeColor
End Sub
Private Sub ckcDay_DragOver(Index As Integer, Source As control, X As Single, Y As Single, State As Integer)
    imDragDest = -1
    If imDragSrce = DRAGLINE Then
        If (State = vbLeave) Or (Not ckcDay(Index).Enabled) Then
            lbcLines.DragIcon = IconTraf!imcIconDrag.DragIcon
            lac30(Index).BackColor = lacTSTime.BackColor
            lac30(Index).ForeColor = lacTSTime.ForeColor
            lac60(Index).BackColor = lacTSTime.BackColor
            lac60(Index).ForeColor = lacTSTime.ForeColor
            lacLen(Index).BackColor = lacTSTime.BackColor
            lacLen(Index).ForeColor = lacTSTime.ForeColor
            Exit Sub
        End If
        lac30(Index).BackColor = lacTSTime.ForeColor
        lac30(Index).ForeColor = lacTSTime.BackColor
        lac60(Index).BackColor = lacTSTime.ForeColor
        lac60(Index).ForeColor = lacTSTime.BackColor
        lacLen(Index).BackColor = lacTSTime.ForeColor
        lacLen(Index).ForeColor = lacTSTime.BackColor
        lbcLines.DragIcon = IconTraf!imcIconInsert.DragIcon
        imDragDest = Index
    End If
End Sub
Private Sub ckcDay_GotFocus(Index As Integer)
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub ckcEmptyAvails_Click()
    If ckcEmptyAvails.Value = vbUnchecked Then
        ckcAdjBreaks.Value = vbUnchecked
    End If
End Sub

Private Sub ckcFAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcFAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    If imBypassAll Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    tmcLine.Enabled = False
    ilValue = Value
    If lbcVehicle.ListCount > 0 Then
        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicle.hWnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    tmcLine.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub ckcFromGame_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRet                         llRg                                                    *
'******************************************************************************************

    Dim ilValue As Integer
    Dim ilLoop As Integer

    If imBypassAll Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    tmcLine.Enabled = False
    ilValue = False
    If ckcFromGame.Value = vbChecked Then
        ilValue = True
    End If
    'If lbcFromGame.ListItems.Count > 0 Then
    '    llRg = CLng(lbcFromGame.ListItems.Count - 1) * &H10000 Or 0
    '    llRet = SendMessageByNum(lbcFromGame.hwnd, LB_SELITEMRANGE, ilValue, llRg)
    'End If
    For ilLoop = 0 To lbcFromGame.ListItems.Count - 1 Step 1
        lbcFromGame.ListItems(ilLoop + 1).Selected = ilValue
    Next ilLoop
    tmcLine.Enabled = True
    Screen.MousePointer = vbDefault

End Sub

Private Sub ckcSpotType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSpotType(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    tmcLine.Enabled = False
    tmcLine.Enabled = True
End Sub
Private Sub ckcSpotType_GotFocus(Index As Integer)
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub
Private Sub ckcTAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcTAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    If imBypassAll Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilValue = Value
    If lbcTVehicle(imFillVehType).ListCount > 0 Then
        llRg = CLng(lbcTVehicle(imFillVehType).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcTVehicle(imFillVehType).hWnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub ckcToGame_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRet                         llRg                                                    *
'******************************************************************************************

    Dim ilValue As Integer
    Dim ilLoop As Integer

    If imBypassAll Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilValue = False
    If ckcToGame.Value = vbChecked Then
        ilValue = True
    End If
    'If lbcToGame.ListItems.Count > 0 Then
    '    llRg = CLng(lbcToGame.ListItems.Count - 1) * &H10000 Or 0
    '    llRet = SendMessageByNum(lbcToGame.hwnd, LB_SELITEMRANGE, ilValue, llRg)
    'End If
    For ilLoop = 0 To lbcToGame.ListItems.Count - 1 Step 1
        lbcToGame.ListItems(ilLoop + 1).Selected = ilValue
    Next ilLoop
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcFDropDown(imFDateIndex).SelStart = 0
    edcFDropDown(imFDateIndex).SelLength = Len(edcFDropDown(imFDateIndex).Text)
    edcFDropDown(imFDateIndex).SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcFDropDown(imFDateIndex).SelStart = 0
    edcFDropDown(imFDateIndex).SelLength = Len(edcFDropDown(imFDateIndex).Text)
    edcFDropDown(imFDateIndex).SetFocus
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilRes As Integer
    Dim slSTime As String
    Dim slETime As String
    Dim llSTime As Long
    Dim llETime As Long
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCode As String
    Dim slLength As String
    Dim ilCount As Integer
    Dim ilGameNo As Integer

    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    ilFound = False
    For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
        If lbcLines.Selected(ilLoop) Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If ilFound Then
        If Not imGameVehicle Then
            slSTime = edcTDropDown(0).Text
            If (Not gValidTime(slSTime)) Or (slSTime = "") Then
                Beep
                '2/7/09: Added to handle case where focus can't be set
                On Error Resume Next
                edcTDropDown(0).SetFocus
                On Error GoTo 0
                Exit Sub
            End If
            slETime = edcTDropDown(1).Text
            If (Not gValidTime(slETime)) Or (slETime = "") Then
                Beep
                '2/7/09: Added to handle case where focus can't be set
                On Error Resume Next
                edcTDropDown(1).SetFocus
                On Error GoTo 0
                Exit Sub
            End If
            llSTime = CLng(gTimeToCurrency(slSTime, False))
            llETime = CLng(gTimeToCurrency(slETime, True))
            If llETime < llSTime Then
                Beep
                '2/7/09: Added to handle case where focus can't be set
                On Error Resume Next
                edcTDropDown(1).SetFocus
                On Error GoTo 0
                Exit Sub
            End If
            If imFillVehType = 0 Then
                If lbcTVehicle(0).ListIndex < 0 Then
                    Beep
                    If lbcTVehicle(0).ListCount > 0 Then
                        '2/7/09: Added to handle case where focus can't be set
                        On Error Resume Next
                        lbcTVehicle(0).SetFocus
                        On Error GoTo 0
                    Else
                        mVehPop
                        If lbcTVehicle(0).ListCount > 0 Then
                            '2/7/09: Added to handle case where focus can't be set
                            On Error Resume Next
                            lbcTVehicle(0).SetFocus
                            On Error GoTo 0
                        Else
                            MsgBox "No 'To' vehicles available, press Cancel and restart Traffic"
                        End If
                    End If
                    Exit Sub
                End If
            End If
        Else
            If imMixtureOfVehicles > 0 Then
                slSTime = edcTDropDown(0).Text
                If (Not gValidTime(slSTime)) Or (slSTime = "") Then
                    Beep
                    '2/7/09: Added to handle case where focus can't be set
                    On Error Resume Next
                    edcTDropDown(0).SetFocus
                    On Error GoTo 0
                    Exit Sub
                End If
                slETime = edcTDropDown(1).Text
                If (Not gValidTime(slETime)) Or (slETime = "") Then
                    Beep
                    '2/7/09: Added to handle case where focus can't be set
                    On Error Resume Next
                    edcTDropDown(1).SetFocus
                    On Error GoTo 0
                    Exit Sub
                End If
                llSTime = CLng(gTimeToCurrency(slSTime, False))
                llETime = CLng(gTimeToCurrency(slETime, True))
                If llETime < llSTime Then
                    Beep
                    '2/7/09: Added to handle case where focus can't be set
                    On Error Resume Next
                    edcTDropDown(1).SetFocus
                    On Error GoTo 0
                    Exit Sub
                End If
            Else
                llSTime = 0
                llETime = 86400
            End If
            ilCount = 0
            For ilLoop = 0 To lbcToGame.ListItems.Count - 1 Step 1
                If lbcToGame.ListItems(ilLoop + 1).Selected Then
                    ilCount = ilCount + 1
                End If
            Next ilLoop
            If ilCount <= 0 Then
                Beep
                '2/7/09: Added to handle case where focus can't be set
                On Error Resume Next
                lbcToGame.SetFocus
                On Error GoTo 0
                Exit Sub
            End If
        End If
        ilRes = MsgBox("Ok to Fill Avails", vbYesNo + vbQuestion, "Fill")
        If ilRes = vbNo Then
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        'ReDim tgSvCntSpot(0 To UBound(tgCntSpot)) As CNTSPOT
        'For ilLoop = 0 To UBound(tgCntSpot) Step 1
        '    tgSvCntSpot(ilLoop) = tgCntSpot(ilLoop)
        'Next ilLoop
        mMixCntSpot
        If Not imGameVehicle Then
            If imFillVehType = 0 Then
                slNameCode = tgUserVehicle(lbcTVehicle(0).ListIndex).sKey 'Traffic!lbcUserVehicle.List(ilVef)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                imToVefCode = Val(slCode)
                If bmFirstCallToVpfFind Then
                    imVpfIndex = gVpfFind(SpotFill, imToVefCode)
                    bmFirstCallToVpfFind = False
                Else
                    imVpfIndex = gVpfFindIndex(imToVefCode)
                End If
                If (tgVpf(imVpfIndex).iLLD(0) <> 0) Or (tgVpf(imVpfIndex).iLLD(1) <> 0) Then
                    gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLastLogDate
                Else
                    lmLastLogDate = -1
                End If
                If tgVpf(imVpfIndex).sSCompType = "T" Then
                    gUnpackLength tgVpf(imVpfIndex).iSCompLen(0), tgVpf(imVpfIndex).iSCompLen(1), "3", False, slLength
                    lmCompTime = CLng(gLengthToCurrency(slLength))
                Else
                    lmCompTime = 0&
                End If
                For ilLoop = 0 To 6 Step 1
                    If ckcDay(ilLoop).Value = vbChecked Then
                        imDays(ilLoop) = True
                    Else
                        imDays(ilLoop) = False
                    End If
                Next ilLoop
                mBookSpot lmSDate, lmEDate, llSTime, llETime, -1, 0
            Else
                For ilLoop = 0 To 6 Step 1
                    If ckcDay(ilLoop + 7).Value = vbChecked Then
                        imDays(ilLoop) = True
                    Else
                        imDays(ilLoop) = False
                    End If
                Next ilLoop
                For ilVef = 0 To lbcTVehicle(1).ListCount - 1 Step 1
                    If lbcTVehicle(1).Selected(ilVef) Then
                        slNameCode = tgUserVehicle(ilVef).sKey 'Traffic!lbcUserVehicle.List(ilVef)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        imToVefCode = Val(slCode)
                        If bmFirstCallToVpfFind Then
                            imVpfIndex = gVpfFind(SpotFill, imToVefCode)
                            bmFirstCallToVpfFind = False
                        Else
                            imVpfIndex = gVpfFindIndex(imToVefCode)
                        End If
                        If (tgVpf(imVpfIndex).iLLD(0) <> 0) Or (tgVpf(imVpfIndex).iLLD(1) <> 0) Then
                            gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLastLogDate
                        Else
                            lmLastLogDate = -1
                        End If
                        If tgVpf(imVpfIndex).sSCompType = "T" Then
                            gUnpackLength tgVpf(imVpfIndex).iSCompLen(0), tgVpf(imVpfIndex).iSCompLen(1), "3", False, slLength
                            lmCompTime = CLng(gLengthToCurrency(slLength))
                        Else
                            lmCompTime = 0&
                        End If
                        mBookSpot lmSDate, lmEDate, llSTime, llETime, -1, 0
                    End If
                Next ilVef
                'imFromVefCode = igFillVefCode
            End If
        Else
            If bmFirstCallToVpfFind Then
                imVpfIndex = gVpfFind(SpotFill, imFromVefCode)
                bmFirstCallToVpfFind = False
            Else
                imVpfIndex = gVpfFindIndex(imFromVefCode)
            End If
            For ilLoop = 0 To lbcToGame.ListItems.Count - 1 Step 1
                If lbcToGame.ListItems(ilLoop + 1).Selected Then
                    '5/5/11: Active manual contract for games
                    lmSDate = gDateValue(lbcToGame.ListItems(ilLoop + 1).SubItems(1))
                    lmEDate = lmSDate
                    ilGameNo = Val(lbcToGame.ListItems(ilLoop + 1).Text)
                    mBookSpot lmSDate, lmEDate, llSTime, llETime, -1, ilGameNo
                End If
            Next ilLoop
        End If
        ReDim tgCntSpot(0 To UBound(tgDCntSpot)) As CNTSPOT
        For ilLoop = 0 To UBound(tgDCntSpot) Step 1
            tgCntSpot(ilLoop) = tgDCntSpot(ilLoop)
        Next ilLoop
        lmPSTime = -1
        If Not imGameVehicle Then
            mGet30Count 0, False
        Else
            For ilLoop = 0 To lbcToGame.ListItems.Count - 1 Step 1
                If lbcToGame.ListItems(ilLoop + 1).Selected Then
                    ilGameNo = Val(lbcToGame.ListItems(ilLoop + 1).Text)
                    mGet30Count ilGameNo, False
                    lbcToGame.ListItems(ilLoop + 1).SubItems(2) = lac30(0).Caption   'Number of 30s
                    lbcToGame.ListItems(ilLoop + 1).SubItems(3) = lac60(0).Caption   'Number of 60s
                End If
            Next ilLoop
        End If
        tmcLine.Enabled = False
        mMakeSummary
        mLoadListBox
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub cmcDone_GotFocus()
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcFDropDown_Click(Index As Integer)
    If (imFDateIndex = 2) Or (imFDateIndex = 3) Then
        plcCalendar.Visible = Not plcCalendar.Visible
    Else
        plcTme.Visible = Not plcTme.Visible
    End If
    edcFDropDown(Index).SelStart = 0
    edcFDropDown(Index).SelLength = Len(edcFDropDown(Index).Text)
    edcFDropDown(Index).SetFocus
End Sub
Private Sub cmcFDropDown_GotFocus(Index As Integer)
    Dim slStr As String
    tmcLine.Enabled = False
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    If imFTimeIndex <> Index Then
        plcTme.Visible = False
    End If
    If imFDateIndex <> Index Then
        plcCalendar.Visible = False
        slStr = edcFDropDown(Index).Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    If Index <= 1 Then
        imFTimeIndex = Index
        imFDateIndex = -1
        plcTme.Move plcInv.Left + frcFrom.Left + edcFDropDown(Index).Left, plcInv.Top + frcFrom.Top + edcFDropDown(Index).Top + edcFDropDown(Index).height
    Else
        imFDateIndex = Index
        imFTimeIndex = -1
        plcCalendar.Move plcInv.Left + frcFrom.Left + edcFDropDown(Index).Left, plcInv.Top + frcFrom.Top + edcFDropDown(Index).Top + edcFDropDown(Index).height
    End If
    imTTimeIndex = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcTDropDown_Click(Index As Integer)
    plcTme.Visible = Not plcTme.Visible
    edcTDropDown(Index).SelStart = 0
    edcTDropDown(Index).SelLength = Len(edcTDropDown(Index).Text)
    edcTDropDown(Index).SetFocus
End Sub
Private Sub cmcTDropDown_GotFocus(Index As Integer)
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    If imTTimeIndex <> Index Then
        plcTme.Visible = False
    End If
    plcCalendar.Visible = False
    imTTimeIndex = Index
    imFTimeIndex = -1
    plcTme.Move plcInv.Left + frcTo.Left + edcTDropDown(Index).Left, plcInv.Top + frcTo.Top + edcTDropDown(Index).Top + edcTDropDown(Index).height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcTVeh_Click()
    If imFillVehType = 0 Then
        If lbcTVehicle(0).Visible Then
            plcSVeh.Visible = True
        Else
            plcSVeh.Visible = False
        End If
        lbcTVehicle(0).Visible = Not lbcTVehicle(0).Visible
    End If
End Sub
Private Sub edcFDropDown_Change(Index As Integer)
    Dim slTime As String
    Dim slDate As String
    tmcLine.Enabled = False
    If (imFDateIndex = 2) Or (imFDateIndex = 3) Then
        slDate = edcFDropDown(Index).Text
        If Not gValidDate(slDate) Then
            lacDate.Visible = False
        Else
            lacDate.Visible = True
            gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
            tmcLine.Enabled = True
        End If
    Else
        slTime = edcFDropDown(Index).Text
        If gValidTime(slTime) Then
            tmcLine.Enabled = True
        End If
    End If
End Sub
Private Sub edcFDropDown_GotFocus(Index As Integer)
    Dim slStr As String
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    If imFTimeIndex <> Index Then
        plcTme.Visible = False
    End If
    If imFDateIndex <> Index Then
        plcCalendar.Visible = False
    End If
    If Index <= 1 Then
        imFTimeIndex = Index
        imFDateIndex = -1
        plcTme.Move plcInv.Left + frcFrom.Left + edcFDropDown(Index).Left, plcInv.Top + frcFrom.Top + edcFDropDown(Index).Top + edcFDropDown(Index).height
    Else
        imFDateIndex = Index
        imFTimeIndex = -1
        plcCalendar.Move plcInv.Left + frcFrom.Left + edcFDropDown(Index).Left, plcInv.Top + frcFrom.Top + edcFDropDown(Index).Top + edcFDropDown(Index).height
        If (Index = 3) And (Trim$(edcFDropDown(Index).Text) = "") Then
            slStr = edcFDropDown(2).Text
            edcFDropDown(Index).Text = gObtainNextSunday(slStr)
        End If
    End If
    imTTimeIndex = -1
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcFDropDown_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcFDropDown_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If ActiveControl.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    If (imFDateIndex = 2) Or (imFDateIndex = 3) Then
        'Filter characters (allow only BackSpace, numbers 0 thru 9
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Else
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            ilFound = False
            For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                If KeyAscii = igLegalTime(ilLoop) Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
        gTimeOutLine KeyAscii, imcTmeOutline
    End If
End Sub
Private Sub edcFDropDown_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim slDate As String

    If (imFDateIndex = 2) Or (imFDateIndex = 3) Then
        If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
            If (Shift And vbAltMask) > 0 Then
                plcCalendar.Visible = Not plcCalendar.Visible
            Else
                slDate = edcFDropDown(Index).Text
                If gValidDate(slDate) Then
                    If KeyCode = KEYUP Then 'Up arrow
                        slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                    Else
                        slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                    End If
                    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                    edcFDropDown(Index).Text = slDate
                End If
            End If
            edcFDropDown(Index).SelStart = 0
            edcFDropDown(Index).SelLength = Len(edcFDropDown(Index).Text)
        End If
        If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
            If (Shift And vbAltMask) > 0 Then
            Else
                slDate = edcFDropDown(Index).Text
                If gValidDate(slDate) Then
                    If KeyCode = KEYLEFT Then 'Up arrow
                        slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                    Else
                        slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                    End If
                    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                    edcFDropDown(Index).Text = slDate
                End If
            End If
            edcFDropDown(Index).SelStart = 0
            edcFDropDown(Index).SelLength = Len(edcFDropDown(Index).Text)
        End If
    Else
        If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
            If (Shift And vbAltMask) > 0 Then
                plcTme.Visible = Not plcTme.Visible
            End If
            edcFDropDown(Index).SelStart = 0
            edcFDropDown(Index).SelLength = Len(edcFDropDown(Index).Text)
        End If
    End If
End Sub
Private Sub edcTDropDown_Change(Index As Integer)
    Dim slTime As String
    slTime = edcTDropDown(Index).Text
    If gValidTime(slTime) Then
        mGet30Count 0, False
    End If
End Sub
Private Sub edcTDropDown_GotFocus(Index As Integer)
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    If imTTimeIndex <> Index Then
        plcTme.Visible = False
    End If
    plcCalendar.Visible = False
    imTTimeIndex = Index
    imFTimeIndex = -1
    plcTme.Move plcInv.Left + frcTo.Left + edcTDropDown(Index).Left, plcInv.Top + frcTo.Top + edcTDropDown(Index).Top + edcTDropDown(Index).height
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcTDropDown_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTDropDown_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If ActiveControl.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        ilFound = False
        For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
            If KeyAscii = igLegalTime(ilLoop) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    gTimeOutLine KeyAscii, imcTmeOutline
End Sub
Private Sub edcTDropDown_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcTme.Visible = Not plcTme.Visible
        End If
        edcTDropDown(Index).SelStart = 0
        edcTDropDown(Index).SelLength = Len(edcTDropDown(Index).Text)
    End If
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(SPOTSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    If plcFillInv.Visible Then
'        If tgSpf.sDefFillInv = "Y" Then
'            rbcFillInv(0).Value = True
'        Else
'            rbcFillInv(1).Value = True
'        end if
        rbcFillInv(2).Value = True

    Else
        ckc10To15.Left = ckcDaysTimes.Left
        ckcDaysTimes.Top = plcAC.Top
        ckcLock.Top = plcAC.Top
        plcAC.Top = plcFillInv.Top
    End If
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
        plcCalendar.Visible = False
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    Erase tmTeam
    Erase tmGsfInfo
    Erase tgCntSpot
    Erase tgTCntSpot
    Erase tgDCntSpot
    Erase tgSCntSpot
    Erase tgSvCntSpot
    Erase tgClfSpot
    Erase tgCffSpot
    Erase tmChfAdvtExt
    Erase lmDatesClearedResv
    Erase tmVcf0
    Erase tmVcf6
    Erase tmVcf7
    Erase tmSpotMove
    ilRet = btrClose(hmRlf)
    btrDestroy hmRlf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmSxf)
    btrDestroy hmSxf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmFnf)
    btrDestroy hmFnf
    ilRet = btrClose(hmFsf)
    btrDestroy hmFsf
    btrExtClear hmStf   'Clear any previous extend operation
    ilRet = btrClose(hmStf)
    btrDestroy hmStf
    btrExtClear hmSmf   'Clear any previous extend operation
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    btrExtClear hmCgf   'Clear any previous extend operation
    ilRet = btrClose(hmCgf)
    btrDestroy hmCgf
    btrExtClear hmCff   'Clear any previous extend operation
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    btrExtClear hmCHF   'Clear any previous extend operation
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    btrExtClear hmClf   'Clear any previous extend operation
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    btrExtClear hmSsf   'Clear any previous extend operation
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    btrExtClear hmSdf   'Clear any previous extend operation
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf

    Set SpotFill = Nothing   'Remove data segment

End Sub

Private Sub lbcFromGame_Click()
    tmcLine.Enabled = False
    imBypassAll = True
    ckcFromGame.Value = vbUnchecked
    imBypassAll = False
    tmcLine.Enabled = True
End Sub

Private Sub lbcFromGame_GotFocus()
    tmcLine.Enabled = False
End Sub

Private Sub lbcLines_Click()
    pbcLbcLines_Paint
End Sub

Private Sub lbcLines_GotFocus()
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcLines_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 2) Or (imDetail = 1) Then  'Right Mouse
        Exit Sub
    End If
    fmDragX = X
    fmDragY = Y
    imDragButton = Button
    imDragType = 0
    imDragShift = Shift
    imDragSrce = DRAGLINE
    imDragIndexDest = -1
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub lbcLines_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub lbcLines_Scroll()
    pbcLbcLines_Paint
End Sub

Private Sub lbcToGame_Click()
    imBypassAll = True
    ckcToGame.Value = vbUnchecked
    imBypassAll = False
End Sub

Private Sub lbcTVehicle_Click(Index As Integer)
    imBypassAll = True
    ckcTAll.Value = vbUnchecked
    imBypassAll = False
    If imFillVehType = 0 Then
        If lbcTVehicle(0).ListIndex >= 0 Then
            smSingleName = Trim$(lbcTVehicle(0).List(lbcTVehicle(0).ListIndex))
        Else
            smSingleName = ""
        End If
        lmPETime = -1
        mGet30Count 0, True
        pbcFillVehType.Cls
        pbcFillVehType_Paint
        plcSVeh.Visible = True
        lbcTVehicle(0).Visible = False
    End If
End Sub
Private Sub lbcVehicle_Click()
    tmcLine.Enabled = False
    imBypassAll = True
    ckcFAll.Value = vbUnchecked
    imBypassAll = False
    tmcLine.Enabled = True
End Sub
Private Sub lbcVehicle_GotFocus()
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    plcTme.Visible = False
    plcCalendar.Visible = False
    tmcLine.Enabled = False
End Sub
Private Sub lbcVehicle_TopIndexChange(TopIndex As Integer)
    If tmcLine.Enabled Then
        tmcLine.Enabled = False
        tmcLine.Enabled = True
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAnyConflicts                   *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test if Ok to book spot        *
'*                                                     *
'*******************************************************
Private Function mAnyConflicts(ilAvailIndex As Integer, ilAdfCode As Integer, ilMnfComp0 As Integer, ilMnfComp1 As Integer, slType As String, ilDay As Integer) As Integer
    Dim ilSpotIndex As Integer
    Dim ilMatchComp As Integer
    Dim ilAvHour As Integer
    Dim ilHour As Integer
    Dim ilEvt As Integer
    Dim ilPSACount As Integer
    Dim ilPromoCount As Integer
    Dim tlAvail As AVAILSS
    Dim llSsfRecPos As Long
   LSet tlAvail = tmSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
    If (tlAvail.iRecType = 2) And ((slType = "S") Or (slType = "M")) Then
        'Get count- and test if max exceded
        ilAvHour = tlAvail.iTime(1) \ 256  'Obtain month
        ilPSACount = 0
        ilPromoCount = 0
        'Get start of hour
        For ilEvt = 1 To tmSsf.iCount Step 1
           LSet tlAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
            If tlAvail.iRecType = 2 Then
                ilHour = tlAvail.iTime(1) \ 256  'Obtain month
                If ilHour > ilAvHour Then
                    Exit For
                End If
                If (ilAvHour = ilHour) Then
                    For ilSpotIndex = ilEvt + 1 To ilEvt + tlAvail.iNoSpotsThis Step 1
                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                        If (tmSpot.iRank And RANKMASK) = 1050 Then 'Promo
                            ilPromoCount = ilPromoCount + 1
                        ElseIf (tmSpot.iRank And RANKMASK) = 1060 Then 'PSA
                            ilPSACount = ilPSACount + 1
                        End If
                    Next ilSpotIndex
                End If
            End If
        Next ilEvt
        If slType = "S" Then
            If ilDay <= 4 Then
                If ilPSACount >= tgVpf(imVpfIndex).iMMFPSA(ilAvHour) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            ElseIf ilDay = 5 Then
                If ilPSACount >= tgVpf(imVpfIndex).iMSAPSA(ilAvHour) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            ElseIf ilDay = 6 Then
                If ilPSACount >= tgVpf(imVpfIndex).iMSUPSA(ilAvHour) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            End If
        Else
            If ilDay <= 4 Then
                If ilPromoCount >= tgVpf(imVpfIndex).iMMFPromo(ilAvHour) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            ElseIf ilDay = 5 Then
                If ilPromoCount >= tgVpf(imVpfIndex).iMSAPromo(ilAvHour) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            ElseIf ilDay = 6 Then
                If ilPromoCount >= tgVpf(imVpfIndex).iMSUPromo(ilAvHour) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            End If
        End If
    End If
    If rbcAC(2).Value Then
        mAnyConflicts = False
        Exit Function
    ElseIf rbcAC(0).Value Then
        llSsfRecPos = 0 'Only used if preempting- Not preempting
        If Not gAdvtTest(hmSsf, tmSsf, llSsfRecPos, tmSpotMove(), imVpfIndex, lmSepLength, ilAvailIndex, tmChf.iAdfCode, tmChf.iMnfComp(0), tmChf.iMnfComp(1), 0, 0, "I", "N", imPriceLevel, True) Then
            mAnyConflicts = True
            Exit Function
        End If
        If Not gCompetitiveTest(lmCompTime, hmSsf, tmSsf, llSsfRecPos, tmSpotMove(), imVpfIndex, tmClf.iLen, tmChf.iMnfComp(0), tmChf.iMnfComp(1), ilAvailIndex, tmVcf0(), tmVcf6(), tmVcf7(), 0, 0, "I", "N", imPriceLevel, True) Then
            mAnyConflicts = True
            Exit Function
        End If
    Else
       LSet tlAvail = tmSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
        For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tlAvail.iNoSpotsThis Step 1
           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
            ilMatchComp = False
            If (ilMnfComp0 = 0) And (ilMnfComp1 = 0) And (tmSpot.iMnfComp(0) = 0) And (tmSpot.iMnfComp(1) = 0) Then
                ilMatchComp = True
            Else
                If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpot.iMnfComp(0)) Or (ilMnfComp0 = tmSpot.iMnfComp(1))) Then
                    ilMatchComp = True
                End If
                If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpot.iMnfComp(0)) Or (ilMnfComp1 = tmSpot.iMnfComp(1))) Then
                    ilMatchComp = True
                End If
            End If
            If (tmSpot.iAdfCode = ilAdfCode) And (ilMatchComp) Then
                mAnyConflicts = True
                Exit Function
            End If
            If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpot.iMnfComp(0)) Or (ilMnfComp0 = tmSpot.iMnfComp(1))) Then
                mAnyConflicts = True
                Exit Function
            ElseIf (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpot.iMnfComp(0)) Or (ilMnfComp1 = tmSpot.iMnfComp(1))) Then
                mAnyConflicts = True
                Exit Function
            End If
        Next ilSpotIndex
    End If
    mAnyConflicts = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mBookSpot                       *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get 30sec count for week       *
'*                                                     *
'*******************************************************
Private Sub mBookSpot(llSDate As Long, llEDate As Long, llSTime As Long, llETime As Long, ilCntSpotIndex As Integer, ilGameNo As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVef                                                                                 *
'******************************************************************************************

'
'   ilCntSpotIndex(I) - >=0 Fill one avail from Drag/Drop spot, -1 fill all avails from selected spots
'
    Dim slDate As String
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilEvt As Integer
    Dim ilAvEvt As Integer
    Dim llAvTime As Long
    Dim ilUnits As Integer
    Dim ilLen As Integer
    Dim ilSpotLen As Integer
    Dim ilPosition As Integer
    Dim llSsfRecPos As Long
    Dim llSdfRecPos As Long
    Dim llTime As Long
    Dim ilType As Integer
    Dim ilBkQH As Integer
    Dim slSchStatus As String
    Dim ilSpot As Integer
    Dim ilDay As Integer
    Dim ilCntAllIndex As Integer
    Dim ilStartCntAllIndex As Integer
    Dim ilTimeOk As Integer
    Dim ilSetStartLpTest As Integer   'Used to indicate that the start of the loop test must be set
                                    'after a spot is booked
    Dim slLnStartDate As String
    Dim slLnEndDate As String
    Dim slNoSpots As String
    Dim llChfCode As Long
    Dim ilAdfCode As Integer
    Dim ilVehComp As Integer
    Dim llStartDateLen As Long
    Dim llEndDateLen As Long
    Dim ilLineNo As Integer
    Dim llCntrSDate As Long
    Dim llCntrEDate As Long
    Dim ilTest As Integer
    Dim ilLenOk As Integer
    Dim ilAvailOk As Integer
    Dim llLockRecCode As Long
    Dim slUserName As String
    Dim ilPriceLevel As Integer
    Dim ilGsf As Integer
    ReDim ilDays(0 To 6) As Integer   'Day status

    If Not imGameVehicle Then
        For ilDay = 0 To 6 Step 1
            ilDays(ilDay) = imDays(ilDay)
        Next ilDay
        For llDate = llSDate To llEDate Step 1
            If (llDate < lmNowDate + 1) Or (llDate < lmLastLogDate + 1) Then
                ilDay = gWeekDayLong(llDate)
                If lmLastLogDate >= lmNowDate + 1 Then
                    If tgVpf(imVpfIndex).sMoveLLD <> "Y" Then
                        If llDate < lmLastLogDate + 1 Then
                            ilDays(ilDay) = False
                        End If
                    Else
                        If llDate < lmNowDate + 1 Then
                            ilDays(ilDay) = False
                        End If
                    End If
                Else
                    If llDate < lmNowDate + 1 Then
                        ilDays(ilDay) = False
                    End If
                End If
            End If
        Next llDate
    End If
    If Not imGameVehicle Then
        ilType = 0
        slDate = edcFDropDown(2).Text
        llCntrSDate = gDateValue(slDate)
        slDate = edcFDropDown(3).Text
        llCntrEDate = gDateValue(slDate)
        ReDim tmGsf(0 To 1) As GSF
        tmGsf(0).iGameNo = 0
        mRemoveResvSpots llSDate, llEDate
    Else
        '5/5/11: Active manual contract for games
        For ilDay = 0 To 6 Step 1
            ilDays(ilDay) = False
        Next ilDay
        ilType = ilGameNo
        imToVefCode = imFromVefCode
        '5/5/11: Active manual contract for games
        ilDay = gWeekDayLong(llSDate)
        ilDays(ilDay) = True
        ReDim tmGsf(0 To 1) As GSF
        tmGsf(0).iGameNo = ilGameNo
        If imMixtureOfVehicles > 0 Then
            slDate = edcFDropDown(2).Text
            llCntrSDate = gDateValue(slDate)
            slDate = edcFDropDown(3).Text
            llCntrEDate = gDateValue(slDate)
        End If
    End If
    ilPosition = -1
    ilBkQH = 1045   'ignore booking quarter hour since manual move
    If ilCntSpotIndex >= 0 Then
        ilSpotLen = Val(tgCntSpot(ilCntSpotIndex).sLen)
        'Allow drag to any day- ignore test
        'If imToVefCode = tgCntSpot(ilCntSpotIndex).iLnVefCode Then
        '    ilDay = gWeekDayLong(llSDate)
        '    If Not tgCntSpot(ilCntSpotIndex).iAllowedDays(ilDay) Then
        '        Exit Sub
        '    End If
        'End If
    Else
        ilCntAllIndex = LBound(tgCntSpot)
    End If
    For llDate = llSDate To llEDate Step 1
        If Not imGameVehicle Then
            ilDay = gWeekDayLong(llDate)
        End If
        If (ilDays(ilDay)) Or (ilGameNo > 0) Then   'ckcDay(ilDay).Value Then
            For ilGsf = 0 To UBound(tmGsf) - 1 Step 1
                If Not imGameVehicle Then
                    slDate = Format$(llDate, "m/d/yy")
                    llLockRecCode = gCreateLockRec(hmRlf, "S", "S", 65536 * imToVefCode + llDate, False, slUserName)
                    ilType = 0
                Else
                    llLockRecCode = gCreateLockRec(hmRlf, "S", "S", 65536 * imToVefCode + ilGameNo, False, slUserName)
                    ilType = ilGameNo
                    'If imMixtureOfVehicles > 0 Then
                        slDate = Format$(llDate, "m/d/yy")
                    'End If
                End If
                If llLockRecCode > 0 Then
                    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    If Not imGameVehicle Then
                        gPackDate slDate, ilLogDate0, ilLogDate1
                        tmSsfSrchKey.iType = ilType
                        tmSsfSrchKey.iVefCode = imToVefCode
                        tmSsfSrchKey.iDate(0) = ilLogDate0
                        tmSsfSrchKey.iDate(1) = ilLogDate1
                        tmSsfSrchKey.iStartTime(0) = 0
                        tmSsfSrchKey.iStartTime(1) = 0
                        ilRet = gSSFGetEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                    Else
                        'tmSsfSrchKey1.iVefCode = imToVefCode
                        'tmSsfSrchKey1.iType = ilType
                        'ilRet = gSSFGetEqualKey1(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                        gPackDate slDate, ilLogDate0, ilLogDate1
                        tmSsfSrchKey2.iVefCode = imToVefCode
                        tmSsfSrchKey2.iDate(0) = ilLogDate0
                        tmSsfSrchKey2.iDate(1) = ilLogDate1
                        ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = imToVefCode)
                            If (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1) Then
                                If tmSsf.iType = ilGameNo Then
                                    Exit Do
                                End If
                            End If
                            imSsfRecLen = Len(tmSsf)
                            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        Loop
                    End If
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = imToVefCode) And (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1)
                        ilRet = gSSFGetPosition(hmSsf, llSsfRecPos)
                        ilEvt = 1
                        Do While (ilEvt <= tmSsf.iCount) And (tmSsf.iCount < UBound(tmSsf.tPas))
                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                            'If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                            If ((tmAvail.iRecType = 2) Or (tmAvail.iRecType = 8) Or (tmAvail.iRecType = 9)) And (((tmAvail.iAvInfo And SSLOCK) <> SSLOCK) Or (((tmAvail.iAvInfo And SSLOCK) = SSLOCK) And (ckcLock.Value = vbUnchecked))) Then
                                gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                If llTime >= llETime Then
                                    Exit Do
                                End If
                                ilAvailOk = True
                                If (ckcEmptyAvails.Value = vbChecked) And (tmAvail.iNoSpotsThis <= 0) Then
                                    ilAvailOk = False
                                End If
                                If (ckcAdjBreaks.Value = vbChecked) And (Not ilAvailOk) Then
                                    ilAvEvt = ilEvt
                                    llAvTime = llTime + tmAvail.iLen
                                    ilEvt = ilEvt + 1
                                    Do While (ilEvt <= tmSsf.iCount) And (tmSsf.iCount < UBound(tmSsf.tPas))
                                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                        'If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                                        If ((tmAvail.iRecType = 2) Or (tmAvail.iRecType = 8) Or (tmAvail.iRecType = 9)) And (((tmAvail.iAvInfo And SSLOCK) <> SSLOCK) Or (((tmAvail.iAvInfo And SSLOCK) = SSLOCK) And (ckcLock.Value = vbUnchecked))) Then
                                            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                            If llTime = llAvTime Then
                                                If tmAvail.iNoSpotsThis > 0 Then
                                                    ilAvailOk = True
                                                End If
                                                Exit Do
                                            End If
                                        End If
                                        ilEvt = ilEvt + 1
                                    Loop
                                    ilEvt = ilAvEvt
                                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llAvTime
                                    ilEvt = ilEvt - 1
                                    Do While ilEvt >= 1
                                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                        'If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                                        If ((tmAvail.iRecType = 2) Or (tmAvail.iRecType = 8) Or (tmAvail.iRecType = 9)) And (((tmAvail.iAvInfo And SSLOCK) <> SSLOCK) Or (((tmAvail.iAvInfo And SSLOCK) = SSLOCK) And (ckcLock.Value = vbUnchecked))) Then
                                            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                            llTime = llTime + tmAvail.iLen
                                            If llTime = llAvTime Then
                                                If tmAvail.iNoSpotsThis > 0 Then
                                                    ilAvailOk = True
                                                End If
                                                Exit Do
                                            End If
                                        End If
                                        ilEvt = ilEvt - 1
                                    Loop
                                    ilEvt = ilAvEvt
                                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                End If
                                If (llTime >= llSTime) And ilAvailOk Then
                                    ilAvEvt = ilEvt
                                    'Test if within selected times
                                    ilLen = tmAvail.iLen
                                    ilUnits = tmAvail.iAvInfo And &H1F
                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                        ilEvt = ilEvt + 1
                                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                        If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                            ilUnits = ilUnits - 1
                                            ilLen = ilLen - (tmSpot.iPosLen And &HFFF)
                                        End If
                                    Next ilSpot
                                    If ilCntSpotIndex >= 0 Then
                                        'If (ilUnits > 0) And (Val(tgCntSpot(ilCntSpotIndex).sLen) <= ilLen) And (llTime >= tgCntSpot(ilCntSpotIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntSpotIndex).lAllowedETime) Then
                                        If (ilUnits > 0) And (Val(tgCntSpot(ilCntSpotIndex).sLen) <= ilLen) Then
                                            ilLenOk = False
                                            If (Val(tgCntSpot(ilCntSpotIndex).sLen) = 30) Or (Val(tgCntSpot(ilCntSpotIndex).sLen) = 60) Or (Val(tgCntSpot(ilCntSpotIndex).sLen) = tmAvail.iLen) Then
                                                ilLenOk = True
                                            End If
                                            'If (Val(tgCntSpot(ilCntSpotIndex).sLen) = tmAvail.iLen) Then
                                            '    ilLenOk = True
                                            'End If
                                            'Book if spot is same length as remaining length
                                            If (Val(tgCntSpot(ilCntSpotIndex).sLen) = ilLen) Then
                                                ilLenOk = True
                                            End If
                                            If (Val(tgCntSpot(ilCntSpotIndex).sLen) = 10) And (tmAvail.iLen = 15) And (ckc10To15.Value = vbChecked) Then
                                                ilLenOk = True
                                            End If
                                            If ilLenOk Then
                                                'If tgCntSpot(ilCntSpotIndex).sType = "S" Then
                                                '    If (llTime >= tgCntSpot(ilCntSpotIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntSpotIndex).lAllowedETime) Then
                                                '        ilTimeOk = True
                                                '    Else
                                                '        ilTimeOk = False
                                                '    End If
                                                'ElseIf tgCntSpot(ilCntSpotIndex).sType = "M" Then
                                                '    If (llTime >= tgCntSpot(ilCntSpotIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntSpotIndex).lAllowedETime) Then
                                                '        ilTimeOk = True
                                                '    Else
                                                '        ilTimeOk = False
                                                '    End If
                                                'ElseIf tgCntSpot(ilCntSpotIndex).sType = "Q" Then
                                                '    If (llTime >= tgCntSpot(ilCntSpotIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntSpotIndex).lAllowedETime) Then
                                                '        ilTimeOk = True
                                                '    Else
                                                '        ilTimeOk = False
                                                '    End If
                                                'Else
                                                '    ilTimeOk = True
                                                'End If
                                                If ckcDaysTimes.Value = vbChecked Then
                                                    'If (llTime >= tgCntSpot(ilCntSpotIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntSpotIndex).lAllowedETime) Then
                                                    '    If tgCntSpot(ilCntSpotIndex).iAllowedDays(ilDay) Then
                                                    '        ilTimeOk = True
                                                    '    Else
                                                    '        ilTimeOk = False
                                                    '    End If
                                                    'Else
                                                    '    ilTimeOk = False
                                                    'End If
                                                    ilTimeOk = False
                                                    For ilTest = LBound(tgCntSpot(ilCntSpotIndex).lAllowedSTime) To UBound(tgCntSpot(ilCntSpotIndex).lAllowedSTime) Step 1
                                                        If tgCntSpot(ilCntSpotIndex).lAllowedSTime(ilTest) <> -1 Then
                                                            If (llTime >= tgCntSpot(ilCntSpotIndex).lAllowedSTime(ilTest)) And (llTime <= tgCntSpot(ilCntSpotIndex).lAllowedETime(ilTest)) Then
                                                                If tgCntSpot(ilCntSpotIndex).iAllowedDays(ilDay) Then
                                                                    ilTimeOk = True
                                                                End If
                                                            End If
                                                        End If
                                                    Next ilTest
                                                Else
                                                    ilTimeOk = True
                                                End If
                                                If ilTimeOk Then
                                                    If (tmAvail.iRecType = 2) Or ((tmAvail.iRecType = 8) And (tgCntSpot(ilCntSpotIndex).sType = "S")) Or ((tmAvail.iRecType = 9) And (tgCntSpot(ilCntSpotIndex).sType = "M")) Then
                                                        'Moved test so that imPriceLevel set
                                                        'If (rbcAC(0).Value) Or (ckcAvailNames.Value = vbChecked) Then
                                                        ''5/5/11: Active manual contract mode fpr games
                                                        'If Not imGameVehicle Then
                                                            ilRet = mReadChfClfRdfCffRec(tgCntSpot(ilCntSpotIndex).lChfCode, tgCntSpot(ilCntSpotIndex).iLineNo, tgCntSpot(ilCntSpotIndex).lFsfCode, tgCntSpot(ilCntSpotIndex).iGameNo, slDate, slLnStartDate, slLnEndDate, slNoSpots)
                                                        'Else
                                                        '    ilRet = mReadChfClfRdfCffRec(tgCntSpot(ilCntSpotIndex).lChfCode, tgCntSpot(ilCntSpotIndex).iLineNo, tgCntSpot(ilCntSpotIndex).lFsfCode, ilGameNo, slDate, slLnStartDate, slLnEndDate, slNoSpots)
                                                        'End If
                                                        If ilRet Then
                                                        '5/5/11
                                                            llChfCode = 0   'Force to recompute value so ilBkQH gets commputed
                                                            ''gGetLineSchParameters hmSsf, tgSsf(), lgSsfDate(), lgSsfRecPos(), gDateValue(slDate), imToVefCode, tmChf.iAdfCode, ilGameNo, tmICff(), tmClf, tmRdf, lmSepLength, llStartDateLen, llEndDateLen, llChfCode, ilLineNo, ilAdfCode, ilVehComp, imHour(), imDay(), imQH(), imAHour(), imADay(), imAQH(), lmTBStartTime(), lmTBEndTime(), 0, imSkip(), tmChf.sType, tmChf.iPctTrade, imBkQH, False, imPriceLevel
                                                            'gGetLineSchParameters hmSsf, tgSsf(), lgSsfDate(), lgSsfRecPos(), gDateValue(slDate), imToVefCode, tmChf.iadfCode, tgCntSpot(ilCntSpotIndex).iGameNo, tmICff(), tmClf, tmRdf, lmSepLength, llStartDateLen, llEndDateLen, llChfCode, ilLineNo, ilAdfCode, ilVehComp, imHour(), imDay(), imQH(), imAHour(), imADay(), imAQH(), lmTBStartTime(), lmTBEndTime(), 0, imSkip(), tmChf.sType, tmChf.iPctTrade, imBkQH, False, imPriceLevel, False
                                                            If tgCntSpot(ilCntSpotIndex).lSepLength <= 0 Then
                                                                gGetLineSchParameters hmSsf, tgSsf(), lgSsfDate(), lgSsfRecPos(), gDateValue(slDate), imToVefCode, tmChf.iAdfCode, tgCntSpot(ilCntSpotIndex).iGameNo, tmICff(), tmClf, tmRdf, lmSepLength, llStartDateLen, llEndDateLen, llChfCode, ilLineNo, ilAdfCode, ilVehComp, imHour(), imDay(), imQH(), imAHour(), imADay(), imAQH(), lmTBStartTime(), lmTBEndTime(), 0, imSkip(), tmChf.sType, tmChf.iPctTrade, imBkQH, False, imPriceLevel, False
                                                                tgCntSpot(ilCntSpotIndex).lSepLength = lmSepLength
                                                            Else
                                                                lmSepLength = tgCntSpot(ilCntSpotIndex).lSepLength
                                                            End If
                                                            If (rbcAC(0).Value) Or (ckcAvailNames.Value = vbChecked) Then
                                                                If ckcAvailNames.Value = vbChecked Then
                                                                    If tmRdf.sInOut = "I" Then
                                                                        If (tmAvail.ianfCode <> tmRdf.ianfCode) Then
                                                                            ilRet = False
                                                                        End If
                                                                    ElseIf tmRdf.sInOut = "O" Then
                                                                        If (tmAvail.ianfCode = tmRdf.ianfCode) Then
                                                                            ilRet = False
                                                                        End If
                                                                    Else    'Book into any avail which allows sustaining spots
                                                                        If (tmAvail.iAvInfo And SSSUSTAINING) <> SSSUSTAINING Then
                                                                            ilRet = False
                                                                        End If
                                                                    End If
                                                                End If
                                                            Else
                                                                ilRet = True
                                                            End If
                                                        '5/5/11: Active manual contract mode for games
                                                        End If
                                                        '5/5/11
                                                        If ilRet Then
                                                            If Not mAnyConflicts(ilAvEvt, tgCntSpot(ilCntSpotIndex).iAdfCode, tgCntSpot(ilCntSpotIndex).iMnfComp0, tgCntSpot(ilCntSpotIndex).iMnfComp1, tgCntSpot(ilCntSpotIndex).sType, ilDay) Then
                                                                If (rbcAC(1).Value) Or (rbcAC(2).Value) Then
                                                                    ilRet = mReadChfClfRdfRec(tgCntSpot(ilCntSpotIndex).lChfCode, tgCntSpot(ilCntSpotIndex).iLineNo, tgCntSpot(ilCntSpotIndex).lFsfCode)
                                                                Else
                                                                    ilRet = True
                                                                End If
                                                                'Book spots
                                                                'mMakeUnschSpot tgCntSpot(ilCntSpotIndex).lChfCode, tgCntSpot(ilCntSpotIndex).iAdfCode, tgCntSpot(ilCntSpotIndex).iLineNo, slDate, imToVefCode, ilSpotLen, llSdfRecPos
                                                                If ilRet Then
                                                                    ilRet = btrBeginTrans(hmSdf, 1000)
                                                                    If ilRet <> BTRV_ERR_NONE Then
                                                                        ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                                        Exit Sub
                                                                    End If
                                                                    'ilRet = mMakeUnschSpot(tgCntSpot(ilCntSpotIndex).lChfCode, tgCntSpot(ilCntSpotIndex).iAdfCode, tgCntSpot(ilCntSpotIndex).iLineNo, tgCntSpot(ilCntSpotIndex).lFsfCode, ilGameNo, slDate, tgCntSpot(ilCntSpotIndex).iLnVefCode, ilSpotLen, llSdfRecPos)
                                                                    ilRet = mMakeUnschSpot(tgCntSpot(ilCntSpotIndex).lChfCode, tgCntSpot(ilCntSpotIndex).iAdfCode, tgCntSpot(ilCntSpotIndex).iLineNo, tgCntSpot(ilCntSpotIndex).lFsfCode, tgCntSpot(ilCntSpotIndex).iGameNo, slDate, tgCntSpot(ilCntSpotIndex).iLnVefCode, ilSpotLen, llSdfRecPos)
                                                                    If ilRet <> BTRV_ERR_NONE Then
                                                                        If ilRet >= 30000 Then
                                                                            ilRet = csiHandleValue(0, 7)
                                                                        End If
                                                                        ilCRet = btrAbortTrans(hmSdf)
                                                                        ilCRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                                        ilRet = MsgBox("Task could not be completed successfully because of " & str$(ilRet) & ", Redo Task", vbOKOnly + vbExclamation, "Spot")
                                                                        Exit Sub
                                                                    End If
                                                                    If imToVefCode = tgCntSpot(ilCntSpotIndex).iLnVefCode Then
                                                                        'If (llTime >= tgCntSpot(ilCntSpotIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntSpotIndex).lAllowedETime) And tgCntSpot(ilCntSpotIndex).iAllowedDays(ilDay) Then
                                                                        '    'slSchStatus = "S"
                                                                        '    If (llCntrSDate <> lmSDate) Or (llCntrEDate <> lmEDate) Then
                                                                        '        slSchStatus = "O"
                                                                        '    Else
                                                                        '        slSchStatus = "S"
                                                                        '    End If
                                                                        'Else
                                                                        '    slSchStatus = "O"
                                                                        'End If
                                                                        slSchStatus = "O"
                                                                        For ilTest = LBound(tgCntSpot(ilCntSpotIndex).lAllowedSTime) To UBound(tgCntSpot(ilCntSpotIndex).lAllowedSTime) Step 1
                                                                            If tgCntSpot(ilCntSpotIndex).lAllowedSTime(ilTest) <> -1 Then
                                                                                If (llTime >= tgCntSpot(ilCntSpotIndex).lAllowedSTime(ilTest)) And (llTime <= tgCntSpot(ilCntSpotIndex).lAllowedETime(ilTest)) And tgCntSpot(ilCntSpotIndex).iAllowedDays(ilDay) Then
                                                                                    If (llCntrSDate = lmSDate) And (llCntrEDate = lmEDate) Then
                                                                                        slSchStatus = "S"
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        Next ilTest
                                                                    Else
                                                                        slSchStatus = "O"
                                                                    End If
                                                                    If tgCntSpot(ilCntSpotIndex).sType = "S" Then
                                                                        ilBkQH = PSARANK    '1060
                                                                    ElseIf tgCntSpot(ilCntSpotIndex).sType = "M" Then
                                                                        ilBkQH = PROMORANK  '1050
                                                                    ElseIf tgCntSpot(ilCntSpotIndex).sType = "Q" Then
                                                                        ilBkQH = PERINQUIRYRANK '1030
                                                                    ElseIf tgCntSpot(ilCntSpotIndex).sType = "T" Then
                                                                        ilBkQH = REMNANTRANK    '1020
                                                                    ElseIf tgCntSpot(ilCntSpotIndex).sType = "R" Then
                                                                        ilBkQH = DIRECTRESPONSERANK '1010
                                                                    Else
                                                                        ilBkQH = EXTRARANK  '1045
                                                                    End If
                                                                    ilPriceLevel = 0
                                                                    'BookSpot Re-Read Ssf so handle is correct
                                                                    ilRet = gBookSpot(slSchStatus, hmSdf, tmSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf, llSsfRecPos, ilAvEvt, ilPosition, tmChf, tmClf, tmRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, ilPriceLevel, False, hmSxf, hmGsf)
                                                                    If ilRet Then
                                                                        If imToVefCode = igFillVefCode Then
                                                                            igSpotFillReturn = 1
                                                                        End If
                                                                        cmcCancel.Caption = "&Done"
                                                                        'mMakeTracer llSdfRecPos, "S"
                                                                        ilRet = gMakeTracer(hmSdf, tmSdf, llSdfRecPos, hmStf, lmLastLogDate, "S", "M", tmSdf.iRotNo, hmGsf)
                                                                        If ilRet Then
                                                                            tgCntSpot(ilCntSpotIndex).iNoTimesUsed = tgCntSpot(ilCntSpotIndex).iNoTimesUsed + 1
                                                                            tgCntSpot(ilCntSpotIndex).iNoESpots = tgCntSpot(ilCntSpotIndex).iNoESpots + 1
                                                                            ilRet = btrEndTrans(hmSdf)
                                                                            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                                            Exit Sub
                                                                        Else
                                                                            ilCRet = btrAbortTrans(hmSdf)
                                                                            ilCRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                                            'ilRet = MsgBox("Task could not be completed successfully because of " & Str$(ilRet) & ", Redo Task", vbOkOnly + vbExclamation, "Spot")
                                                                            ilRet = MsgBox("Task could not be completed successfully because of " & str$(igBtrError) & ": " & sgErrLoc & ", Redo Task", vbOKOnly + vbExclamation, "Spot")
                                                                            Exit Sub
                                                                        End If
                                                                    End If
                                                                    ilCRet = btrAbortTrans(hmSdf)
                                                                    ilCRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                                    'ilRet = MsgBox("Task could not be completed successfully because of " & Str$(ilRet) & ", Redo Task", vbOkOnly + vbExclamation, "Spot")
                                                                    ilRet = MsgBox("Task could not be completed successfully because of " & str$(igBtrError) & ": " & sgErrLoc & ", Redo Task", vbOKOnly + vbExclamation, "Spot")
                                                                    Exit Sub
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        ilStartCntAllIndex = ilCntAllIndex
                                        ilSetStartLpTest = False
                                        Do While (ilUnits > 0) And (ilLen > 0)
                                            'If tgCntSpot(ilCntAllIndex).sType = "S" Then
                                            '    If (llTime >= tgCntSpot(ilCntAllIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntAllIndex).lAllowedETime) Then
                                            '        ilTimeOk = True
                                            '    Else
                                            '        ilTimeOk = False
                                            '    End If
                                            'ElseIf tgCntSpot(ilCntAllIndex).sType = "M" Then
                                            '    If (llTime >= tgCntSpot(ilCntAllIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntAllIndex).lAllowedETime) Then
                                            '        ilTimeOk = True
                                            '    Else
                                            '        ilTimeOk = False
                                            '    End If
                                            'ElseIf tgCntSpot(ilCntAllIndex).sType = "Q" Then
                                            '    If (llTime >= tgCntSpot(ilCntAllIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntAllIndex).lAllowedETime) Then
                                            '        ilTimeOk = True
                                            '    Else
                                            '        ilTimeOk = False
                                            '    End If
                                            'Else
                                            '    ilTimeOk = True
                                            'End If
                                            If (ilCntAllIndex < LBound(tgCntSpot)) Or (ilCntAllIndex >= UBound(tgCntSpot)) Then
                                                Exit Sub
                                            End If
                                            If ckcDaysTimes.Value = vbChecked Then
                                                'If (llTime >= tgCntSpot(ilCntAllIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntAllIndex).lAllowedETime) Then
                                                '    If tgCntSpot(ilCntAllIndex).iAllowedDays(ilDay) Then
                                                '        ilTimeOk = True
                                                '    Else
                                                '        ilTimeOk = False
                                                '    End If
                                                'Else
                                                '    ilTimeOk = False
                                                'End If
                                                ilTimeOk = False
                                                For ilTest = LBound(tgCntSpot(ilCntAllIndex).lAllowedSTime) To UBound(tgCntSpot(ilCntAllIndex).lAllowedSTime) Step 1
                                                    If tgCntSpot(ilCntAllIndex).lAllowedSTime(ilTest) <> -1 Then
                                                        If (llTime >= tgCntSpot(ilCntAllIndex).lAllowedSTime(ilTest)) And (llTime <= tgCntSpot(ilCntAllIndex).lAllowedETime(ilTest)) Then
                                                            If tgCntSpot(ilCntAllIndex).iAllowedDays(ilDay) Then
                                                                ilTimeOk = True
                                                            End If
                                                        End If
                                                    End If
                                                Next ilTest
                                            Else
                                                ilTimeOk = True
                                            End If
                                            If ilTimeOk Then
                                                If (tmAvail.iRecType = 2) Or ((tmAvail.iRecType = 8) And (tgCntSpot(ilCntAllIndex).sType = "S")) Or ((tmAvail.iRecType = 9) And (tgCntSpot(ilCntAllIndex).sType = "M")) Then
                                                    'If (Val(tgCntSpot(ilCntAllIndex).sLen) = 30) Or (Val(tgCntSpot(ilCntAllIndex).sLen) = 60) Or (Val(tgCntSpot(ilCntAllIndex).sLen) = tmAvail.iLen) Then
                                                    ilLenOk = False
                                                    If (Val(tgCntSpot(ilCntAllIndex).sLen) = 30) Or (Val(tgCntSpot(ilCntAllIndex).sLen) = 60) Or (Val(tgCntSpot(ilCntAllIndex).sLen) = tmAvail.iLen) Then
                                                        ilLenOk = True
                                                    End If
                                                    'If (Val(tgCntSpot(ilCntAllIndex).sLen) = tmAvail.iLen) Then
                                                    '    ilLenOk = True
                                                    'End If
                                                    'Book if spot is same length as remaining length
                                                    If (Val(tgCntSpot(ilCntAllIndex).sLen) = ilLen) Then
                                                        ilLenOk = True
                                                    End If
                                                    If (Val(tgCntSpot(ilCntAllIndex).sLen) = 10) And (tmAvail.iLen = 15) And (ckc10To15.Value = vbChecked) Then
                                                        ilLenOk = True
                                                    End If
                                                    If ilLenOk Then
                                                        If (Val(tgCntSpot(ilCntAllIndex).sLen) <= ilLen) And ((imToVefCode <> tgCntSpot(ilCntAllIndex).iLnVefCode) Or (tgCntSpot(ilCntAllIndex).iAllowedDays(ilDay))) Then
                                                            'Moved test so that imPriceLevel set
                                                            'If (rbcAC(0).Value) Or (ckcAvailNames.Value = vbChecked) Then
                                                            ''5/5/11: Active manual contract fill for games
                                                            'If Not imGameVehicle Then
                                                                ilRet = mReadChfClfRdfCffRec(tgCntSpot(ilCntAllIndex).lChfCode, tgCntSpot(ilCntAllIndex).iLineNo, tgCntSpot(ilCntAllIndex).lFsfCode, tgCntSpot(ilCntAllIndex).iGameNo, slDate, slLnStartDate, slLnEndDate, slNoSpots)
                                                            'Else
                                                            '    ilRet = mReadChfClfRdfCffRec(tgCntSpot(ilCntAllIndex).lChfCode, tgCntSpot(ilCntAllIndex).iLineNo, tgCntSpot(ilCntAllIndex).lFsfCode, ilGameNo, slDate, slLnStartDate, slLnEndDate, slNoSpots)
                                                            'End If
                                                            If ilRet Then
                                                            '5/5/11
                                                                llChfCode = 0   'Force to recompute value so ilBkQH gets commputed
                                                                'gGetLineSchParameters hmSsf, tgSsf(), lgSsfDate(), lgSsfRecPos(), gDateValue(slDate), imToVefCode, tmChf.iadfCode, ilType, tmICff(), tmClf, tmRdf, lmSepLength, llStartDateLen, llEndDateLen, llChfCode, ilLineNo, ilAdfCode, ilVehComp, imHour(), imDay(), imQH(), imAHour(), imADay(), imAQH(), lmTBStartTime(), lmTBEndTime(), 0, imSkip(), tmChf.sType, tmChf.iPctTrade, imBkQH, False, imPriceLevel, False
                                                                If tgCntSpot(ilCntAllIndex).lSepLength <= 0 Then
                                                                    gGetLineSchParameters hmSsf, tgSsf(), lgSsfDate(), lgSsfRecPos(), gDateValue(slDate), imToVefCode, tmChf.iAdfCode, ilType, tmICff(), tmClf, tmRdf, lmSepLength, llStartDateLen, llEndDateLen, llChfCode, ilLineNo, ilAdfCode, ilVehComp, imHour(), imDay(), imQH(), imAHour(), imADay(), imAQH(), lmTBStartTime(), lmTBEndTime(), 0, imSkip(), tmChf.sType, tmChf.iPctTrade, imBkQH, False, imPriceLevel, False
                                                                    tgCntSpot(ilCntAllIndex).lSepLength = lmSepLength
                                                                Else
                                                                    lmSepLength = tgCntSpot(ilCntAllIndex).lSepLength
                                                                End If
                                                                If (rbcAC(0).Value) Or (ckcAvailNames.Value = vbChecked) Then
                                                                    If ckcAvailNames.Value = vbChecked Then
                                                                        If tmRdf.sInOut = "I" Then
                                                                            If (tmAvail.ianfCode <> tmRdf.ianfCode) Then
                                                                                ilRet = False
                                                                            End If
                                                                        ElseIf tmRdf.sInOut = "O" Then
                                                                            If (tmAvail.ianfCode = tmRdf.ianfCode) Then
                                                                                ilRet = False
                                                                            End If
                                                                        Else    'Book into any avail which allows sustaining spots
                                                                            If (tmAvail.iAvInfo And SSSUSTAINING) <> SSSUSTAINING Then
                                                                                ilRet = False
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Else
                                                                    ilRet = True
                                                                End If
                                                            '5/5/11: Active manual contract for game
                                                            End If
                                                            '5/5/11
                                                            If ilRet Then
                                                                If Not mAnyConflicts(ilAvEvt, tgCntSpot(ilCntAllIndex).iAdfCode, tgCntSpot(ilCntAllIndex).iMnfComp0, tgCntSpot(ilCntAllIndex).iMnfComp1, tgCntSpot(ilCntAllIndex).sType, ilDay) Then
                                                                    If (rbcAC(1).Value) Or (rbcAC(2).Value) Then
                                                                        ilRet = mReadChfClfRdfRec(tgCntSpot(ilCntAllIndex).lChfCode, tgCntSpot(ilCntAllIndex).iLineNo, tgCntSpot(ilCntAllIndex).lFsfCode)
                                                                    Else
                                                                        ilRet = True
                                                                    End If
                                                                    If ilRet Then
                                                                        ilRet = btrBeginTrans(hmSdf, 1000)
                                                                        If ilRet <> BTRV_ERR_NONE Then
                                                                            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                                            Exit Sub
                                                                        End If
                                                                        'Book spots
                                                                        ilSpotLen = Val(tgCntSpot(ilCntAllIndex).sLen)
                                                                        'mMakeUnschSpot tgCntSpot(ilCntAllIndex).lChfCode, tgCntSpot(ilCntAllIndex).iAdfCode, tgCntSpot(ilCntAllIndex).iLineNo, slDate, imToVefCode, ilSpotLen, llSdfRecPos
                                                                        'ilRet = mMakeUnschSpot(tgCntSpot(ilCntAllIndex).lChfCode, tgCntSpot(ilCntAllIndex).iAdfCode, tgCntSpot(ilCntAllIndex).iLineNo, tgCntSpot(ilCntAllIndex).lFsfCode, ilGameNo, slDate, tgCntSpot(ilCntAllIndex).iLnVefCode, ilSpotLen, llSdfRecPos)
                                                                        ilRet = mMakeUnschSpot(tgCntSpot(ilCntAllIndex).lChfCode, tgCntSpot(ilCntAllIndex).iAdfCode, tgCntSpot(ilCntAllIndex).iLineNo, tgCntSpot(ilCntAllIndex).lFsfCode, tgCntSpot(ilCntAllIndex).iGameNo, slDate, tgCntSpot(ilCntAllIndex).iLnVefCode, ilSpotLen, llSdfRecPos)
                                                                        If ilRet <> BTRV_ERR_NONE Then
                                                                            If ilRet >= 30000 Then
                                                                                ilRet = csiHandleValue(0, 7)
                                                                            End If
                                                                            ilCRet = btrAbortTrans(hmSdf)
                                                                            ilCRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                                            ilRet = MsgBox("Task could not be completed successfully because of " & str$(ilRet) & ", Redo Task", vbOKOnly + vbExclamation, "Spot")
                                                                            Exit Sub
                                                                        End If
                                                                        If imToVefCode = tgCntSpot(ilCntAllIndex).iLnVefCode Then
                                                                            'If (llTime >= tgCntSpot(ilCntAllIndex).lAllowedSTime) And (llTime <= tgCntSpot(ilCntAllIndex).lAllowedETime) And tgCntSpot(ilCntAllIndex).iAllowedDays(ilDay) Then
                                                                            '    'slSchStatus = "S"
                                                                            '    If (llCntrSDate <> lmSDate) Or (llCntrEDate <> lmEDate) Then
                                                                            '        slSchStatus = "O"
                                                                            '    Else
                                                                            '        slSchStatus = "S"
                                                                            '    End If
                                                                            'Else
                                                                            '    slSchStatus = "O"
                                                                            'End If
                                                                            slSchStatus = "O"
                                                                            For ilTest = LBound(tgCntSpot(ilCntAllIndex).lAllowedSTime) To UBound(tgCntSpot(ilCntAllIndex).lAllowedSTime) Step 1
                                                                                If tgCntSpot(ilCntAllIndex).lAllowedSTime(ilTest) <> -1 Then
                                                                                    If (llTime >= tgCntSpot(ilCntAllIndex).lAllowedSTime(ilTest)) And (llTime <= tgCntSpot(ilCntAllIndex).lAllowedETime(ilTest)) And tgCntSpot(ilCntAllIndex).iAllowedDays(ilDay) Then
                                                                                        If (llCntrSDate = lmSDate) And (llCntrEDate = lmEDate) Then
                                                                                            slSchStatus = "S"
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            Next ilTest
                                                                        Else
                                                                            slSchStatus = "O"
                                                                        End If
                                                                        If tgCntSpot(ilCntAllIndex).sType = "S" Then    'PSA
                                                                            ilBkQH = PSARANK    '1060
                                                                        ElseIf tgCntSpot(ilCntAllIndex).sType = "M" Then    'Promo
                                                                            ilBkQH = PROMORANK  '1050
                                                                        ElseIf tgCntSpot(ilCntAllIndex).sType = "Q" Then    'Per Inquiry
                                                                            ilBkQH = PERINQUIRYRANK '1030
                                                                        ElseIf tgCntSpot(ilCntAllIndex).sType = "T" Then    'Remnant
                                                                            ilBkQH = REMNANTRANK    '1020
                                                                        ElseIf tgCntSpot(ilCntAllIndex).sType = "R" Then    'DR
                                                                            ilBkQH = DIRECTRESPONSERANK '1010
                                                                        Else
                                                                            ilBkQH = EXTRARANK  '1045   'Extra
                                                                        End If
                                                                        ilPriceLevel = 0
                                                                        'BookSpot Re-Read Ssf so handle is correct
                                                                        ilRet = gBookSpot(slSchStatus, hmSdf, tmSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf, llSsfRecPos, ilAvEvt, ilPosition, tmChf, tmClf, tmRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, ilPriceLevel, False, hmSxf, hmGsf)
                                                                        If ilRet Then
                                                                            If imToVefCode = igFillVefCode Then
                                                                                igSpotFillReturn = 1
                                                                            End If
                                                                            cmcCancel.Caption = "&Done"
                                                                            'mMakeTracer llSdfRecPos, "S"
                                                                            ilRet = gMakeTracer(hmSdf, tmSdf, llSdfRecPos, hmStf, lmLastLogDate, "S", "M", tmSdf.iRotNo, hmGsf)
                                                                            If ilRet Then
                                                                                ilUnits = ilUnits - 1
                                                                                ilLen = ilLen - Val(tgCntSpot(ilCntAllIndex).sLen)
                                                                                tgCntSpot(ilCntAllIndex).iNoTimesUsed = tgCntSpot(ilCntAllIndex).iNoTimesUsed + 1
                                                                                tgDCntSpot(tgCntSpot(ilCntAllIndex).iUpdateIndex).iNoESpots = tgDCntSpot(tgCntSpot(ilCntAllIndex).iUpdateIndex).iNoESpots + 1
                                                                                ilSetStartLpTest = True
                                                                                ilRet = btrEndTrans(hmSdf)
                                                                            Else
                                                                                ilCRet = btrAbortTrans(hmSdf)
                                                                                ilCRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                                                'ilRet = MsgBox("Task could not be completed successfully, Redo Task", vbOkOnly + vbExclamation, "Spot")
                                                                                ilRet = MsgBox("Task could not be completed successfully because of " & str$(igBtrError) & ": " & sgErrLoc & ", Redo Task", vbOKOnly + vbExclamation, "Spot")
                                                                                Exit Sub
                                                                            End If
                                                                        Else
                                                                            ilCRet = btrAbortTrans(hmSdf)
                                                                            ilCRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                                            'ilRet = MsgBox("Task could not be completed successfully, Redo Task", vbOkOnly + vbExclamation, "Spot")
                                                                            ilRet = MsgBox("Task could not be completed successfully because of " & str$(igBtrError) & ": " & sgErrLoc & ", Redo Task", vbOKOnly + vbExclamation, "Spot")
                                                                            Exit Sub
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            ilCntAllIndex = ilCntAllIndex + 1
                                            If ilCntAllIndex >= UBound(tgCntSpot) Then
                                                ilCntAllIndex = LBound(tgCntSpot)
                                            End If
                                            If Not ilSetStartLpTest Then
                                                If ilCntAllIndex = ilStartCntAllIndex Then
                                                    Exit Do
                                                End If
                                            Else
                                                ilStartCntAllIndex = ilCntAllIndex
                                            End If
                                            ilSetStartLpTest = False
                                        Loop
                                    End If
                                Else
                                    ilEvt = ilEvt + tmAvail.iNoSpotsThis
                                End If
                            End If
                            ilEvt = ilEvt + 1
                            DoEvents
                        Loop
                        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                        ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                End If
            Next ilGsf
        End If
    Next llDate
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    If (imFDateIndex <> 2) And (imFDateIndex <> 3) Then
        'lacDate.Visible = False
        Exit Sub
    End If
    slStr = edcFDropDown(imFDateIndex).Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGet30Count                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get 30sec count for week       *
'*                                                     *
'*******************************************************
Private Sub mGet30Count(ilGameNo As Integer, ilResetDays As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVef                         llGsfDate                                               *
'******************************************************************************************

    Dim llDate As Long
    Dim slDate As String
    Dim ilDay As Integer
    Dim ilRet As Integer
    Dim ilType As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim slSTime As String
    Dim llSTime As Long
    Dim slETime As String
    Dim llETime As Long
    Dim llTime As Long
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilSvLen As Integer
    Dim ilSvUnits As Integer
    Dim ilEvt As Integer
    Dim il30 As Integer
    Dim il60 As Integer
    Dim ilSpot As Integer
    Dim ilSelLen As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilGsf As Integer

    If Not imGameVehicle Then
        ilType = 0
        slSTime = edcTDropDown(0).Text
        slETime = edcTDropDown(1).Text
        If (Not gValidTime(slSTime)) Or (slSTime = "") Then
            For ilDay = 0 To 6 Step 1
                lac30(ilDay).Caption = ""
                lac60(ilDay).Caption = ""
                lacLen(ilDay).Caption = ""
            Next ilDay
            lmPSTime = -1
            Exit Sub
        End If
        If Not gValidTime(slETime) Or (slETime = "") Then
            For ilDay = 0 To 6 Step 1
                lac30(ilDay).Caption = ""
                lac60(ilDay).Caption = ""
                lacLen(ilDay).Caption = ""
            Next ilDay
            lmPETime = -1
            Exit Sub
        End If
        llSTime = CLng(gTimeToCurrency(slSTime, False))
        llETime = CLng(gTimeToCurrency(slETime, True)) - 1
        If (llSTime = lmPSTime) And (llETime = lmPETime) Then
            Exit Sub
        End If
        If imFillVehType <> 0 Then
            lmPETime = -1
            Exit Sub
        End If
        If lbcTVehicle(0).ListIndex < 0 Then
            lmPETime = -1
            Exit Sub
        End If
        slNameCode = tgUserVehicle(lbcTVehicle(0).ListIndex).sKey 'Traffic!lbcUserVehicle.List(ilVef)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        imToVefCode = Val(slCode)
        lmPSTime = llSTime
        lmPETime = llETime
        Screen.MousePointer = vbHourglass
        llSTime = CLng(gTimeToCurrency(slSTime, False))
        llETime = CLng(gTimeToCurrency(slETime, True))
        ReDim tmGsf(0 To 1) As GSF
        tmGsf(0).iGameNo = 0
        If ilResetDays Then
            For ilDay = 0 To 6 Step 1
                ckcDay(ilDay).Value = vbUnchecked
                ckcDay(ilDay).Enabled = False
            Next ilDay
        End If
    Else
        imToVefCode = imFromVefCode
        If imMixtureOfVehicles <= 0 Then
            ilDay = 0
            lmSDate = 0
            lmEDate = 0
            llSTime = 0
            llETime = 86400
        Else
            ilDay = 0
            'lmSDate = 0
            'lmEDate = 0
            llSTime = 0
            llETime = 86400
        End If
        ReDim tmGsf(0 To 1) As GSF
        tmGsf(0).iGameNo = ilGameNo
        For ilGsf = 0 To lbcToGame.ListItems.Count - 1 Step 1
            If Val(lbcToGame.ListItems(ilGsf + 1).Text) = ilGameNo Then
                lmSDate = gDateValue(lbcToGame.ListItems(ilGsf + 1).SubItems(1))
                lmEDate = lmSDate
                Exit For
            End If
        Next ilGsf
    End If
    For ilDay = 0 To 6 Step 1
        lac30(ilDay).Caption = ""
        lac60(ilDay).Caption = ""
        lacLen(ilDay).Caption = ""
    Next ilDay
    For llDate = lmSDate To lmEDate Step 1
        il30 = 0
        il60 = 0
        ilSelLen = 0
        For ilGsf = 0 To UBound(tmGsf) - 1 Step 1
            ilType = tmGsf(ilGsf).iGameNo
            If Not imGameVehicle Then
                ilDay = gWeekDayLong(llDate)
                If ilResetDays Then
                    ckcDay(ilDay).Value = vbChecked
                    ckcDay(ilDay).Enabled = True
                End If
                slDate = Format$(llDate, "m/d/yy")
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                gPackDate slDate, ilLogDate0, ilLogDate1
                tmSsfSrchKey.iType = ilType
                tmSsfSrchKey.iVefCode = imToVefCode
                tmSsfSrchKey.iDate(0) = ilLogDate0
                tmSsfSrchKey.iDate(1) = ilLogDate1
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilType) Or (tmSsf.iVefCode <> imToVefCode) Or (tmSsf.iDate(0) <> ilLogDate0) Or (tmSsf.iDate(1) <> ilLogDate1) Then
                    ckcDay(ilDay).Value = vbUnchecked
                    ckcDay(ilDay).Enabled = False
                End If
                If llDate < lgFillAllowDate Then
                    ckcDay(ilDay).Value = vbUnchecked
                    ckcDay(ilDay).Enabled = False
                End If
            Else
                'imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                'tmSsfSrchKey1.iVefCode = imToVefCode
                'tmSsfSrchKey1.iType = ilType
                'ilRet = gSSFGetEqualKey1(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                slDate = Format$(llDate, "m/d/yy")
                gPackDate slDate, ilLogDate0, ilLogDate1
                tmSsfSrchKey2.iVefCode = imToVefCode
                tmSsfSrchKey2.iDate(0) = ilLogDate0
                tmSsfSrchKey2.iDate(1) = ilLogDate1
                ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = imToVefCode)
                    If (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1) Then
                        If tmSsf.iType = ilGameNo Then
                            Exit Do
                        End If
                    End If
                    imSsfRecLen = Len(tmSsf)
                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                Loop
            End If
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = imToVefCode) And (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1)
                ilEvt = 1
                Do While ilEvt <= tmSsf.iCount
                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    'If ((tmAvail.iRecType = 2)) Or ((tmAvail.iRecType = 8) And (ckcSpotType(1).Value)) Or ((tmAvail.iRecType = 9) And (ckcSpotType(2).Value)) Then
                    If (tmAvail.iRecType = 2) Then
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        If llTime > llETime Then
                            Exit Do
                        End If
                        If llTime >= llSTime Then
                            'Test if within selected times
                            ilLen = tmAvail.iLen
                            ilUnits = tmAvail.iAvInfo And &H1F
                            For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                ilEvt = ilEvt + 1
                               LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                    ilUnits = ilUnits - 1
                                    ilLen = ilLen - (tmSpot.iPosLen And &HFFF)
                                End If
                            Next ilSpot
                            If (ilUnits > 0) And (ilLen > 0) And (imSelLen > 0) And (tmAvail.iLen = imSelLen) Then
                                ilSelLen = ilSelLen + 1
                            End If
                            ilSvUnits = ilUnits
                            ilSvLen = ilLen
                            Do While (ilUnits > 0) And (ilLen >= 30)
                                il30 = il30 + 1
                                ilUnits = ilUnits - 1
                                ilLen = ilLen - 30
                            Loop
                            ilUnits = ilSvUnits
                            ilLen = ilSvLen
                            Do While (ilUnits > 0) And (ilLen >= 60)
                                il60 = il60 + 1
                                ilUnits = ilUnits - 1
                                ilLen = ilLen - 60
                            Loop
                        Else
                            ilEvt = ilEvt + tmAvail.iNoSpotsThis
                        End If
                    End If
                    ilEvt = ilEvt + 1
                Loop
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilGsf
        If Not imGameVehicle Then
            lac30(ilDay).Caption = Trim$(str$(il30))
            lac60(ilDay).Caption = Trim$(str$(il60))
            lacLen(ilDay).Caption = Trim$(str$(ilSelLen))
        End If
        DoEvents
    Next llDate
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetSpotSumm                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Spot Summary               *
'*                                                     *
'*******************************************************
Private Sub mGetSpotSum(ilVefCode As Integer, slSDate As String, slEDate As String, llSTime As Long, llETime As Long, ilGameNo As Integer)
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilOffSet As Integer
    Dim ilIndex As Integer
    Dim ilSpotOK As Integer
    Dim llRecPos As Long
    Dim llTime As Long
    Dim llSpotDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record

    If ilGameNo <= 0 Then
        ilExtLen = Len(tmSdf) 'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
        btrExtClear hmSdf   'Clear any previous extend operation
        tmSdfSrchKey1.iVefCode = ilVefCode
        gPackDate slSDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
        tmSdfSrchKey1.iTime(0) = 0
        tmSdfSrchKey1.iTime(1) = 0
        tmSdfSrchKey1.sSchStatus = " "
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            Exit Sub
        End If
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilVefCode
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        gPackDate slSDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        gPackDate slEDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilRet = btrExtAddField(hmSdf, 0, Len(tmSdf))  'Extract Name
        ilUpper = UBound(tgCntSpot)
        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            ilExtLen = Len(tmSdf)
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilSpotOK = False
                gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                If (llTime >= llSTime) And (llTime <= llETime) Then
                    'filter out spot only if not allowed to run on selected days (cff)
                    'mLinePop removed code to set start/end date range
                    'gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                    'ilDay = gWeekDayLong(llDate)
                    'If ckcDay(ilDay).Value Then
                        ilSpotOK = True
                    'End If
                End If
                '5/18/11: Bypass Open/Close BB's
                If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
                    ilSpotOK = False
                End If
                If ilSpotOK Then
                    ilFound = False
                    For ilTest = LBound(tgCntSpot) To UBound(tgCntSpot) - 1 Step 1
                        If (tgCntSpot(ilTest).lChfCode = tmSdf.lChfCode) And (tgCntSpot(ilTest).iLineNo = tmSdf.iLineNo) Then
                            ilFound = True
                            ilIndex = ilTest
                            Exit For
                        End If
                    Next ilTest
                    If Not ilFound Then
                        tgCntSpot(ilUpper).sType = "C"
                        tgCntSpot(ilUpper).iAdfCode = tmSdf.iAdfCode
                        tgCntSpot(ilUpper).lChfCode = tmSdf.lChfCode
                        tgCntSpot(ilUpper).lFsfCode = tmSdf.lFsfCode
                        tgCntSpot(ilUpper).iVefCode = tmSdf.iVefCode
                        tgCntSpot(ilUpper).iLineNo = tmSdf.iLineNo
                        tgCntSpot(ilUpper).iNoSSpots = 0
                        tgCntSpot(ilUpper).iNoGSpots = 0
                        tgCntSpot(ilUpper).iNoMSpots = 0
                        tgCntSpot(ilUpper).iNoESpots = 0
                        tgCntSpot(ilUpper).lSdfRecPos = llRecPos
                        tgCntSpot(ilUpper).iGameNo = tmSdf.iGameNo
                        tgCntSpot(ilUpper).lSepLength = -1
                        ilIndex = ilUpper
                        ilUpper = ilUpper + 1
                        ReDim Preserve tgCntSpot(0 To ilUpper) As CNTSPOT
                    End If
                    If tmSdf.sSpotType <> "X" Then
                        If tmSdf.sSchStatus = "S" Then
                            tgCntSpot(ilIndex).iNoSSpots = tgCntSpot(ilIndex).iNoSSpots + 1
                        ElseIf (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
                            tgCntSpot(ilIndex).iNoGSpots = tgCntSpot(ilIndex).iNoGSpots + 1
                        Else
                            tgCntSpot(ilIndex).iNoMSpots = tgCntSpot(ilIndex).iNoMSpots + 1
                        End If
                    Else
                        tgCntSpot(ilIndex).iNoESpots = tgCntSpot(ilIndex).iNoESpots + 1
                    End If
                End If
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    Else
        llSDate = gDateValue(slSDate)
        llEDate = gDateValue(slEDate)
        tmSdfSrchKey6.iVefCode = ilVefCode
        tmSdfSrchKey6.iGameNo = ilGameNo
        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey6, INDEXKEY6, BTRV_LOCK_NONE, SETFORWRITE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.iGameNo = ilGameNo)
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSpotDate
            If (llSpotDate >= llSDate) And (llSpotDate <= llEDate) Then
                ilFound = False
                For ilTest = LBound(tgCntSpot) To UBound(tgCntSpot) - 1 Step 1
                    If (tgCntSpot(ilTest).lChfCode = tmSdf.lChfCode) And (tgCntSpot(ilTest).iLineNo = tmSdf.iLineNo) And (tgCntSpot(ilTest).iGameNo = tmSdf.iGameNo) Then
                        ilFound = True
                        ilIndex = ilTest
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    ilRet = btrGetPosition(hmSdf, llRecPos)
                    tgCntSpot(ilUpper).sType = "C"
                    tgCntSpot(ilUpper).iAdfCode = tmSdf.iAdfCode
                    tgCntSpot(ilUpper).lChfCode = tmSdf.lChfCode
                    tgCntSpot(ilUpper).lFsfCode = tmSdf.lFsfCode
                    tgCntSpot(ilUpper).iVefCode = tmSdf.iVefCode
                    tgCntSpot(ilUpper).iLineNo = tmSdf.iLineNo
                    tgCntSpot(ilUpper).iNoSSpots = 0
                    tgCntSpot(ilUpper).iNoGSpots = 0
                    tgCntSpot(ilUpper).iNoMSpots = 0
                    tgCntSpot(ilUpper).iNoESpots = 0
                    tgCntSpot(ilUpper).lSdfRecPos = llRecPos
                    tgCntSpot(ilUpper).iGameNo = tmSdf.iGameNo
                    tgCntSpot(ilUpper).lSepLength = -1
                    ilIndex = ilUpper
                    ilUpper = ilUpper + 1
                    ReDim Preserve tgCntSpot(0 To ilUpper) As CNTSPOT
                End If
                If tmSdf.sSpotType <> "X" Then
                    If tmSdf.sSchStatus = "S" Then
                        tgCntSpot(ilIndex).iNoSSpots = tgCntSpot(ilIndex).iNoSSpots + 1
                    ElseIf (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
                        tgCntSpot(ilIndex).iNoGSpots = tgCntSpot(ilIndex).iNoGSpots + 1
                    Else
                        tgCntSpot(ilIndex).iNoMSpots = tgCntSpot(ilIndex).iNoMSpots + 1
                    End If
                Else
                    tgCntSpot(ilIndex).iNoESpots = tgCntSpot(ilIndex).iNoESpots + 1
                End If
            End If
            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilVpf As Integer
    Dim ilLen As Integer
    Dim ilFound As Integer
    Dim ilLenMin As Integer
    Dim ilLenMax As Integer
    Dim ilVef As Integer
    Dim ilVff As Integer
    
    Screen.MousePointer = vbHourglass
    bmFirstCallToVpfFind = True
    imLBCtrls = 1
    imLBCDCtrls = 1
    imLBSdfMdExt = 1
    imFirstActivate = True
    imFromVefCode = igFillVefCode
    If igFillDW = 0 Then
        lmSDate = gDateValue(sgFillStartDate)
        lmEDate = gDateValue(sgFillStartDate)
    Else
        lmSDate = gDateValue(gObtainPrevMonday(sgFillStartDate))
        lmEDate = gDateValue(gObtainNextSunday(sgFillStartDate))
    End If
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    ilRet = gObtainVef()
    ilRet = gVffRead()
    imGameVehicle = False
    imMixtureOfVehicles = 0
    ilVef = gBinarySearchVef(imFromVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType = "G" Then
            imGameVehicle = True
            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                If tgVff(ilVff).iVefCode = imFromVefCode Then
                    If tgVff(ilVff).sMoveNonToSport = "Y" Then
                        imMixtureOfVehicles = imMixtureOfVehicles Or MOVENONTOSPORT
                    End If
                Else
                    If tgVff(ilVff).sMoveSportToSport = "Y" Then
                        imMixtureOfVehicles = imMixtureOfVehicles Or MOVESPORTTOSPORT
                    End If
                End If
            Next ilVff
            'lmSDate = gDateValue(sgFillStartDate)
            'lmEDate = gDateValue(sgFillEndDate)
            If imMixtureOfVehicles = 0 Then
                frcFrom.Visible = False
                frcTo.Visible = False
                pbcFillVehType.Visible = False
                cmcTVeh.Visible = False
                frcFromGame.Visible = True
                frcToGame.Visible = True
            Else
                frcFrom.Visible = True
                frcFromGame.Visible = False
                frcToGame.Visible = True
                frcTo.Visible = False
                pbcFillVehType.Visible = False
                cmcTVeh.Visible = False
            End If
            '5/5/11: Active Manual contract mode
            'For ilLoop = 2 To 6 Step 1
            '    ckcSpotType(ilLoop).Visible = False
            'Next ilLoop
            '5/5/11
        Else
            frcFromGame.Visible = False
            frcToGame.Visible = False
            frcFrom.Visible = True
            frcTo.Visible = True
            pbcFillVehType.Visible = True
            cmcTVeh.Visible = True
        End If
    End If
    'lacTitle.Caption = "Line Vehicle              Daypart                                  Current Dates"
    igTerminateReturn = 0   'Cancel selected
    ReDim lmDatesClearedResv(0 To 0) As Long
    hmSdf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", SpotFill
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)  'Get and save SDF record length
    ' Spot schedule File
    hmSsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", SpotFill
    On Error GoTo 0
    imSsfRecLen = Len(tmSsf)  'Get and save ADF record length
    'Header
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", SpotFill
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'Get and save CHF record length
    'Line
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", SpotFill
    On Error GoTo 0
    imClfRecLen = Len(tmClf)  'Get and save CLF record length
    'Daypart
    hmRdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rdf.Btr)", SpotFill
    On Error GoTo 0
    imRdfRecLen = Len(tmRdf)  'Get and save RPF record length
    'Advertiser
    hmAdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", SpotFill
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)  'Get and save ADF record length
    'Line flights
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", SpotFill
    On Error GoTo 0
    imCffRecLen = Len(tmCff)  'Get and save CFF record length
    hmGhf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", SpotFill
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)  'Get and save CFF record length
    hmGsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gsf.Btr)", SpotFill
    On Error GoTo 0
    ReDim tmGsf(0 To 0) As GSF
    imGsfRecLen = Len(tmGsf(0))  'Get and save CFF record length
    hmCgf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cgf.Btr)", SpotFill
    On Error GoTo 0
    imCgfRecLen = Len(tmCgf)  'Get and save CFF record length
    ' Spot MG File
    hmSmf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", SpotFill
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)  'Get and save SMF record length
    'Spot tracking
    hmStf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmStf, "", sgDBPath & "Stf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Stf.Btr)", SpotFill
    On Error GoTo 0
    imStfRecLen = Len(tmStf)  'Get and save STF record length
    'Feed Spot
    hmFsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Fsf.Btr)", SpotFill
    On Error GoTo 0
    imFsfRecLen = Len(tmFsf)  'Get and save STF record length
    'Feed Name
    hmFnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmFnf, "", sgDBPath & "Fnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Fnf.Btr)", SpotFill
    On Error GoTo 0
    'Product
    hmPrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Prf.Btr)", SpotFill
    On Error GoTo 0
    'Library Calendar
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", SpotFill
    On Error GoTo 0
    'Spot Extension
    hmSxf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sxf.Btr)", SpotFill
    On Error GoTo 0
    'Copy Rotation
    hmCrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Crf.Btr)", SpotFill
    On Error GoTo 0
    hmRlf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rlf.Btr)", SpotFill
    On Error GoTo 0
    'If tgSpf.sSchdRemnant = "Y" Then
    '    'ckcSpotType(5).Visible = False
    '    'ckcSpotType(6).Left = ckcSpotType(5).Left
    '    ckcSpotType(5).Enabled = False
    'End If
    'If tgSpf.sSchdPromo = "Y" Then
    '    ckcSpotType(3).Enabled = False
    'End If
    'If tgSpf.sSchdPSA = "Y" Then
    '    ckcSpotType(2).Enabled = False
    'End If
    imFillVehType = 1
    pbcFillVehType_MouseUp 0, 0, 0, 0
    imDetail = 0
    SpotFill.height = cmcCancel.Top + 5 * cmcCancel.height / 3
    'gCenterModalForm SpotFill
    gCenterModalForm SpotFill
    imListField(1) = 15
    imListField(2) = 29 * igAlignCharWidth
    imListField(3) = 36 * igAlignCharWidth
    imListField(4) = 50 * igAlignCharWidth
    imListField(5) = 68 * igAlignCharWidth
    imListField(6) = 76 * igAlignCharWidth
    imListField(7) = 80 * igAlignCharWidth
    imListField(8) = 88 * igAlignCharWidth
    imListField(9) = 92 * igAlignCharWidth
    imListField(10) = 120 * igAlignCharWidth

    ReDim tgCntSpot(0 To 0) As CNTSPOT
    ReDim tgDCntSpot(0 To 0) As CNTSPOT
    ReDim tgSCntSpot(0 To 0) As CNTSPOT
    ReDim tmVcf0(0 To 0) As VCF
    ReDim tmVcf6(0 To 0) As VCF
    ReDim tmVcf7(0 To 0) As VCF
    imCalType = 0   'Standard
    mInitBox
    lbcVehicle.Clear 'Force population
    mVehPop
    If imTerminate Then ' this is set by mVehPop if error occurs
        Exit Sub
    End If
    If Not imGameVehicle Then
        ReDim imLengths(0 To 0) As Integer
        ilLenMin = 32000
        ilLenMax = 0
        For ilVpf = LBound(tgVpf) To UBound(tgVpf) Step 1
            'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '    If tgMVef(ilVef).iCode = tgVpf(ilVpf).iVefKCode Then
                ilVef = gBinarySearchVef(tgVpf(ilVpf).iVefKCode)
                If ilVef <> -1 Then
                    If (tgMVef(ilVef).sState = "A") And ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S")) Then
                        For ilLen = LBound(tgVpf(ilVpf).iSLen) To UBound(tgVpf(ilVpf).iSLen) Step 1
                            If tgVpf(ilVpf).iSLen(ilLen) > 0 Then
                                If (tgVpf(ilVpf).iSLen(ilLen) <> 30) And (tgVpf(ilVpf).iSLen(ilLen) <> 60) Then

                                    ilFound = False
                                    For ilLoop = LBound(imLengths) To UBound(imLengths) - 1 Step 1
                                        If imLengths(ilLoop) = tgVpf(ilVpf).iSLen(ilLen) Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If Not ilFound Then
                                        If tgVpf(ilVpf).iSLen(ilLen) < ilLenMin Then
                                            ilLenMin = tgVpf(ilVpf).iSLen(ilLen)
                                        End If
                                        If tgVpf(ilVpf).iSLen(ilLen) > ilLenMax Then
                                            ilLenMax = tgVpf(ilVpf).iSLen(ilLen)
                                        End If
                                        imLengths(UBound(imLengths)) = tgVpf(ilVpf).iSLen(ilLen)
                                        ReDim Preserve imLengths(0 To UBound(imLengths) + 1) As Integer
                                    End If
                                End If
                            End If
                        Next ilLen
                    End If
            '        Exit For
                End If
            'Next ilVef
        Next ilVpf
        For ilLoop = ilLenMin To ilLenMax Step 1
            For ilLen = LBound(imLengths) To UBound(imLengths) - 1 Step 1
                If ilLoop = imLengths(ilLen) Then
                    cbcLen.AddItem Trim$(str$(imLengths(ilLen)))
                    Exit For
                End If
            Next ilLen
        Next ilLoop
        If cbcLen.ListCount <= 0 Then
            cbcLen.Visible = False
        Else
            imIgnoreChg = True
            cbcLen.ListIndex = 0
            imSelLen = Val(cbcLen.List(0))
        End If
        'plcScreen.Caption = "Spot Fill from " & Format$(lmSDate, "m/d/yy") & " to " & Format$(lmEDate, "m/d/yy") '& " for " & sgVehName
        For ilLoop = 0 To UBound(tgUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
            slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = imFromVefCode Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", slName)
                'plcScreen.Caption = plcScreen.Caption & " " & slName
                smSingleName = slName
                lbcTVehicle(0).ListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
        If bmFirstCallToVpfFind Then
            imVpfIndex = gVpfFind(SpotFill, imFromVefCode)
            bmFirstCallToVpfFind = False
        Else
            imVpfIndex = gVpfFindIndex(imFromVefCode)
        End If
        If (tgVpf(imVpfIndex).iLLD(0) <> 0) Or (tgVpf(imVpfIndex).iLLD(1) <> 0) Then
            gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLastLogDate
        Else
            lmLastLogDate = -1
        End If
    End If
    sgCntrForDateStamp = ""
    imBypassFocus = False
    imGetMissed = True
    imBypassAll = False
    imLbcHeight = fgListHtArial825
    imFTimeIndex = -1
    imTTimeIndex = -1
    lmPSTime = 0
    lmPETime = 0
    For ilLoop = 0 To 6 Step 1
        lac30(ilLoop).Caption = ""
        lac60(ilLoop).Caption = ""
        lacLen(ilLoop).Caption = ""
    Next ilLoop
    edcFDropDown(2).Text = Format$(lmSDate, "m/d/yy")
    edcFDropDown(3).Text = Format$(lmEDate, "m/d/yy")
    edcFDropDown(0).Text = "12M"
    edcFDropDown(1).Text = "12M"
    edcTDropDown(0).Text = "12M"
    edcTDropDown(1).Text = "12M"
    If Not imGameVehicle Then
        '6/16/09:  Removed menu item Import Contracts because out of memory
        'If tgSpf.sImptCntr = "Y" Then
        '    ckcAvailNames.Value = vbChecked
        'End If
    End If
'    If plcFillInv.Visible Then
''        If tgSpf.sDefFillInv = "Y" Then
''            rbcFillInv(0).Value = True
''        Else
''            rbcFillInv(1).Value = True
''        end if
'        rbcFillInv(2).Value = True
'
'    Else
'        ckc10To15.Left = ckcDaysTimes.Left
'        ckcDaysTimes.Top = plcAC.Top
'        ckcLock.Top = plcAC.Top
'        plcAC.Top = plcFillInv.Top
'    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         llWidth                                                 *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    Dim ilLoop As Integer
    Dim llRet As Long

    'flTextHeight = pbcContract.TextHeight("1") - 35
    frcFromGame.Move frcFrom.Left, frcFrom.Top, frcFrom.Width, frcFrom.height
    frcToGame.Move frcTo.Left, frcTo.Top, frcTo.Width, frcTo.height
    lbcFromGame.Top = 210
    lbcFromGame.height = frcFromGame.height - 2 * lbcFromGame.Top - ckcFromGame.height + 30
    lbcFromGame.Width = (41 * frcFromGame.Width) / 45
    lbcFromGame.Left = (frcFromGame.Width - lbcFromGame.Width) / 2
    lbcToGame.Top = 210
    lbcToGame.height = frcToGame.height - 2 * lbcToGame.Top - ckcToGame.height + 30
    lbcToGame.Width = (41 * frcTo.Width) / 45
    lbcToGame.Left = (frcToGame.Width - lbcToGame.Width) / 2
    ckcFromGame.Left = lbcFromGame.Left
    ckcFromGame.Top = lbcFromGame.Top + lbcFromGame.height + 105
    ckcToGame.Left = lbcToGame.Left
    ckcToGame.Top = lbcToGame.Top + lbcToGame.height + 105
    mListColumnWidths
    llRet = SendMessageByNum(lbcFromGame.hWnd, LV_SETEXTENDEDLISTVIEWSTYLE, 0, LV_FULLROWSSELECT + LV_GRIDLINES)
    llRet = SendMessageByNum(lbcToGame.hWnd, LV_SETEXTENDEDLISTVIEWSTYLE, 0, LV_FULLROWSSELECT + LV_GRIDLINES)
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLinePop                        *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate line list box with    *
'*                      contracts                      *
'*                                                     *
'*******************************************************
Private Sub mLinePop(llSTime As Long, llETime As Long)
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim slCntrNo As String
    Dim slStr As String
    Dim slDate As String
    Dim ilDay As Integer
    Dim ilVpfIndex As Integer
    Dim ilUpper As Integer
    Dim llAllowedSTime As Long
    Dim llAllowedETime As Long
    Dim ilAllowedTimeIndex As Integer
    Dim slAdvtName As String
    ReDim ilAllowedDays(0 To 6) As Integer
    'Build sort key and filter out trades, N/C lines, lines only with MG's,..
    ReDim tgTCntSpot(0 To UBound(tgCntSpot)) As CNTSPOT
    For ilLoop = 0 To UBound(tgCntSpot) - 1 Step 1
        tgTCntSpot(ilLoop) = tgCntSpot(ilLoop)
    Next ilLoop
    ReDim tgCntSpot(0 To 0) As CNTSPOT
    For ilLoop = 0 To UBound(tgTCntSpot) - 1 Step 1
        tgCntSpot(UBound(tgCntSpot)) = tgTCntSpot(ilLoop)
        ilUpper = UBound(tgCntSpot)
        If tgCntSpot(ilUpper).iNoSSpots > 0 Then
            ilRet = mReadChfClfRdfRec(tgCntSpot(ilUpper).lChfCode, tgCntSpot(ilUpper).iLineNo, tgCntSpot(ilUpper).lFsfCode)
            'Eliminate Split Network buys
            If (ilRet) And (tmClf.lRafCode <= 0) And ((tmChf.sType = "C") Or ((tgSpf.sSchdRemnant = "Y") And (tmChf.sType = "T")) Or ((tgSpf.sSchdPromo = "Y") And (tmChf.sType = "M")) Or ((tgSpf.sSchdPSA = "Y") And (tmChf.sType = "S"))) Then
                tgCntSpot(ilUpper).sType = tmChf.sType
                ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tgCntSpot(ilUpper).lSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                ilRet = gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff)
                If ilRet Then
                    If tmCff.sDyWk = "D" Then
                        For ilDay = 0 To 6 Step 1
                            If tmCff.iDay(ilDay) > 0 Then
                                ilAllowedDays(ilDay) = True
                            Else
                                ilAllowedDays(ilDay) = False
                            End If
                        Next ilDay
                    Else
                        For ilDay = 0 To 6 Step 1
                            If (tmCff.iDay(ilDay) > 0) Or (tmCff.sXDay(ilDay) = "Y") Then
                                ilAllowedDays(ilDay) = True
                            Else
                                ilAllowedDays(ilDay) = False
                            End If
                        Next ilDay
                    End If
                    'Only require days to match for same vehicles
                    If imFillVehType = 0 Then
                        If imFromVefCode = tmClf.iVefCode Then
                            ilRet = False
                            For ilDay = 0 To 6 Step 1
                                If (ckcDay(ilDay).Value = vbChecked) And ilAllowedDays(ilDay) Then
                                    ilRet = True
                                    Exit For
                                End If
                            Next ilDay
                        End If
                    End If
                End If
                If ilRet Then
                    'If ((tmCff.sPriceType = "T") And (tmCff.lActPrice > 0) And (ckcSpotType(0).Value = vbChecked)) Or ((tmCff.sPriceType = "T") And (tmCff.lActPrice = 0) And (ckcSpotType(1).Value = vbChecked)) Or ((tmCff.sPriceType = "N") And (ckcSpotType(1).Value = vbChecked)) Then
                    'Else
                    '    ilRet = False
                    'End If
                    If tmChf.sType = "C" Then
                        If ((tmCff.sPriceType = "T") And (tmCff.lActPrice > 0) And (ckcSpotType(0).Value = vbChecked)) Or ((tmCff.sPriceType = "T") And (tmCff.lActPrice = 0) And (ckcSpotType(1).Value = vbChecked)) Or ((tmCff.sPriceType = "N") And (ckcSpotType(1).Value = vbChecked)) Or ((tmCff.sPriceType = "B") And (ckcSpotType(1).Value = vbChecked)) Then
                        Else
                            ilRet = False
                        End If
                    Else
                        If (tmChf.sType = "S") And (ckcSpotType(2).Value = vbUnchecked) Then
                            ilRet = False
                        End If
                        If (tmChf.sType = "M") And (ckcSpotType(3).Value = vbUnchecked) Then
                            ilRet = False
                        End If
                        If (tmChf.sType = "T") And (ckcSpotType(5).Value = vbUnchecked) Then
                            ilRet = False
                        End If
                    End If
                End If
                If (ilRet) Then
                    If tmSdf.iAdfCode <> tmAdf.iCode Then
                        tmAdfSrchKey.iCode = tmSdf.iAdfCode
                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    End If
                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tmClf.iVefCode = tgMVef(ilVef).iCode Then
                        ilVef = gBinarySearchVef(tmClf.iVefCode)
                        If ilVef <> -1 Then
                            If bmFirstCallToVpfFind Then
                                ilVpfIndex = gVpfFind(SpotFill, tmClf.iVefCode)
                                bmFirstCallToVpfFind = False
                            Else
                                ilVpfIndex = gVpfFindIndex(tmClf.iVefCode)
                            End If
                            slCntrNo = Trim$(str$(tmChf.lCntrNo))
                            Do While Len(slCntrNo) < 8
                                slCntrNo = "0" & slCntrNo
                            Loop
                            If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                                slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                            Else
                                slAdvtName = Trim$(tmAdf.sName)
                            End If
                            tgCntSpot(ilUpper).sKey = "1" & "|" & slAdvtName & "|" & slCntrNo & "|" & tgMVef(ilVef).sName
                            tgCntSpot(ilUpper).sLen = Trim$(str$(tmClf.iLen))
                            tgCntSpot(ilUpper).sProduct = tmChf.sProduct
                            tgCntSpot(ilUpper).iLnVefCode = tmClf.iVefCode
                            gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slDate
                            slStr = slDate
                            gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slDate
                            slStr = slStr & "-" & slDate
                            tgCntSpot(ilUpper).sDate = slStr
                            For ilTest = LBound(tgCntSpot(ilUpper).lAllowedSTime) To UBound(tgCntSpot(ilUpper).lAllowedSTime) Step 1
                                tgCntSpot(ilUpper).lAllowedSTime(ilTest) = -1
                                tgCntSpot(ilUpper).lAllowedETime(ilTest) = -1
                            Next ilTest
                            ilAllowedTimeIndex = LBound(tgCntSpot(ilUpper).lAllowedSTime)
                            If (tmRdf.iLtfCode(0) <> 0) Or (tmRdf.iLtfCode(1) <> 0) Or (tmRdf.iLtfCode(2) <> 0) Then
                            Else
                                If ((tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0)) Or (tgVpf(ilVpfIndex).sGMedium = "S") Then
                                    For ilTest = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1
                                        If (tmRdf.iStartTime(0, ilTest) <> 1) Or (tmRdf.iStartTime(1, ilTest) <> 0) Then
                                            gUnpackTimeLong tmRdf.iStartTime(0, ilTest), tmRdf.iStartTime(1, ilTest), False, llAllowedSTime
                                            gUnpackTimeLong tmRdf.iEndTime(0, ilTest), tmRdf.iEndTime(1, ilTest), True, llAllowedETime
                                            mChkXMid llSTime, llETime, ilAllowedTimeIndex, llAllowedSTime, llAllowedETime
                                        End If
                                    Next ilTest
                                Else
                                    gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llAllowedSTime
                                    gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llAllowedETime
                                    mChkXMid llSTime, llETime, ilAllowedTimeIndex, llAllowedSTime, llAllowedETime
                                End If
                            End If
                            If ilAllowedTimeIndex > LBound(tgCntSpot(ilUpper).lAllowedSTime) Then
                                tgCntSpot(ilUpper).iNoTimesUsed = 0
                                tgCntSpot(ilUpper).iMnfComp0 = tmChf.iMnfComp(0)
                                tgCntSpot(ilUpper).iMnfComp1 = tmChf.iMnfComp(1)
                                For ilDay = 0 To 6 Step 1
                                    tgCntSpot(ilUpper).iAllowedDays(ilDay) = ilAllowedDays(ilDay)
                                Next ilDay
                                tgCntSpot(ilUpper).sPriceType = tmCff.sPriceType
                                tgCntSpot(ilUpper).lPrice = tmCff.lActPrice
                                tgCntSpot(ilUpper).lSepLength = -1
                                ReDim Preserve tgCntSpot(0 To UBound(tgCntSpot) + 1) As CNTSPOT
                            End If
                    '        Exit For
                        End If
                    'Next ilVef
                End If
            End If
        End If
    Next ilLoop
    'ReDim tgCntSpot(0 To UBound(tgTCntSpot)) As CNTSPOT
    'For ilLoop = 0 To UBound(tgTCntSpot) - 1 Step 1
    '    tgCntSpot(ilLoop) = tgTCntSpot(ilLoop)
    'Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLoadListBox                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Load Contracts into list box   *
'*                                                     *
'*******************************************************
Private Sub mLoadListBox()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim slDay As String
    Dim ilDay As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slRate As String
    lbcLines.Clear
    If imDetail = 0 Then
        ReDim tgCntSpot(0 To UBound(tgDCntSpot)) As CNTSPOT
        For ilLoop = 0 To UBound(tgDCntSpot) Step 1
            tgCntSpot(ilLoop) = tgDCntSpot(ilLoop)
        Next ilLoop
    Else
        ReDim tgCntSpot(0 To UBound(tgSCntSpot)) As CNTSPOT
        For ilLoop = 0 To UBound(tgSCntSpot) Step 1
            tgCntSpot(ilLoop) = tgSCntSpot(ilLoop)
        Next ilLoop
    End If
    For ilLoop = 0 To UBound(tgCntSpot) - 1 Step 1
        slNameCode = tgCntSpot(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "|", slCode)  'Advertsier
        If Len(Trim$(tgCntSpot(ilLoop).sProduct)) <= 20 Then
            slStr = Left$(Trim$(slCode), 30 - Len(Trim$(tgCntSpot(ilLoop).sProduct))) & ", " & Trim$(tgCntSpot(ilLoop).sProduct)
        Else
            slStr = Left$(Trim$(slCode), 10) & ", " & Left$(Trim$(tgCntSpot(ilLoop).sProduct), 20)
        End If
        ilRet = gParseItem(slNameCode, 3, "|", slCode)  'Contract #
        slStr = slStr & "|" & Trim$(str$(Val(slCode)))
        ilRet = gParseItem(slNameCode, 4, "|", slCode)  'Vehicle
        slDay = ""
        If (Not imGameVehicle) Or ((imMixtureOfVehicles > 0) And (tgCntSpot(ilLoop).iGameNo <= 0)) Then
            For ilDay = 0 To 6 Step 1
                If tgCntSpot(ilLoop).iAllowedDays(ilDay) Then
                    Select Case ilDay
                        Case 0
                            slDay = slDay & "m"
                        Case 1
                            slDay = slDay & "t"
                        Case 2
                            slDay = slDay & "w"
                        Case 3
                            slDay = slDay & "t"
                        Case 4
                            slDay = slDay & "f"
                        Case 5
                            slDay = slDay & "s"
                        Case 6
                            slDay = slDay & "s"
                    End Select
                    'slDay = slDay & "y"
                Else
                    slDay = slDay & "-" '"n"
                End If
            Next ilDay
        End If
        If tgCntSpot(ilLoop).sType = "S" Then
            slRate = "PSA"
        ElseIf tgCntSpot(ilLoop).sType = "M" Then
            slRate = "Promo"
        ElseIf tgCntSpot(ilLoop).sType = "Q" Then
            slRate = "PI"
        ElseIf tgCntSpot(ilLoop).sType = "T" Then
            slRate = "Rem"
        ElseIf tgCntSpot(ilLoop).sType = "R" Then
            slRate = "DR"
        Else
            Select Case tgCntSpot(ilLoop).sPriceType
                Case "T"
                    slRate = gLongToStrDec(tgCntSpot(ilLoop).lPrice, 2)
                    slRate = Left$(slRate, Len(slRate) - 3)
                    If slRate = "" Then
                        slRate = "0"
                    End If
                Case "N"
                    slRate = "N/C"
                Case "M"
                    slRate = "MG"
                Case "B"
                    slRate = "Bonus"
                Case "S"
                    slRate = "Spinoff"
                Case "R"
                    slRate = "Recapturable"
                Case "A"
                    slRate = "ADU"
                Case Else
                    slRate = ""
            End Select
        End If
        If imDetail = 0 Then
            If (Not imGameVehicle) Or ((imMixtureOfVehicles > 0) And (tgCntSpot(ilLoop).iGameNo <= 0)) Then
                slStr = slStr & "|" & Trim$(slCode) & "|" & tgCntSpot(ilLoop).sDate & "|" & slDay & "|" & tgCntSpot(ilLoop).sLen & "|" & slRate & "|" & Trim$(str$(tgCntSpot(ilLoop).iNoSSpots)) & "|" & Trim$(str$(tgCntSpot(ilLoop).iNoESpots))
            Else
                slStr = slStr & "|" & Trim$(slCode) & "|" & " " & "|" & " " & "|" & tgCntSpot(ilLoop).sLen & "|" & slRate & "|" & Trim$(str$(tgCntSpot(ilLoop).iNoSSpots)) & "|" & Trim$(str$(tgCntSpot(ilLoop).iNoESpots))
            End If
        Else
            slStr = slStr & "|" & " " & "|" & " " & "|" & " " & "|" & tgCntSpot(ilLoop).sLen & "|" & slRate & "|" & Trim$(str$(tgCntSpot(ilLoop).iNoSSpots)) & "|" & Trim$(str$(tgCntSpot(ilLoop).iNoESpots))
        End If
        'lbcLines.AddItem gAlignStringByPixel(slStr, "|", imListField(), imListFieldChar())
        lbcLines.AddItem slStr
    Next ilLoop
    pbcLbcLines_Paint
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeSummary                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Make summary from contracts    *
'*                                                     *
'*******************************************************
Private Sub mMakeSummary()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilLn As Integer
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim ilLnFound As Integer
    ReDim tgSCntSpot(0 To 0) As CNTSPOT
    ReDim tgDCntSpot(0 To UBound(tgCntSpot)) As CNTSPOT
    For ilLoop = 0 To UBound(tgCntSpot) Step 1
        tgDCntSpot(ilLoop) = tgCntSpot(ilLoop)
    Next ilLoop
    ilUpper = LBound(tgSCntSpot)
    For ilLoop = LBound(tgDCntSpot) To UBound(tgDCntSpot) - 1 Step 1
        ilFound = False
        For ilIndex = LBound(tgSCntSpot) To UBound(tgSCntSpot) - 1 Step 1
            If (tgSCntSpot(ilIndex).lChfCode = tgDCntSpot(ilLoop).lChfCode) And (Val(tgSCntSpot(ilIndex).sLen) = Val(tgDCntSpot(ilLoop).sLen)) Then
                If (tgSCntSpot(ilIndex).iLineNo <> tgDCntSpot(ilLoop).iLineNo) Then
                    ilLnFound = False
                    For ilLn = LBound(tgDCntSpot) To ilLoop - 1 Step 1
                        If (tgDCntSpot(ilLoop).lChfCode = tgDCntSpot(ilLn).lChfCode) And (tgDCntSpot(ilLoop).iLineNo = tgDCntSpot(ilLn).iLineNo) Then
                            ilLnFound = True
                            Exit For
                        End If
                    Next ilLn
                    If Not ilLnFound Then
                        tgSCntSpot(ilIndex).iNoSSpots = tgSCntSpot(ilIndex).iNoSSpots + tgDCntSpot(ilLoop).iNoSSpots
                        tgSCntSpot(ilIndex).iNoGSpots = tgSCntSpot(ilIndex).iNoGSpots + tgDCntSpot(ilLoop).iNoGSpots
                        tgSCntSpot(ilIndex).iNoMSpots = tgSCntSpot(ilIndex).iNoMSpots + tgDCntSpot(ilLoop).iNoMSpots
                        tgSCntSpot(ilIndex).iNoESpots = tgSCntSpot(ilIndex).iNoESpots + tgDCntSpot(ilLoop).iNoESpots
                    End If
                End If
                ilFound = True
                Exit For
            End If
        Next ilIndex
        If Not ilFound Then
            tgSCntSpot(ilUpper) = tgDCntSpot(ilLoop)
            ilUpper = ilUpper + 1
            ReDim Preserve tgSCntSpot(LBound(tgSCntSpot) To ilUpper) As CNTSPOT
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeUnschSpot                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Create a Sdf records            *
'*                                                     *
'*                     Similar to code within          *
'*                     CntSchd.Bas                     *
'*                                                     *
'*******************************************************
Private Function mMakeUnschSpot(llChfCode As Long, ilAdfCode As Integer, ilLineNo As Integer, llFsfCode As Long, ilGameNo As Integer, slDate As String, ilVefCode As Integer, ilLen As Integer, llSdfRecPos As Long) As Integer
'
'   mMakeUnschSpot llChfCode, ilAdfCode, ilLineNo, lDate, ilVefCode, ilLen, llSdfRecPos
'   Where:
'       llChfCode(I)- Chf Code
'       ilLineNo(I)- Line number
'       slDate(I)- Date to create spot for
'       ilExtraSpot(I)- True=Extra Bonus Spot
'       llSdfRecPos(O)- Sdf Record position
'
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilRet As Integer
    Dim llSeasonStart As Long
    Dim llSeasonEnd As Long
    Dim tlGsf As GSF

    tmSdf.lCode = 0
    tmSdf.iVefCode = ilVefCode      'Vehicle Code (combos not allowed)
    tmSdf.lChfCode = llChfCode    'Contract code
    tmSdf.iLineNo = ilLineNo    'Line number
    tmSdf.lFsfCode = llFsfCode
    tmSdf.iAdfCode = ilAdfCode 'Advertiser code number
    If ilGameNo <= 0 Then
        gPackDate slDate, tmSdf.iDate(0), tmSdf.iDate(1)
    Else
        gPackDate slDate, tmSdf.iDate(0), tmSdf.iDate(1)
        tmGsfSrchKey3.iVefCode = ilVefCode
        tmGsfSrchKey3.iGameNo = ilGameNo
        ilRet = btrGetEqual(hmGsf, tlGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        Do While (ilRet = BTRV_ERR_NONE) And (tlGsf.iVefCode = ilVefCode) And (tlGsf.iGameNo = ilGameNo)
            tmGhfSrchKey0.lCode = tlGsf.lghfcode
            ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStart
                gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEnd
                If (gDateValue(slDate) >= llSeasonStart) And (gDateValue(slDate) <= llSeasonEnd) Then
                    tmSdf.iDate(0) = tlGsf.iAirDate(0)
                    tmSdf.iDate(1) = tlGsf.iAirDate(1)
                    Exit Do
                End If
            End If
            ilRet = btrGetNext(hmGsf, tlGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        Loop
    End If
    llDate = gDateValue(slDate)
    ilDay = gWeekDayLong(llDate)
    tmSdf.iTime(0) = 0
    tmSdf.iTime(1) = 0
    tmSdf.sSchStatus = "M"    'S=Scheduled, M=Missed,
                                'G=Makegood, A=on alternate log but not MG, B=on alternate Log and MG,
                                'C=Cancelled
    tmSdf.iMnfMissed = igDefaultMnfMissed   'Missed reason
    tmSdf.sTracer = " "   'M=Mouse move, N=On demand & mouse moved, C=Created in post log,
                            'N=N/A, D=on Demand & created in post log
    tmSdf.sAffChg = " "   'T=Time change, C=Copy change, B=Time and copy changed, blank=no change
    tmSdf.sPtType = "0"
    tmSdf.lCopyCode = 0        'Copy inventory code
    tmSdf.iRotNo = 0
    tmSdf.iLen = ilLen         'Spot length
    mReadChf llChfCode, llFsfCode
    If (llFsfCode > 0) And (tmFsf.lCifCode > 0) Then
        tmSdf.sPtType = "1"
        tmSdf.lCopyCode = tmFsf.lCifCode
    End If
'    If (tmChf.sType = "C") Or (tmChf.sType = "V") Then
    If (tmChf.sType = "C") Or (tmChf.sType = "V") Or ((tgSpf.sSchdRemnant = "Y") And (tmChf.sType = "T")) Or ((tgSpf.sSchdPromo = "Y") And (tmChf.sType = "M")) Or ((tgSpf.sSchdPSA = "Y") And (tmChf.sType = "S")) Then
        If tmChf.iAdfCode <> tmAdf.iCode Then
            tmAdfSrchKey.iCode = tmChf.iAdfCode
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
        'If rbcFillInv(0).Value Then
        If rbcFillInv(0).Value Then
            tmSdf.sPriceType = "+"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
            tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
        ElseIf rbcFillInv(1).Value Then
            tmSdf.sPriceType = "-"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
            tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
        Else
            If tmAdf.sBonusOnInv <> "N" Then
                tmSdf.sPriceType = "B"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
                tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
            Else
                tmSdf.sPriceType = "N"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
                tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
            End If
        End If
    Else
        tmSdf.sPriceType = "L"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
        tmSdf.sSpotType = tmChf.sType   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo
    End If
    tmSdf.sBill = "N"
    tmSdf.lSmfCode = 0
    tmSdf.iGameNo = ilGameNo
    tmSdf.iUrfCode = tgUrf(0).iCode      'Last user who modified spot
    tmSdf.sXCrossMidnight = "N"
    tmSdf.sWasMG = "N"
    tmSdf.sFromWorkArea = "N"
    tmSdf.sUnused = ""
    ilRet = btrInsert(hmSdf, tmSdf, imSdfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        mMakeUnschSpot = ilRet
        Exit Function
    End If
    ilRet = btrGetPosition(hmSdf, llSdfRecPos)
    mMakeUnschSpot = ilRet
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mManCntrPop                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mManCntrPop(slCntrType As String, llSDate As Long, llEDate As Long, llSTime As Long, llETime As Long)
'
'   mManCntrPop ilManSch
'   Where:
'       ilManSch(I)- 1=Deferred; 2=Remnant; 3=Direct Response; 4=per Inquire; 5=PSA; 6= Promo
'
    Dim ilRet As Integer 'btrieve status
    Dim ilFound As Integer
    Dim slNameCode As String  'Name and code
    Dim slCode As String    'Code number
    Dim slDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilLoop As Integer
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim slStr As String
    Dim slCntrStatus As String
    Dim ilHOType As Integer
    Dim ilUpper As Integer
    Dim slCntrNo As String
    Dim ilDay As Integer
    Dim ilTest As Integer
    Dim ilVpfIndex As Integer
    Dim llAllowedSTime As Long
    Dim llAllowedETime As Long
    Dim ilAllowedTimeIndex As Integer
    Dim slAdvtName As String
    '5/5/11: Active manual contract for games
    Dim ilGame As Integer
    Dim slVefType As String

    ReDim ilAllowedDays(0 To 6) As Integer
    If slCntrType = "" Then
        Exit Sub
    End If
    slCntrStatus = "HO"
    ilHOType = 1
    slStartDate = Format$(llSDate, "m/d/yy")
    slEndDate = Format$(llEDate, "m/d/yy")
    sgCntrForDateStamp = ""
    ilRet = gObtainCntrForDate(SpotFill, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOType, tmChfAdvtExt())
    If (ilRet <> CP_MSG_NOPOPREQ) And (ilRet <> CP_MSG_NONE) Then
        Exit Sub
    End If
    For ilLoop = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        ilRet = gObtainChfClf(hmCHF, hmClf, tmChfAdvtExt(ilLoop).lCode, False, tmChf, tgClfSpot())
        If ilRet Then
            For ilClf = LBound(tgClfSpot) To UBound(tgClfSpot) - 1 Step 1
                tmClf = tgClfSpot(ilClf).ClfRec
                ilFound = False
                '5/5/11: Active Manual mode for games
                slVefType = ""
                If imMixtureOfVehicles > 0 Then
                    ilVef = gBinarySearchVef(tmClf.iVefCode)
                    If ilVef <> -1 Then
                        slVefType = Trim$(tgMVef(ilVef).sType)
                    End If
                End If
                'If (Not imGameVehicle) Or ((imMixtureOfVehicles > 0) And (slVefType = "G")) Then
                If (Not imGameVehicle) Or (imMixtureOfVehicles > 0) Then
                '5/5/11
                    For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
                        If lbcVehicle.Selected(ilVef) Then
                            slNameCode = tgUserVehicle(ilVef).sKey 'Traffic!lbcUserVehicle.List(ilVef)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tmClf.iVefCode Then
                                ilFound = True
                                Exit For
                            End If
                        End If
                    Next ilVef
                '5/5/11: Active Manual contract mode for games
                Else
                    If imFromVefCode = tmClf.iVefCode Then
                        ilFound = True
                    End If
                End If
                '5/5/11
                If ilFound Then
                    If tmRdf.iCode <> tmClf.iRdfCode Then
                        tmRdfSrchKey.iCode = tmClf.iRdfCode  ' Rate card program/time File Code
                        ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    End If
                    ilFound = False
                    tmCffSrchKey.lChfCode = tmChf.lCode
                    tmCffSrchKey.iClfLine = tmClf.iLine
                    tmCffSrchKey.iCntRevNo = tmClf.iCntRevNo
                    tmCffSrchKey.iPropVer = tmClf.iPropVer
                    tmCffSrchKey.iStartDate(0) = 0
                    tmCffSrchKey.iStartDate(1) = 0
                    imCffRecLen = Len(tmCff)
                    ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmCff.lChfCode = tmChf.lCode) And (tmCff.iClfLine = tmClf.iLine)
                        If (tmCff.iCntRevNo = tmClf.iCntRevNo) And (tmCff.iPropVer = tmClf.iPropVer) Then 'And (tmCff(2).sDelete <> "Y") Then
                            gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llStartDate    'Week Start date
                            gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), llEndDate    'Week Start date
                            If (llEDate >= llStartDate) And (llSDate <= llEndDate) Then
                                ilFound = True
                                Exit Do
                            End If
                        End If
                        ilRet = btrGetNext(hmCff, tmCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If ilFound Then
                        If tmCff.sDyWk = "D" Then
                            For ilDay = 0 To 6 Step 1
                                If tmCff.iDay(ilDay) > 0 Then
                                    ilAllowedDays(ilDay) = True
                                Else
                                    ilAllowedDays(ilDay) = False
                                End If
                            Next ilDay
                        Else
                            For ilDay = 0 To 6 Step 1
                                If (tmCff.iDay(ilDay) > 0) Or (tmCff.sXDay(ilDay) = "Y") Then
                                    ilAllowedDays(ilDay) = True
                                Else
                                    ilAllowedDays(ilDay) = False
                                End If
                            Next ilDay
                        End If
                        If imFillVehType = 0 Then
                            ilRet = False
                            For ilDay = 0 To 6 Step 1
                                If (ckcDay(ilDay).Value = vbChecked) And ilAllowedDays(ilDay) Then
                                    ilRet = True
                                    Exit For
                                End If
                            Next ilDay
                        Else
                            ilRet = True
                        End If
                    Else
                        ilRet = False
                    End If
                    If (ilRet) Then
                        If tmChf.iAdfCode <> tmAdf.iCode Then
                            tmAdfSrchKey.iCode = tmChf.iAdfCode
                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        End If
                        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If tmClf.iVefCode = tgMVef(ilVef).iCode Then
                            ilVef = gBinarySearchVef(tmClf.iVefCode)
                            If ilVef <> -1 Then
                                If bmFirstCallToVpfFind Then
                                    ilVpfIndex = gVpfFind(SpotFill, tmClf.iVefCode)
                                    bmFirstCallToVpfFind = False
                                Else
                                    ilVpfIndex = gVpfFindIndex(tmClf.iVefCode)
                                End If
                                slCntrNo = Trim$(str$(tmChf.lCntrNo))
                                Do While Len(slCntrNo) < 8
                                    slCntrNo = "0" & slCntrNo
                                Loop
                                ilUpper = UBound(tgCntSpot)
                                tgCntSpot(ilUpper).sType = tmChf.sType
                                tgCntSpot(ilUpper).iAdfCode = tmChf.iAdfCode
                                tgCntSpot(ilUpper).lChfCode = tmChf.lCode
                                tgCntSpot(ilUpper).lFsfCode = 0
                                tgCntSpot(ilUpper).iVefCode = tmClf.iVefCode
                                tgCntSpot(ilUpper).iLnVefCode = tmClf.iVefCode
                                tgCntSpot(ilUpper).iLineNo = tmClf.iLine
                                tgCntSpot(ilUpper).iNoSSpots = 0
                                tgCntSpot(ilUpper).iNoGSpots = 0
                                tgCntSpot(ilUpper).iNoMSpots = 0
                                tgCntSpot(ilUpper).iNoESpots = 0
                                tgCntSpot(ilUpper).lSdfRecPos = 0
                                ilRet = True
                                If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                                    slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                                Else
                                    slAdvtName = Trim$(tmAdf.sName)
                                End If
                                If tmChf.sType = "S" Then
                                    tgCntSpot(ilUpper).sKey = "2" & "|" & slAdvtName & "|" & slCntrNo & "|" & tgMVef(ilVef).sName
                                ElseIf tmChf.sType = "M" Then
                                    tgCntSpot(ilUpper).sKey = "3" & "|" & slAdvtName & "|" & slCntrNo & "|" & tgMVef(ilVef).sName
                                ElseIf tmChf.sType = "Q" Then
                                    tgCntSpot(ilUpper).sKey = "4" & "|" & slAdvtName & "|" & slCntrNo & "|" & tgMVef(ilVef).sName
                                ElseIf tmChf.sType = "T" Then
                                    tgCntSpot(ilUpper).sKey = "5" & "|" & slAdvtName & "|" & slCntrNo & "|" & tgMVef(ilVef).sName
                                ElseIf tmChf.sType = "R" Then
                                    tgCntSpot(ilUpper).sKey = "6" & "|" & slAdvtName & "|" & slCntrNo & "|" & tgMVef(ilVef).sName
                                Else
                                    ilRet = False
                                End If
                                If ilRet Then
                                    tgCntSpot(ilUpper).sLen = Trim$(str$(tmClf.iLen))
                                    tgCntSpot(ilUpper).sProduct = tmChf.sProduct
                                    gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slDate
                                    slStr = slDate
                                    gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slDate
                                    slStr = slStr & "-" & slDate
                                    tgCntSpot(ilUpper).sDate = slStr
                                    For ilTest = LBound(tgCntSpot(ilUpper).lAllowedSTime) To UBound(tgCntSpot(ilUpper).lAllowedSTime) Step 1
                                        tgCntSpot(ilUpper).lAllowedSTime(ilTest) = -1
                                        tgCntSpot(ilUpper).lAllowedETime(ilTest) = -1
                                    Next ilTest
                                    ilAllowedTimeIndex = LBound(tgCntSpot(ilUpper).lAllowedSTime)
                                    If (tmRdf.iLtfCode(0) <> 0) Or (tmRdf.iLtfCode(1) <> 0) Or (tmRdf.iLtfCode(2) <> 0) Then
                                    Else
                                        If ((tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0)) Or (tgVpf(ilVpfIndex).sGMedium = "S") Then
                                            For ilTest = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1
                                                If (tmRdf.iStartTime(0, ilTest) <> 1) Or (tmRdf.iStartTime(1, ilTest) <> 0) Then
                                                    gUnpackTimeLong tmRdf.iStartTime(0, ilTest), tmRdf.iStartTime(1, ilTest), False, llAllowedSTime
                                                    gUnpackTimeLong tmRdf.iEndTime(0, ilTest), tmRdf.iEndTime(1, ilTest), True, llAllowedETime
                                                    mChkXMid llSTime, llETime, ilAllowedTimeIndex, llAllowedSTime, llAllowedETime
                                                End If
                                            Next ilTest
                                        Else
                                            gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llAllowedSTime
                                            gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llAllowedETime
                                            mChkXMid llSTime, llETime, ilAllowedTimeIndex, llAllowedSTime, llAllowedETime
                                        End If
                                    End If
                                    'If (llETime >= tgCntSpot(ilUpper).lAllowedSTime) And (llSTime <= tgCntSpot(ilUpper).lAllowedETime) Then
                                    If ilAllowedTimeIndex > LBound(tgCntSpot(ilUpper).lAllowedSTime) Then
                                        tgCntSpot(ilUpper).iNoTimesUsed = 0
                                        tgCntSpot(ilUpper).iMnfComp0 = tmChf.iMnfComp(0)
                                        tgCntSpot(ilUpper).iMnfComp1 = tmChf.iMnfComp(1)
                                        For ilDay = 0 To 6 Step 1
                                            tgCntSpot(ilUpper).iAllowedDays(ilDay) = ilAllowedDays(ilDay)
                                        Next ilDay
                                        tgCntSpot(ilUpper).lPrice = tmCff.lActPrice
                                        '5/5/11: Active manual contract for games
                                        If tgMVef(ilVef).sType <> "G" Then
                                            tgCntSpot(ilUpper).iGameNo = 0
                                            tgCntSpot(ilUpper).lSepLength = -1
                                            ReDim Preserve tgCntSpot(0 To UBound(tgCntSpot) + 1) As CNTSPOT
                                        Else
                                            For ilGame = 0 To lbcFromGame.ListItems.Count - 1 Step 1
                                                If lbcFromGame.ListItems(ilGame + 1).Selected Then
                                                    ilUpper = UBound(tgCntSpot)
                                                    tgCntSpot(ilUpper).iGameNo = Val(lbcFromGame.ListItems(ilGame + 1).Text)
                                                    tgCntSpot(ilUpper).lSepLength = -1
                                                    ReDim Preserve tgCntSpot(0 To UBound(tgCntSpot) + 1) As CNTSPOT
                                                End If
                                            Next ilGame
                                        End If
                                        '5/5/11
                                    End If
                                End If
                        '        Exit For
                            End If
                        'Next ilVef
                    End If
                End If
            Next ilClf
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMixCntSpot                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Randomly mix spots             *
'*                                                     *
'*******************************************************
Private Sub mMixCntSpot()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpper As Integer
    Dim ilNoToMove As Integer
    Dim ilLn As Integer
    ReDim tgTCntSpot(0 To 0) As CNTSPOT
    If imDetail = 0 Then
        For ilLoop = 0 To UBound(tgCntSpot) - 1 Step 1
            If lbcLines.Selected(ilLoop) Then
                tgTCntSpot(UBound(tgTCntSpot)) = tgCntSpot(ilLoop)
                tgTCntSpot(UBound(tgTCntSpot)).iUpdateIndex = ilLoop
                ReDim Preserve tgTCntSpot(0 To UBound(tgTCntSpot) + 1) As CNTSPOT
            End If
        Next ilLoop
    Else
        For ilLoop = 0 To UBound(tgCntSpot) - 1 Step 1
            If lbcLines.Selected(ilLoop) Then
                For ilLn = LBound(tgDCntSpot) To UBound(tgDCntSpot) - 1 Step 1
                    If (tgCntSpot(ilLoop).lChfCode = tgDCntSpot(ilLn).lChfCode) And (Val(tgCntSpot(ilLoop).sLen) = Val(tgDCntSpot(ilLn).sLen)) And (tgCntSpot(ilLoop).lFsfCode = tgDCntSpot(ilLn).lFsfCode) Then
                        tgTCntSpot(UBound(tgTCntSpot)) = tgDCntSpot(ilLn)
                        tgTCntSpot(UBound(tgTCntSpot)).iUpdateIndex = ilLn
                        ReDim Preserve tgTCntSpot(0 To UBound(tgTCntSpot) + 1) As CNTSPOT
                    End If
                Next ilLn
            End If
        Next ilLoop
    End If
    ReDim tgCntSpot(0 To UBound(tgTCntSpot)) As CNTSPOT
    Randomize
    ilNoToMove = UBound(tgTCntSpot)
    ilUpper = LBound(tgCntSpot)
    Do While ilNoToMove >= 1
        ilIndex = Int((ilNoToMove) * Rnd + 1) - 1
        tgCntSpot(ilUpper) = tgTCntSpot(ilIndex)
        For ilLoop = ilIndex To UBound(tgTCntSpot) - 1 Step 1
            tgTCntSpot(ilLoop) = tgTCntSpot(ilLoop + 1)
        Next ilLoop
        ilUpper = ilUpper + 1
        ilNoToMove = ilNoToMove - 1
    Loop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfClfRdfCffRec            *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read records                   *
'*                                                     *
'*******************************************************
Private Sub mReadChf(llChfCode As Long, llFsfCode As Long)
    Dim ilRet As Integer
    If llChfCode > 0 Then
        If llChfCode <> tmChf.lCode Then
            tmChfSrchKey.lCode = llChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmChf.sType = "C"   'Force spot to be fill
                Exit Sub
            End If
        End If
    Else
        tmFSFSrchKey.lCode = llFsfCode
        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        gMoveFeedToCntr tmFsf, tmRdf, tmChf, tmClf, tmFCff(), hmFnf, hmPrf
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfClfRdfCffRec            *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read records                   *
'*                                                     *
'*******************************************************
Private Function mReadChfClfRdfCffRec(llChfCode As Long, ilLineNo As Integer, llFsfCode As Long, ilOrderedGameNo As Integer, slSpotDate As String, slLnStartDate As String, slLnEndDate As String, slNoSpots As String) As Integer
'
'   iRet = mReadChfClfRdfCffRec(llChfCode, ilLineNo, slMissedDate, SlStartDate, slEndDate, slNoSpots)
'   Where:
'       llChfCode (I) - Contract code
'       ilLineNo(I)- Line number
'       slMissedDate(I)- Missed date or date to find bracketing week
'       slLnStartdate(O)- line start date
'       slLnEndDate(O)- line end date
'       slNoSpots(O)- if "" then invalid week
'       tmICff(1)(O)- contains valid flight week (if sDelete = "Y", then week is invalid)
'       iRet (O)- True if all records read,
'                 False if error in read
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim llSpotDate As Long
    Dim ilNoSpots As Integer
    Dim ilVef As Integer

    slLnStartDate = ""
    slLnEndDate = ""
    slNoSpots = ""
    tmICff(1).sDelete = "Y"  'Set as flag that illegal week
    If mReadChfClfRdfRec(llChfCode, ilLineNo, llFsfCode) Then
        llStartDate = 0
        llEndDate = 0
        llSpotDate = gDateValue(slSpotDate)
        If llChfCode > 0 Then
            ilVef = gBinarySearchVef(tmClf.iVefCode)
            If ilVef <> -1 Then
                If tgMVef(ilVef).sType <> "G" Then
                    tmCffSrchKey.lChfCode = llChfCode
                    tmCffSrchKey.iClfLine = ilLineNo
                    tmCffSrchKey.iCntRevNo = tmClf.iCntRevNo
                    tmCffSrchKey.iPropVer = tmClf.iPropVer
                    tmCffSrchKey.iStartDate(0) = 0
                    tmCffSrchKey.iStartDate(1) = 0
                    ilRet = btrGetGreaterOrEqual(hmCff, tmICff(2), imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Else
                    tmCgfSrchKey1.lClfCode = tmClf.lCode
                    ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If tmClf.lCode = tmCgf.lClfCode Then
                        gCgfToCff tmClf, tmCgf, tmGCff()
                        tmICff(2) = tmGCff(0)   'tmGCff(1)
                        If tmCgf.iGameNo <> ilOrderedGameNo Then
                            tmICff(2).iCntRevNo = -1 'Force to read next
                        End If
                    Else
                        tmICff(2).lChfCode = -1
                    End If
                End If
            Else
                mReadChfClfRdfCffRec = False
                Exit Function
            End If
        Else
            tmICff(2) = tmFCff(1)
            tmICff(2).lChfCode = llChfCode
            tmICff(2).iClfLine = ilLineNo
            ilRet = BTRV_ERR_NONE
        End If
        Do While (ilRet = BTRV_ERR_NONE) And (tmICff(2).lChfCode = llChfCode) And (tmICff(2).iClfLine = ilLineNo)
            If (tmICff(2).iCntRevNo = tmClf.iCntRevNo) And (tmICff(2).iPropVer = tmClf.iPropVer) Then 'And (tmICff(2).sDelete <> "Y") Then
                tmICff(2).sDelete = "N"  'Set flight as if not deleted (delete is set if line replaced)
                                        'Only if line is altered (not scheduled will this happen)
                gUnpackDate tmICff(2).iStartDate(0), tmICff(2).iStartDate(1), slStartDate    'Week Start date
                gUnpackDate tmICff(2).iEndDate(0), tmICff(2).iEndDate(1), slEndDate    'Week Start date
                If llStartDate = 0 Then
                    llStartDate = gDateValue(slStartDate)
                    llEndDate = gDateValue(slEndDate)
                Else
                    If gDateValue(slStartDate) < llStartDate Then
                        llStartDate = gDateValue(slStartDate)
                    End If
                    If gDateValue(slEndDate) > llEndDate Then
                        llEndDate = gDateValue(slEndDate)
                    End If
                End If
                If (llSpotDate >= gDateValue(slStartDate)) And (llSpotDate <= gDateValue(slEndDate)) Then
                    tmICff(1) = tmICff(2)
                    ilNoSpots = 0
                    'If (tmCff(1).iSpotsWk <> 0) Or (tmCff(1).iXSpotsWk <> 0) Then 'Weekly
                    If (tmICff(1).sDyWk <> "D") Then  'Weekly
                        ilNoSpots = tmICff(1).iSpotsWk + tmICff(1).iXSpotsWk
                    Else    'Daily
                        For ilLoop = 0 To 6 Step 1
                            ilNoSpots = ilNoSpots + tmICff(1).iDay(ilLoop)
                        Next ilLoop
                    End If
                    slNoSpots = Trim$(str$(ilNoSpots))
                    'Don't exit as end date of all flights must be determined
                End If
            End If
            If llChfCode > 0 Then
                If tgMVef(ilVef).sType <> "G" Then
                    ilRet = btrGetNext(hmCff, tmICff(2), imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Else
                    ilRet = btrGetNext(hmCgf, tmCgf, imCgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If tmClf.lCode <> tmCgf.lClfCode Then
                        Exit Do
                    End If
                    gCgfToCff tmClf, tmCgf, tmGCff()
                    tmICff(2) = tmGCff(0)   'tmGCff(1)
                    If tmCgf.iGameNo <> ilOrderedGameNo Then
                        tmICff(2).iCntRevNo = -1 'Force to read next
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
        If llStartDate > 0 Then
            slLnStartDate = Format$(llStartDate, "m/d/yy")
            slLnEndDate = Format$(llEndDate, "m/d/yy")
        End If
        mReadChfClfRdfCffRec = True
    Else
        mReadChfClfRdfCffRec = False
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfClfRdfRec               *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read records                   *
'*                                                     *
'*******************************************************
Private Function mReadChfClfRdfRec(llChfCode As Long, ilLineNo As Integer, llFsfCode As Long) As Integer
'
'   iRet = mReadChfClfRdfRec(llChfCode, ilLineNo)
'   Where:
'       llChfCode (I) - Contract code
'       ilLineNo(I)- Line number
'       iRet (O)- True if all records read,
'                 False if error in read
'
    Dim ilRet As Integer
    If llChfCode > 0 Then
        'If llChfCode <> tmChf.lCode Then
            tmChfSrchKey.lCode = llChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                mReadChfClfRdfRec = False
                Exit Function
            End If
        'End If
        'If (tmClf.lChfCode <> llChfCode) Or (tmClf.iLine <> ilLineNo) Then
            tmClfSrchKey.lChfCode = llChfCode
            tmClfSrchKey.iLine = ilLineNo
            tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
            tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iLine = ilLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))    'And (tmClf.sSchStatus = "A")
                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        'Else
        '    ilRet = BTRV_ERR_NONE
        'End If
        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iLine = ilLineNo) Then
            If tmRdf.iCode <> tmClf.iRdfCode Then
                tmRdfSrchKey.iCode = tmClf.iRdfCode  ' Rate card program/time File Code
                ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    mReadChfClfRdfRec = False
                    Exit Function
                End If
            End If
            mReadChfClfRdfRec = True
        Else
            mReadChfClfRdfRec = False
        End If
    Else
        tmFSFSrchKey.lCode = llFsfCode
        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        gMoveFeedToCntr tmFsf, tmRdf, tmChf, tmClf, tmFCff(), hmFnf, hmPrf
        mReadChfClfRdfRec = True
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mRemoveResvSpots                *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove any Reservation Spots   *
'*                                                     *
'*******************************************************
Private Sub mRemoveResvSpots(llStartDate As Long, llEndDate As Long)
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilEvt As Integer
    Dim ilSpot As Integer
    Dim ilRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim slDate As String
    Dim ilType As Integer
    Dim llSsfDate As Long
    Dim llSsfRecPos As Long
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llTestDate As Long
    Dim tlSdf As SDF
    Dim llLockRecCode As Long
    Dim slUserName As String
    Dim ilVef As Integer

    If imGetMissed Then
        slStartDate = Format$(lmSDate, "m/d/yy")
        slEndDate = Format$(lmEDate, "m/d/yy")
        ilVef = gBinarySearchVef(imToVefCode)
        'If ilRet <> -1 Then
        If ilVef <> -1 Then
            If tgMVef(ilVef).sType <> "G" Then
                ilRet = gObtainMissedSpot("M", imToVefCode, -1, 0, slStartDate, slEndDate, 1, tmSdfMdExt(), smSdfMdExtTag)
            Else
                ilRet = gObtainMissedSpot("M", imToVefCode, -1, -1, slStartDate, slEndDate, 1, tmSdfMdExt(), smSdfMdExtTag)
            End If
        End If
        imGetMissed = False
    End If
    ilType = 0
    For llDate = llStartDate To llEndDate Step 1
        ilFound = False
        For ilLoop = 0 To UBound(lmDatesClearedResv) - 1 Step 1
            If llDate = lmDatesClearedResv(ilLoop) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            'For ilLoop = LBound(tmSdfMdExt) To UBound(tmSdfMdExt) - 1 Step 1
            For ilLoop = imLBSdfMdExt To UBound(tmSdfMdExt) - 1 Step 1
                If tmSdfMdExt(ilLoop).sCntrType = "V" Then
                    tmSdfMdExt(ilLoop).sCntrType = ""
                    gUnpackDateLong tmSdfMdExt(ilLoop).iDate(0), tmSdfMdExt(ilLoop).iDate(0), llTestDate
                    If llTestDate = llDate Then
                        Do
                            ilRet = btrGetDirect(hmSdf, tlSdf, imSdfRecLen, tmSdfMdExt(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                            If ilRet <> BTRV_ERR_NONE Then
                                Exit Do
                            End If
                            If tlSdf.sSchStatus <> "M" Then
                                Exit Do
                            End If
                            'tmSRec = tlSdf
                            'ilRet = gGetByKeyForUpdate("SDF", hmSdf, tmSRec)
                            'tlSdf = tmSRec
                            If tlSdf.sSchStatus <> "M" Then
                                ilRet = BTRV_ERR_CONFLICT
                                Exit Do
                            End If
                            tlSdf.sSchStatus = "H"
                            ilRet = btrUpdate(hmSdf, tlSdf, imSdfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                    End If
                End If
            Next ilLoop
            ilDay = gWeekDayLong(llDate)
            slDate = Format$(llDate, "m/d/yy")
            llLockRecCode = gCreateLockRec(hmRlf, "S", "S", 65536 * imToVefCode + llDate, False, slUserName)
            If llLockRecCode > 0 Then
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                gPackDate slDate, ilLogDate0, ilLogDate1
                tmSsfSrchKey.iType = ilType
                tmSsfSrchKey.iVefCode = imToVefCode
                tmSsfSrchKey.iDate(0) = ilLogDate0
                tmSsfSrchKey.iDate(1) = ilLogDate1
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = imToVefCode) And (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1)
                    ilRet = gSSFGetPosition(hmSsf, llSsfRecPos)
                    ilEvt = 1
                    Do While ilEvt <= tmSsf.iCount
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Avail
                            For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                ilEvt = ilEvt + 1
                               LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                If (tmSpot.iRank And RANKMASK) = RESERVATIONRANK Then
                                    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                    ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        llSsfDate = 0   'Force read
                                        If gChgSchSpot("H", hmSdf, tlSdf, hmSmf, tlSdf.iGameNo, tmSmf, hmSsf, tmSsf, llSsfDate, llSsfRecPos, hmSxf, hmGsf, hmGhf) Then
                                            imSsfRecLen = Len(tmSsf)
                                            ilRet = gSSFGetDirect(hmSsf, tmSsf, imSsfRecLen, llSsfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                            ilEvt = 0
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next ilSpot
                        End If
                        ilEvt = ilEvt + 1
                    Loop
                    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                lmDatesClearedResv(UBound(lmDatesClearedResv)) = llDate
                ReDim Preserve lmDatesClearedResv(0 To UBound(lmDatesClearedResv) + 1) As Long
            End If
        End If
    Next llDate
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'
    Screen.MousePointer = vbDefault
    Unload SpotFill
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llFilter As Long
    Dim ilVef As Integer
    Dim llNowDate As Long
    Dim slNowDate As String
    Dim mFromItem As ListItem
    Dim mToItem As ListItem
    Dim ilGsf As Integer

    ilVef = gBinarySearchVef(imFromVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType <> "G" Then
            ''ilRet = gPopUserVehicleBox(SpotFill, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcVehicle, Traffic!lbcUserVehicle)
            'ilRet = gPopUserVehicleBox(SpotFill, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
            If tgSpf.sMktBase = "Y" Then
                'llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHEXCLUDESPORT + ACTIVEVEH + VEHBYMKT ' Airing and all conventional vehicles (except with Log) and Log
                llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHMOVESPORTTONONSPORT + ACTIVEVEH + VEHEXCLUDEIMPORTINVOICESPOTS + VEHEXCLUDEPODNOPRGM   ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 EXLUDE CPM VEHICLES
                ilRet = gPopUserVehicleByMkt(SpotFill, llFilter, igSpotMktCode(), lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
            Else
                'llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHEXCLUDESPORT + ACTIVEVEH ' Airing and all conventional vehicles (except with Log) and Log
                llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHMOVESPORTTONONSPORT + ACTIVEVEH + VEHEXCLUDEIMPORTINVOICESPOTS + VEHEXCLUDEPODNOPRGM  ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 EXLUDE CPM VEHICLES
                ilRet = gPopUserVehicleBox(SpotFill, llFilter, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
            End If
            If ilRet <> CP_MSG_NOPOPREQ Then
                On Error GoTo mVehPopErr
                gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", SpotFill
                On Error GoTo 0
            End If
            For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                lbcTVehicle(0).AddItem lbcVehicle.List(ilLoop)
                lbcTVehicle(1).AddItem lbcVehicle.List(ilLoop)
            Next ilLoop
        Else
            'If tgSpf.sMktBase = "Y" Then
            '    llFilter = VEHSPORT + ACTIVEVEH + VEHBYMKT   ' Airing and all conventional vehicles (except with Log) and Log
            '    ilRet = gPopUserVehicleByMkt(SpotFill, llFilter, igSpotMktCode(), lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
            'Else
            '    llFilter = VEHSPORT + ACTIVEVEH ' Airing and all conventional vehicles (except with Log) and Log
            '    ilRet = gPopUserVehicleBox(SpotFill, llFilter, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
            'End If
            slNowDate = Format$(gNow(), "m/d/yy")
            llNowDate = gDateValue(slNowDate)
            mTeamPop
            lbcFromGame.ListItems.Clear
            lbcToGame.ListItems.Clear
            ilRet = gGetGameDates(hmLcf, hmGhf, hmGsf, tgMVef(ilVef).iCode, tmTeam(), tmGsfInfo())
            For ilGsf = LBound(tmGsfInfo) To UBound(tmGsfInfo) - 1 Step 1
                If imMixtureOfVehicles = 0 Then
                    Set mFromItem = lbcFromGame.ListItems.Add()
                    mFromItem.Text = Trim$(str$(tmGsfInfo(ilGsf).iGameNo))
                    mFromItem.SubItems(1) = Format$(tmGsfInfo(ilGsf).lGameDate, "m/d/yy")
                    mFromItem.SubItems(2) = Trim$(Left$(tmGsfInfo(ilGsf).sVisitName, 4)) & " @" & Trim$(Left$(tmGsfInfo(ilGsf).sHomeName, 4))
                End If
                If (tmGsfInfo(ilGsf).lGameDate > llNowDate) And (tmGsfInfo(ilGsf).sGameStatus <> "C") Then
                    Set mToItem = lbcToGame.ListItems.Add()
                    mToItem.Text = Trim$(str$(tmGsfInfo(ilGsf).iGameNo))
                    mToItem.SubItems(1) = Format$(tmGsfInfo(ilGsf).lGameDate, "m/d/yy")
                    mGet30Count tmGsfInfo(ilGsf).iGameNo, True
                    mToItem.SubItems(2) = lac30(0).Caption   'Number of 30s
                    mToItem.SubItems(3) = lac60(0).Caption   'Number of 60s
                End If
            Next ilGsf
            If imMixtureOfVehicles > 0 Then
                If ((imMixtureOfVehicles And MOVENONTOSPORT) = MOVENONTOSPORT) And ((imMixtureOfVehicles And MOVESPORTTOSPORT) = MOVESPORTTOSPORT) Then
                    If tgSpf.sMktBase = "Y" Then
                        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHMOVESPORTTOSPORT + ACTIVEVEH + VEHBYMKT + VEHEXCLUDEIMPORTINVOICESPOTS ' Airing and all conventional vehicles (except with Log) and Log
                        ilRet = gPopUserVehicleByMkt(SpotFill, llFilter, igSpotMktCode(), lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
                    Else
                        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHMOVESPORTTOSPORT + ACTIVEVEH + VEHEXCLUDEIMPORTINVOICESPOTS ' Airing and all conventional vehicles (except with Log) and Log
                        ilRet = gPopUserVehicleBox(SpotFill, llFilter, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
                    End If
                ElseIf ((imMixtureOfVehicles And MOVENONTOSPORT) = MOVENONTOSPORT) Then
                    If tgSpf.sMktBase = "Y" Then
                        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHEXCLUDESPORT + ACTIVEVEH + VEHBYMKT + VEHEXCLUDEIMPORTINVOICESPOTS ' Airing and all conventional vehicles (except with Log) and Log
                        ilRet = gPopUserVehicleByMkt(SpotFill, llFilter, igSpotMktCode(), lbcVehicle, tgUserVehicle(), sgUserVehicleTag, imFromVefCode)
                    Else
                        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHEXCLUDESPORT + ACTIVEVEH + VEHEXCLUDEIMPORTINVOICESPOTS ' Airing and all conventional vehicles (except with Log) and Log
                        ilRet = gPopUserVehicleBox(SpotFill, llFilter, lbcVehicle, tgUserVehicle(), sgUserVehicleTag, imFromVefCode)
                    End If
                ElseIf ((imMixtureOfVehicles And MOVESPORTTOSPORT) = MOVESPORTTOSPORT) Then
                    If tgSpf.sMktBase = "Y" Then
                        llFilter = VEHMOVESPORTTOSPORT + ACTIVEVEH + VEHBYMKT + VEHEXCLUDEIMPORTINVOICESPOTS ' Airing and all conventional vehicles (except with Log) and Log
                        ilRet = gPopUserVehicleByMkt(SpotFill, llFilter, igSpotMktCode(), lbcVehicle, tgUserVehicle(), sgUserVehicleTag, imFromVefCode)
                    Else
                        llFilter = VEHMOVESPORTTOSPORT + ACTIVEVEH + VEHEXCLUDEIMPORTINVOICESPOTS ' Airing and all conventional vehicles (except with Log) and Log
                        ilRet = gPopUserVehicleBox(SpotFill, llFilter, lbcVehicle, tgUserVehicle(), sgUserVehicleTag, imFromVefCode)
                    End If
                End If
                If ilRet <> CP_MSG_NOPOPREQ Then
                    On Error GoTo mVehPopErr
                    gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", SpotFill
                    On Error GoTo 0
                End If
                'For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                '    lbcTVehicle(0).AddItem lbcVehicle.List(ilLoop)
                '    lbcTVehicle(1).AddItem lbcVehicle.List(ilLoop)
                'Next ilLoop
            End If
        End If
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub lbcVehicle_Scroll()
    If tmcLine.Enabled Then
        tmcLine.Enabled = False
        tmcLine.Enabled = True
    End If
End Sub

Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcFDropDown(imFDateIndex).Text = Format$(llDate, "m/d/yy")
                edcFDropDown(imFDateIndex).SelStart = 0
                edcFDropDown(imFDateIndex).SelLength = Len(edcFDropDown(imFDateIndex).Text)
                imBypassFocus = True
                edcFDropDown(imFDateIndex).SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcFDropDown(imFDateIndex).SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcDetail_GotFocus()
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus pbcDetail
End Sub
Private Sub pbcDetail_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("L")) Or (KeyAscii = Asc("l")) Then
        imDetail = 0
        pbcDetail.Cls
        pbcDetail_Paint
        mLoadListBox
    ElseIf (KeyAscii = Asc("C")) Or (KeyAscii = Asc("c")) Then
        imDetail = 1
        pbcDetail.Cls
        pbcDetail_Paint
        mLoadListBox
    End If
    If KeyAscii = Asc(" ") Then
        If imDetail = 1 Then
            imDetail = 0
        ElseIf imDetail = 0 Then
            imDetail = 1
        End If
        pbcDetail.Cls
        pbcDetail_Paint
        mLoadListBox
    End If
End Sub
Private Sub pbcDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDetail = 1 Then
        imDetail = 0
    ElseIf imDetail = 0 Then
        imDetail = 1
    End If
    pbcDetail.Cls
    pbcDetail_Paint
    mLoadListBox
End Sub
Private Sub pbcDetail_Paint()
    pbcDetail.CurrentX = fgBoxInsetX \ 20
    pbcDetail.CurrentY = 0
    If imDetail = 0 Then
        pbcDetail.Print "Line"
    Else
        pbcDetail.Print "Contract"
    End If
End Sub
Private Sub pbcFillVehType_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("S")) Or (KeyAscii = Asc("s")) Then
        imFillVehType = 1
        pbcFillVehType_MouseUp 0, 0, 0, 0
    ElseIf (KeyAscii = Asc("M")) Or (KeyAscii = Asc("m")) Then
        imFillVehType = 0
        pbcFillVehType_MouseUp 0, 0, 0, 0
    End If
    If KeyAscii = Asc(" ") Then
        pbcFillVehType_MouseUp 0, 0, 0, 0
    End If
End Sub
Private Sub pbcFillVehType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imFillVehType = 1 Then
        imFillVehType = 0
        lbcTVehicle(1).Visible = False
        ckcTAll.Visible = False
        plcTDays.Visible = False
        lbcTVehicle(0).Visible = False  'True
        plcSVeh.Visible = True
        lacTSTime.Move 60, 1485
        edcTDropDown(0).Move 975, 1485
        cmcTDropDown(0).Move 1860, 1485
        lacTETime.Move 2145, 1485
        edcTDropDown(1).Move 3090, 1485
        cmcTDropDown(1).Move 3975, 1485
        lmPETime = -1
        mGet30Count 0, True
    ElseIf imFillVehType = 0 Then
        imFillVehType = 1
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = False
        lbcTVehicle(1).Visible = True
        ckcTAll.Move lbcTVehicle(1).Left, 1455
        ckcTAll.Visible = True
        plcTDays.Move 825, 1530, 3375, 240
        plcTDays.Visible = True
        lacTSTime.Move 3090, 255
        edcTDropDown(0).Move 3090, 480
        cmcTDropDown(0).Move 3975, 480
        lacTETime.Move 3090, 735
        edcTDropDown(1).Move 3090, 960
        cmcTDropDown(1).Move 3975, 960
    End If
    pbcFillVehType.Cls
    pbcFillVehType_Paint
End Sub
Private Sub pbcFillVehType_Paint()
    pbcFillVehType.CurrentX = fgBoxInsetX \ 20
    pbcFillVehType.CurrentY = 0
    If imFillVehType = 0 Then
        pbcFillVehType.Print smSingleName   '"Single Vehicle"
    Else
        pbcFillVehType.Print "Multi-Vehicle"
    End If
End Sub

Private Sub pbcLbcLines_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLinesEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 8) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    
    ilLinesEnd = lbcLines.TopIndex + lbcLines.height \ fgListHtArial825
    If ilLinesEnd > lbcLines.ListCount Then
        ilLinesEnd = lbcLines.ListCount
    End If
    If lbcLines.ListCount <= lbcLines.height \ fgListHtArial825 Then
        llWidth = lbcLines.Width - 30
    Else
        llWidth = lbcLines.Width - igScrollBarWidth - 30
    End If
    pbcLbcLines.Width = llWidth
    pbcLbcLines.Cls
    llFgColor = pbcLbcLines.ForeColor
    For ilLoop = lbcLines.TopIndex To ilLinesEnd - 1 Step 1
        pbcLbcLines.ForeColor = llFgColor
        If lbcLines.MultiSelect = 0 Then
            If lbcLines.ListIndex = ilLoop Then
                gPaintArea pbcLbcLines, CSng(0), CSng((ilLoop - lbcLines.TopIndex) * fgListHtArial825), CSng(pbcLbcLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLines.ForeColor = vbWhite
            End If
        Else
            If lbcLines.Selected(ilLoop) Then
                gPaintArea pbcLbcLines, CSng(0), CSng((ilLoop - lbcLines.TopIndex) * fgListHtArial825), CSng(pbcLbcLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLines.ForeColor = vbWhite
            End If
        End If
        slStr = lbcLines.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcLines.CurrentX = imListField(ilField)
            pbcLbcLines.CurrentY = (ilLoop - lbcLines.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcLines, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcLines.Print slStr
        Next ilField
        pbcLbcLines.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    If imFTimeIndex <> -1 Then
                        imBypassFocus = True    'Don't change select text
                        edcFDropDown(imFTimeIndex).SetFocus
                        'SendKeys slKey
                        gSendKeys edcFDropDown(imFTimeIndex), slKey
                    ElseIf imTTimeIndex <> -1 Then
                        imBypassFocus = True    'Don't change select text
                        edcTDropDown(imTTimeIndex).SetFocus
                        'SendKeys slKey
                        gSendKeys edcTDropDown(imTTimeIndex), slKey
                    End If
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub plcInv_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcAC_GotFocus(Index As Integer)
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub
Private Sub rbcFillInv_GotFocus(Index As Integer)
    If imFillVehType = 0 Then
        lbcTVehicle(0).Visible = False
        plcSVeh.Visible = True
    End If
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub
Private Sub tmcDrag_Timer()
    Dim ilListIndex As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            If imDragButton <> 1 Then
                Exit Sub
            End If
            If imFillVehType <> 0 Then
                Exit Sub
            End If
            Select Case imDragSrce
                Case DRAGLINE
                    ilListIndex = (fmDragY \ imLbcHeight) + lbcLines.TopIndex
                    If (ilListIndex >= 0) And (ilListIndex <= lbcLines.ListCount - 1) Then
                        lbcLines.DragIcon = IconTraf!imcIconDrag.DragIcon
                        imDragIndexSrce = ilListIndex
                        lbcLines.Drag vbBeginDrag
                    'Else
                    '    lbcAskDisposition.ListIndex = -1
                    End If
            End Select
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub tmcLine_Timer()
    Dim ilLoop As Integer
    Dim slSTime As String
    Dim slETime As String
    Dim llSTime As Long
    Dim llETime As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slCntrType As String
    Dim slCntrSDate As String
    Dim slCntrEDate As String
    Dim llCntrSDate As Long
    Dim llCntrEDate As Long
    Dim ilGameNo As Integer

    If (Not imGameVehicle) Or (imMixtureOfVehicles > 0) Then
        plcTme.Visible = False
        slSTime = edcFDropDown(0).Text
        slETime = edcFDropDown(1).Text
        If (Not gValidTime(slSTime)) Or (slSTime = "") Then
            ReDim tgCntSpot(0 To 0) As CNTSPOT
            lbcLines.Clear
            Exit Sub
        End If
        If (Not gValidTime(slETime)) Or (slETime = "") Then
            ReDim tgCntSpot(0 To 0) As CNTSPOT
            lbcLines.Clear
            Exit Sub
        End If
        'Use entered dates to obtain contracts
        slCntrSDate = edcFDropDown(2).Text
        slCntrEDate = edcFDropDown(3).Text
        If (Not gValidDate(slCntrSDate)) Or (slCntrSDate = "") Then
            ReDim tgCntSpot(0 To 0) As CNTSPOT
            lbcLines.Clear
            Exit Sub
        End If
        If (Not gValidDate(slCntrEDate)) Or (slCntrEDate = "") Then
            ReDim tgCntSpot(0 To 0) As CNTSPOT
            lbcLines.Clear
            Exit Sub
        End If
    Else
        slSTime = "12m"
        slETime = "12m"
        slCntrSDate = "1/1/1970"
        slCntrEDate = "12/31/2069"
    End If
    Screen.MousePointer = vbHourglass
    tmcLine.Enabled = False
    plcCalendar.Visible = False
    lbcLines.Clear
    ReDim tgCntSpot(0 To 0) As CNTSPOT
    llSTime = CLng(gTimeToCurrency(slSTime, False))
    llETime = CLng(gTimeToCurrency(slETime, True))
    'llSDate = lmSDate
    ''filter out spot only if not allowed to run on selected days (cff)
    ''mGetSpotSum removed code to filter out base on day selected (middle days)
    ''For ilDay = 0 To 6 Step 1
    ''    If ckcDay(ilDay).Value Then
    ''        Exit For
    ''    End If
    ''    llSDate = llSDate + 1
    ''Next ilDay
    'llEDate = lmEDate
    ''For ilDay = 6 To 0 Step -1
    ''    If ckcDay(ilDay).Value Then
    ''        Exit For
    ''    End If
    ''    llEDate = llEDate - 1
    ''Next ilDay
    '
    'Obtain spots from specified dates- Jim 9/12/00
    '
    'slSDate = Format$(lmSDate, "m/d/yy")
    'slEDate = Format$(lmEDate, "m/d/yy")
    If (ckcSpotType(0).Value = vbChecked) Or (ckcSpotType(1).Value = vbChecked) Or ((tgSpf.sSchdRemnant = "Y") And (ckcSpotType(5).Value = vbChecked)) Or ((tgSpf.sSchdPromo = "Y") And (ckcSpotType(3).Value = vbChecked)) Or ((tgSpf.sSchdPSA = "Y") And (ckcSpotType(2).Value = vbChecked)) Then
        If (Not imGameVehicle) Or (imMixtureOfVehicles > 0) Then
            For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                If lbcVehicle.Selected(ilLoop) Then
                    slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    mGetSpotSum Val(slCode), slCntrSDate, slCntrEDate, llSTime, llETime, 0
                End If
            Next ilLoop
        Else
            For ilLoop = 0 To lbcFromGame.ListItems.Count - 1 Step 1
                If lbcFromGame.ListItems(ilLoop + 1).Selected Then
                    ilGameNo = Val(lbcFromGame.ListItems(ilLoop + 1).Text)
                    slCntrSDate = lbcFromGame.ListItems(ilLoop + 1).SubItems(1)
                    slCntrEDate = slCntrSDate
                    mGetSpotSum imFromVefCode, slCntrSDate, slCntrEDate, llSTime, llETime, ilGameNo
                End If
            Next ilLoop
        End If
        mLinePop llSTime, llETime
    End If
    '
    'Obtain contracts (PSA; Promo; PI) from the specified dates- Jim 9/12/00
    '
    '5/5/11: Active Manual mode for games
    'If Not imGameVehicle Then
        slCntrType = ""
        If (ckcSpotType(2).Value = vbChecked) And (tgSpf.sSchdPSA <> "Y") Then
            slCntrType = slCntrType & "S"
        End If
        If (ckcSpotType(3).Value = vbChecked) And (tgSpf.sSchdPromo <> "Y") Then
            slCntrType = slCntrType & "M"
        End If
        If ckcSpotType(4).Value = vbChecked Then
            slCntrType = slCntrType & "Q"
        End If
        If (ckcSpotType(5).Value = vbChecked) And (tgSpf.sSchdRemnant <> "Y") Then
            slCntrType = slCntrType & "T"
        End If
        If (ckcSpotType(6).Value = vbChecked) Then
            slCntrType = slCntrType & "R"
        End If
        llCntrSDate = gDateValue(slCntrSDate)
        llCntrEDate = gDateValue(slCntrEDate)
        mManCntrPop slCntrType, llCntrSDate, llCntrEDate, llSTime, llETime
        If UBound(tgCntSpot) - 1 > 0 Then
            ArraySortTyp fnAV(tgCntSpot(), 0), UBound(tgCntSpot), 0, LenB(tgCntSpot(0)), 0, LenB(tgCntSpot(0).sKey), 0
        End If
    'End If
    mMakeSummary
    mLoadListBox
    Screen.MousePointer = vbDefault
End Sub
Private Sub plcAC_Paint()
    plcAC.CurrentX = 0
    plcAC.CurrentY = 0
    plcAC.Print "Advt/Competitives"
End Sub
Private Sub plcFillInv_Paint()
    plcFillInv.CurrentX = 0
    plcFillInv.CurrentY = 0
    plcFillInv.Print "Show on Inv"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Spot Fill"
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mTeamPop                        *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Tema list box         *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mTeamPop()
'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer
    ilRet = gObtainMnfForType("Z", smTeamTag, tmTeam())
    Exit Sub
End Sub

Private Sub mListColumnWidths()
    Dim ilCol As Integer
    Dim llWidth As Long

    lbcFromGame.ColumnHeaders.item(1).Width = lbcFromGame.Width / 7
    lbcFromGame.ColumnHeaders.item(2).Width = lbcFromGame.Width / 6
    For ilCol = 1 To 2 Step 1
        llWidth = llWidth + lbcFromGame.ColumnHeaders.item(ilCol).Width
    Next ilCol
    lbcFromGame.ColumnHeaders.item(3).Width = lbcFromGame.Width - llWidth - GRIDSCROLLWIDTH - 4 * 240 + 120
    llWidth = 0
    lbcToGame.ColumnHeaders.item(1).Width = lbcFromGame.Width / 7
    lbcToGame.ColumnHeaders.item(2).Width = lbcFromGame.Width / 6
    For ilCol = 1 To 2 Step 1
        llWidth = llWidth + lbcToGame.ColumnHeaders.item(ilCol).Width
    Next ilCol
    lbcToGame.ColumnHeaders.item(3).Width = (lbcToGame.Width - llWidth - GRIDSCROLLWIDTH - 5 * 240) / 2
    llWidth = llWidth + lbcToGame.ColumnHeaders.item(3).Width
    lbcToGame.ColumnHeaders.item(4).Width = lbcToGame.Width - llWidth - GRIDSCROLLWIDTH - 5 * 240 + 60
End Sub

Private Sub mChkXMid(llSTime As Long, llETime As Long, ilAllowedTimeIndex As Integer, llAllowedSTime As Long, llAllowedETime As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilDay                                                                                 *
'******************************************************************************************

    Dim ilUpper As Integer

    ilUpper = UBound(tgCntSpot)
    If llAllowedSTime <= llAllowedETime Then
        If (llETime >= llAllowedSTime) And (llSTime <= llAllowedETime) And (ilAllowedTimeIndex <= UBound(tgCntSpot(ilUpper).lAllowedSTime)) Then
            tgCntSpot(ilUpper).lAllowedSTime(ilAllowedTimeIndex) = llAllowedSTime
            tgCntSpot(ilUpper).lAllowedETime(ilAllowedTimeIndex) = llAllowedETime
            tgCntSpot(ilUpper).lPrice = tmCff.lActPrice
            ilAllowedTimeIndex = ilAllowedTimeIndex + 1
        End If
    Else
        If (llETime >= llAllowedSTime) And (llSTime <= 86400) And (ilAllowedTimeIndex <= UBound(tgCntSpot(ilUpper).lAllowedSTime)) Then
            tgCntSpot(ilUpper).lAllowedSTime(ilAllowedTimeIndex) = llAllowedSTime
            tgCntSpot(ilUpper).lAllowedETime(ilAllowedTimeIndex) = 86400
            ilAllowedTimeIndex = ilAllowedTimeIndex + 1
        End If
        If (llETime >= 0) And (llSTime <= llAllowedETime) And (ilAllowedTimeIndex <= UBound(tgCntSpot(ilUpper).lAllowedSTime)) Then
            tgCntSpot(ilUpper).lAllowedSTime(ilAllowedTimeIndex) = 0
            tgCntSpot(ilUpper).lAllowedETime(ilAllowedTimeIndex) = llAllowedETime
            ilAllowedTimeIndex = ilAllowedTimeIndex + 1
        End If
    End If
End Sub
