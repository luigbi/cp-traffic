VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExpStnFd 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   165
   ClientTop       =   1725
   ClientWidth     =   9495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   9495
   Begin VB.PictureBox plcSelect 
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
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   1
      Left            =   6015
      ScaleHeight     =   435
      ScaleWidth      =   3345
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   3345
      Begin VB.OptionButton rbcEnv 
         Caption         =   "New Only"
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
         Height          =   195
         Index           =   1
         Left            =   2190
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   1125
      End
      Begin VB.OptionButton rbcEnv 
         Caption         =   "All"
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
         Height          =   195
         Index           =   0
         Left            =   1425
         TabIndex        =   10
         Top             =   0
         Width           =   690
      End
      Begin VB.OptionButton rbcEnv 
         Caption         =   "Vantive Rules"
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
         Height          =   195
         Index           =   2
         Left            =   1425
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   210
         Value           =   -1  'True
         Width           =   1860
      End
   End
   Begin VB.PictureBox pbcInterface 
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
      Height          =   390
      Left            =   2520
      ScaleHeight     =   390
      ScaleWidth      =   2385
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   2385
      Begin VB.OptionButton rbcInterface 
         Caption         =   "KenCast"
         Height          =   195
         Index           =   1
         Left            =   1155
         TabIndex        =   3
         Top             =   210
         Width           =   1215
      End
      Begin VB.OptionButton rbcInterface 
         Caption         =   "StarGuide"
         Height          =   195
         Index           =   0
         Left            =   1155
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   8910
      Top             =   4590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".Txt"
      Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
      FilterIndex     =   1
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.PictureBox plcSelect 
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
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   0
      Left            =   5730
      ScaleHeight     =   435
      ScaleWidth      =   3645
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   3645
      Begin VB.OptionButton rbcGen 
         Caption         =   "All Instructions"
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
         Height          =   195
         Index           =   3
         Left            =   1980
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   195
         Width           =   1605
      End
      Begin VB.OptionButton rbcGen 
         Caption         =   "All Spots"
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
         Height          =   195
         Index           =   2
         Left            =   780
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   210
         Width           =   1065
      End
      Begin VB.OptionButton rbcGen 
         Caption         =   "Copy Feed"
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
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   5
         Top             =   0
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton rbcGen 
         Caption         =   "Regional Spots"
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
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   1530
      End
   End
   Begin VB.PictureBox plcRotInfo 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   2730
      ScaleHeight     =   1170
      ScaleWidth      =   6465
      TabIndex        =   60
      Top             =   3075
      Visible         =   0   'False
      Width           =   6525
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Rotation: #  by user name"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   65
         Top             =   45
         Width           =   6330
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Last Assignment Done: Date xx/xx/xx   Time xx:xx:xxam"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   62
         Top             =   720
         Width           =   6315
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Entered Date xx/xx/xx    Version Date xx/xx/xx   Modified xx Times"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   64
         Top             =   270
         Width           =   6330
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Date Range Assigned To:  Earliest xx/xx/xx   Latest xx/xx/xx"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   63
         Top             =   495
         Width           =   6330
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Bulk Feed: Send on xx/xx/xx"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   61
         Top             =   945
         Width           =   6315
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
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
      Left            =   2175
      TabIndex        =   46
      Top             =   5640
      Width           =   1050
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
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
      Left            =   855
      TabIndex        =   45
      Top             =   5640
      Width           =   1185
   End
   Begin VB.CommandButton cmcSuppress 
      Appearance      =   0  'Flat
      Caption         =   "&Suppress Rotation"
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
      Height          =   285
      Left            =   3360
      TabIndex        =   47
      Top             =   5640
      Width           =   1770
   End
   Begin VB.CommandButton cmcReSend 
      Appearance      =   0  'Flat
      Caption         =   "Resen&d Rotation..."
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
      Height          =   285
      Left            =   5235
      TabIndex        =   48
      Top             =   5640
      Width           =   1710
   End
   Begin VB.Timer tmcRot 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   8940
      Top             =   5235
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   2160
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   50
      Top             =   2040
      Visible         =   0   'False
      Width           =   1995
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
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Expstnfd.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   55
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   57
         Top             =   30
         Width           =   1305
      End
   End
   Begin VB.ListBox lbcVehicleCode 
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
      Left            =   9135
      Sorted          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox plcScreen 
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   45
      ScaleHeight     =   270
      ScaleWidth      =   2220
      TabIndex        =   0
      Top             =   0
      Width           =   2220
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8805
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   4710
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8790
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4860
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox plcDates 
      Height          =   4110
      Left            =   75
      ScaleHeight     =   4050
      ScaleWidth      =   9300
      TabIndex        =   13
      Top             =   480
      Width           =   9360
      Begin VB.PictureBox plcCmmlLog 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3405
         ScaleHeight     =   225
         ScaleWidth      =   2700
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1035
         Visible         =   0   'False
         Width           =   2700
         Begin VB.CheckBox ckcCmmlLog 
            Caption         =   "Text"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   2025
            TabIndex        =   39
            Top             =   0
            Width           =   660
         End
         Begin VB.CheckBox ckcCmmlLog 
            Caption         =   "PDF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   1395
            TabIndex        =   38
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lacCmmlLog 
            Caption         =   "Commercial Log"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   1395
         End
      End
      Begin VB.PictureBox plcFormat 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   3285
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1035
         Visible         =   0   'False
         Width           =   3285
         Begin VB.OptionButton rbcFormat 
            Caption         =   "Text"
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
            Index           =   3
            Left            =   2415
            TabIndex        =   35
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton rbcFormat 
            Caption         =   "PDF"
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
            Left            =   1800
            TabIndex        =   34
            Top             =   0
            Value           =   -1  'True
            Width           =   690
         End
      End
      Begin VB.CommandButton cmcGetRot 
         Appearance      =   0  'Flat
         Caption         =   "&Get Rotation"
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
         Left            =   6675
         TabIndex        =   42
         Top             =   1380
         Width           =   2025
      End
      Begin VB.PictureBox pbcLbcRot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
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
         Height          =   2310
         Left            =   45
         ScaleHeight     =   2310
         ScaleWidth      =   8940
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1695
         Width           =   8940
      End
      Begin VB.PictureBox plcFormat 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   3990
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1035
         Visible         =   0   'False
         Width           =   3990
         Begin VB.OptionButton rbcFormat 
            Caption         =   "Acrobat PDF"
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
            Left            =   1275
            TabIndex        =   31
            Top             =   0
            Width           =   1290
         End
         Begin VB.OptionButton rbcFormat 
            Caption         =   "Plain Text"
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
            Left            =   2625
            TabIndex        =   32
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.TextBox edcRunLetter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1860
         MaxLength       =   1
         TabIndex        =   19
         Top             =   510
         Width           =   345
      End
      Begin VB.CommandButton cmcFrom 
         Appearance      =   0  'Flat
         Caption         =   "&Browse..."
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
         Left            =   4950
         TabIndex        =   17
         Top             =   135
         Width           =   1005
      End
      Begin VB.PictureBox plcFrom 
         Height          =   375
         Left            =   1140
         ScaleHeight     =   315
         ScaleWidth      =   3675
         TabIndex        =   15
         Top             =   90
         Width           =   3735
         Begin VB.TextBox edcFrom 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   30
            TabIndex        =   16
            Top             =   30
            Width           =   3615
         End
      End
      Begin VB.CommandButton cmcEndDate 
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
         Left            =   4335
         Picture         =   "Expstnfd.frx":2E1A
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   780
         Width           =   195
      End
      Begin VB.TextBox edcEndDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3390
         MaxLength       =   10
         TabIndex        =   27
         Top             =   780
         Width           =   930
      End
      Begin VB.CommandButton cmcStartDate 
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
         Left            =   2010
         Picture         =   "Expstnfd.frx":2F14
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   780
         Width           =   195
      End
      Begin VB.TextBox edcStartDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   24
         Top             =   780
         Width           =   930
      End
      Begin VB.ListBox lbcVeh 
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
         Height          =   1290
         Left            =   6120
         MultiSelect     =   2  'Extended
         TabIndex        =   40
         Top             =   45
         Width           =   3165
      End
      Begin VB.TextBox edcTranDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4740
         MaxLength       =   10
         TabIndex        =   21
         Top             =   510
         Width           =   930
      End
      Begin VB.CommandButton cmcTranDate 
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
         Left            =   5685
         Picture         =   "Expstnfd.frx":300E
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   510
         Width           =   195
      End
      Begin VB.ListBox lbcRot 
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
         Height          =   2340
         ItemData        =   "Expstnfd.frx":3108
         Left            =   30
         List            =   "Expstnfd.frx":310A
         MultiSelect     =   2  'Extended
         TabIndex        =   44
         Top             =   1680
         Width           =   8970
      End
      Begin VB.VScrollBar vbcRot 
         Height          =   2340
         LargeChange     =   10
         Left            =   8985
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1680
         Width           =   270
      End
      Begin VB.CheckBox ckcAll 
         Caption         =   "All Rotations"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   45
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1455
         Width           =   1410
      End
      Begin VB.PictureBox pbcDateTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   9060
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   29
         Top             =   1380
         Width           =   105
      End
      Begin VB.ListBox lbcRegVeh 
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
         Height          =   1290
         Left            =   6120
         MultiSelect     =   2  'Extended
         TabIndex        =   41
         Top             =   45
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.Label lacRunLetter 
         Appearance      =   0  'Flat
         Caption         =   "Run Letter (A or B)"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   18
         Top             =   495
         Width           =   1725
      End
      Begin VB.Label lbcFrom 
         Appearance      =   0  'Flat
         Caption         =   "Station File"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label lacEndDate 
         Appearance      =   0  'Flat
         Caption         =   "End Date"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2415
         TabIndex        =   26
         Top             =   780
         Width           =   960
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   23
         Top             =   780
         Width           =   945
      End
      Begin VB.Label lacProcessing 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   49
         Top             =   1260
         Width           =   5715
      End
      Begin VB.Label lacTranDate 
         Appearance      =   0  'Flat
         Caption         =   "Transmission Date"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2940
         TabIndex        =   20
         Top             =   495
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmcCopy 
      Appearance      =   0  'Flat
      Caption         =   "C&opy Export..."
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
      Left            =   7080
      TabIndex        =   68
      Top             =   5640
      Width           =   1350
   End
   Begin VB.PictureBox pbcLbcVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   840
      Left            =   660
      ScaleHeight     =   840
      ScaleWidth      =   7815
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   4680
      Width           =   7815
   End
   Begin VB.ListBox lbcVehicle 
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
      Height          =   870
      Left            =   645
      Sorted          =   -1  'True
      TabIndex        =   67
      Top             =   4665
      Width           =   8130
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   90
      Top             =   4665
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "ExpStnFd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpStnFd.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the agency conversion input screen code
'
'
'Note:
'   tmVef and tmVpfInfo are matched and contain the following vehicles:
'   Conventional without Bulk Groups
'   Conventional with Bulk Groups (first vehicle of group alphabetically, remaining vehicles in group are in tmLkVehInfo)
'   Airing without Bulk Groups
'   Airing with Bulk Groups (first vehicle of group alphabetically, remaining vehicles in group are in tmLkVehInfo)
'   Selling
'
'   Only conventional and airing vehicles have groups.
'   Bulk Groups obtained from vehicle options interface area
'   tmVef and tmVpfInfo build in mVehPop
'   Populate lbcVehicle from Airing and Conventional from tmVef
'   Populate lbvVeh by call to gPopUserVehicleBox (with conventional; selling and airing vehicles) in mInit
'
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim tmVehCode() As SORTCODE
Dim smVehCodeTag As String
Dim imVefCode() As Integer
Dim imDVefCode() As Integer
Dim imRegVefCode() As Integer   'List of vehicles that have regions
Dim imExpSpotVefCode() As Integer
Dim imAVefCode() As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmExport As Integer   'file hanle
Dim hmMsg As Integer        'Message File Handle
Dim smFileNames() As String 'Use if abort issued
Dim tmPSAPromoSortCrf() As SORTCRF
Dim tmSvSortCrf() As SORTCRF
Dim lmReadyCRF() As Long    'CRF with crf.sAffFdStatus of Ready to send (need to retain as each vehicle/station is processed this status is changed to S (Sent)
Dim imNoTimesMod() As Integer    'CRF with crf.sAffFdStatus of Ready to send (need to retain as each vehicle/station is processed this status is changed to S (Sent)
'5/31/06:  Show comments on first occurrance only
Dim smRotComment() As String
Dim smTimeRestrictions() As String
Dim smDayRestrictions() As String
'6/14/06
Dim tmDuplComment() As DUPLCOMMENT
'Contract header
Dim tmChfSrchKey As LONGKEY0  'CHF key record image
Dim hmCHF As Integer        'CHF Handle
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF
'Rotation header
Dim tmCrfSrchKey As LONGKEY0  'CRF key record image
Dim tmCrfSrchKey1 As CRFKEY1  'CRF key record image
Dim hmCrf As Integer        'CRF Handle
Dim imCrfRecLen As Integer      'CRF record length
Dim tmCrf As CRF
Dim tmSTCrf() As CRF        'Supersede test
Dim imSelectPrevState() As Integer
Dim lmRotCodeBuild() As Long 'Rotation codes build for export- aviod sending same twicw
'Short Title
Dim tmSif As SIF            'SIF record image
Dim tmSifSrchKey As LONGKEY0  'SIF key record image
Dim hmSif As Integer        'SIF Handle
Dim imSifRecLen As Integer      'SIF record length
'Short Title via Contract
Dim tmVsf As VSF            'VSF record image
Dim tmVsfSrchKey As LONGKEY0  'VSF key record image
Dim hmVsf As Integer        'VSF Handle
Dim imVsfRecLen As Integer      'VSF record length
'Instruction
Dim tmCnf As CNF            'CNF record image
Dim tmCnfSrchKey As CNFKEY0  'CNF key record image
Dim hmCnf As Integer        'CNF Handle
Dim imCnfRecLen As Integer      'CNF record length
Dim tmCnfRot() As CNF
'Media Code
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey As INTKEY0  'MCF key record image
Dim hmMcf As Integer        'MCF Handle
Dim imMcfRecLen As Integer      'MCF record length
'Inventory
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0  'CIF key record image
Dim hmCif As Integer        'CIF Handle
Dim imCifRecLen As Integer      'CIF record length
'Copy Feed
Dim tmCyf As CYF            'CYF record image
Dim tmCyfSrchKey As CYFKEY0  'CYF key record image
Dim hmCyf As Integer        'CYF Handle
Dim imCyfRecLen As Integer      'CYF record length
Dim tmCyfTest() As CYFTEST
'Blackout
Dim tmBof As BOF
Dim tmBofSrchKey As LONGKEY0
Dim hmBof As Integer
Dim imBofRecLen As Integer
'Product
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0  'CPF key record image
Dim hmCpf As Integer        'CPF Handle
Dim imCpfRecLen As Integer      'CPF record length
'Copy Report- used as temporary storage of spots
Dim tmCpr As CPR            'CPR record image
Dim tmCprSrchKey As CPRKEY0  'CPR key record image
Dim hmCpr As Integer        'CPR Handle
Dim imCprRecLen As Integer      'CPR record length
'Text Report- used as temporary storage of spots
Dim tmTxr As TXR            'TXR record image
Dim tmTxrSrchKey As TXRKEY0  'TXR key record image
Dim hmTxr As Integer        'TXR Handle
Dim imTxrRecLen As Integer      'TXR record length
'Comment
Dim tmCsf As CSF            'CSF record image
Dim tmCsfSrchKey As LONGKEY0  'CSF key record image
Dim hmCsf As Integer        'CSF Handle
Dim imCsfRecLen As Integer      'CSF record length
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0 'ADF key record image
Dim imAdfRecLen As Integer  'ADF record length
'Avail name
Dim hmAnf As Integer
Dim tmAnf As ADF
Dim tmAnfSrchKey As INTKEY0 'ANF key record image
Dim imAnfRecLen As Integer  'ANF record length
'Vehicle
Dim hmVef As Integer
Dim tmVef() As VEF
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
Dim imNoRotations As Integer
Dim tmVpfInfo() As VPFINFO
Dim tmLkVehInfo() As LKVEHINFO
Dim tmSALink() As LKVEHINFO
Dim smVehName As String
'Vehicle link
Dim tmVlf As VLF            'VLF record image
Dim hmVlf As Integer        'VLF Handle
Dim imVlfRecLen As Integer      'VLF record length
'Region Code
Dim tmRaf As RAF            'RAF record image
Dim tmRafSrchKey As LONGKEY0  'MCF key record image
Dim tmRafSrchKey2 As LONGKEY0  'MCF key record image
Dim hmRaf As Integer        'MCF Handle
Dim imRafRecLen As Integer      'MCF record length
'Region Schd Copy
Dim tmRsf As RSF            'RSF record image
Dim tmRsfSrchKey1 As LONGKEY0  'RSF key record image
Dim hmRsf As Integer        'RSF Handle
Dim imRsfRecLen As Integer      'RSF record length
'Station Information
Dim smStationFile As String
Dim hmStationFile As Integer
Dim smFieldValues(0 To 30) As String    '30 fields generated in a record, values aged from 0-29 to 1-30
'Field description
'  1
'  2    Vehicle Name
'  3    Region Name
'  4    Region Code
'  5    Call Letters
'  6    Band
'  7    Site ID
'  8    City
'  9    State
' 10    EDAS
' 11    Transportal
' 12    KenCast Address
' 13    Feed Zone
' 14    Zone
' 15
' 16    Number of Airplays
' 17    Commercial Log (L=Generate Commercial Log; N= Don't Generate)
' 18-27 Commercial Log Affiliate Pledge Times
' 28    KenCast Envelope Copy (A=All Copy; N=New Copy)
' 29    Commercial Log Show Daypart (S=Show Daypart; A=Show Affiliate Times {fields 18-27})
' 30    Commercial Log Cart #'s (C=Display cart #'s; N= Don't display cart #'s)
Dim imListFieldRot(0 To 12) As Integer
Dim imListFieldVeh(0 To 3) As Integer
Dim tmStnInfo() As STNINFO
Dim smTranDate As String
Dim smFeedNo As String
Dim smGenTime As String     'Generate Time (HH:MM)
Dim smGenDate As String     'Generate Date (MM/DD/YYYY)
Dim smPDFTime As String     'PDF Generate Time (HH:MM:SS)
Dim smPDFDate As String     'PDF Generate Date (MM/DD/YYYY)
Dim imPDFDate(0 To 1) As Integer
Dim imPDFTime(0 To 1) As Integer
Dim lmPDFSeqNo As Long      'Sequence Number
Dim imCartStnXRef1 As Integer

'Dim tmRec As LPOPREC
'Regional Spots
'Dim hmTo As Integer
Dim tmEVef As VEF
Dim tmAVef As VEF
'Dim tmSVef As VEF
Dim hmSsf As Integer
Dim hmCTSsf As Integer
Dim tmCTSsf As SSF               'Ssf for conflict test
'Dim tmSsfOld As SSF
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmAvailTest As AVAILSS
'Spot record
Dim tmSdf As SDF
Dim hmSdf As Integer
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey1 As SDFKEY1
Dim tmSdfSrchKey3 As LONGKEY0
Dim hmTzf As Integer        'Time zone Copy file handle
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfRecLen As Integer  'TZF record length
Dim tmTzf As TZF            'TZF record image
'Contract Line record information
Dim hmClf As Integer        'Contract Line file handle
Dim imClfRecLen As Integer
Dim tmClf As CLF
Dim tmClfSrchKey As CLFKEY0

Dim imCopyMissing As Integer
Dim imEvtType(0 To 14) As Integer
Dim tmVlfSrchKey1 As VLFKEY1 'VLF key record image
'Delivery file (DLF)
Dim hmDlf As Integer        'Delivery link file
Dim imDlfRecLen As Integer  'DLF record length
Dim tmDlfSrchKey As DLFKEY0 'DLF key record image
Dim tmDlf As DLF            'DLF record image
Dim lmInputStartDate As Long    'Input Start Date
Dim lmInputEndDate As Long  'Input End Date
Dim smWeekNo As String      'Week number from start of the broadcast year
Dim smAllInstFileDate As String
Dim smRunLetter As String
Dim imExptPrevWeek As Integer   'Export Previous Week if not sent
Dim imIgnoreCkcAll As Integer
Dim imLastIndex As Integer
Dim imCurrentIndex As Integer
Dim imShiftKey As Integer   'Bit 0=Shift; 1=Ctrl; 2=Alt
Dim imButton As Integer
Dim imButtonIndex As Integer
Dim imIgnoreRightMove As Integer
Dim imIgnoreVbcChg As Integer
Dim imTypeIndex As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smTodaysDate As String  'mm/dd/yy"
Dim imGenDate(0 To 1) As Integer
Dim imGenTime(0 To 1) As Integer
Dim tmVehTimes() As VEHTIMES
Dim tmRotInfo() As SENDROTINFO
Dim tmAddCyf() As SENDCOPYINFO
Dim tmXRefCyf() As SENDCOPYINFO   'Merge of each tmAddCyf for xref
Dim tmWemCyf() As SENDCOPYINFO
Dim imShowHelpMsg As Integer    'True=Show help message; False=Ignore help message system
'Btrieve wait
Dim imWaitCount As Integer
Dim imTimeDelay As Integer
Dim imLockValue As Integer
Dim imTranLog As Integer
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA  'index zero ignored
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imDateBox As Integer    '1=Start Date; 2=End Date; 3=Transmission Date
' MsgBox parameters
Const vbOkOnly = 0                 ' OK button only
Const vbCritical = 16          ' Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0
Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilValue As Integer
    Dim llRg As Long
    Dim llRet As Long
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slTime As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    If imIgnoreCkcAll Then
        Exit Sub
    End If
    imIgnoreVbcChg = True
    If lbcRot.ListCount <= 0 Then
        If Value Then
            ckcAll.Value = vbUnchecked
        End If
        imIgnoreVbcChg = False
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilValue = Value
    If UBound(tgSortCrf) < vbcRot.LargeChange + 1 Then
        llRg = CLng(UBound(tgSortCrf) - 1) * &H10000 Or 0
    Else
        llRg = CLng(vbcRot.LargeChange) * &H10000 Or 0
    End If
    llRet = SendMessageByNum(lbcRot.HWnd, LB_SELITEMRANGE, ilValue, llRg)
    DoEvents
    ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
    For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
        tgSortCrf(ilLoop).iSelected = Value
    Next ilLoop
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        slNameCode = lbcVehicle.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "|", slTime)
        slTime = "0:0"
        lbcVehicle.List(ilLoop) = slName & "|" & slTime
    Next ilLoop
    If Value Then
        For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
            tmCrf = tgSortCrf(ilLoop).tCrf
            mComputeTime True
            ilIndex = tgSortCrf(ilLoop).iCombineIndex
            Do While ilIndex >= 0
                tmCrf = tgCombineCrf(ilIndex).tCrf
                mComputeTime True
                ilIndex = tgCombineCrf(ilIndex).iCombineIndex
            Loop
            ilIndex = tgSortCrf(ilLoop).iDuplIndex
            Do While ilIndex >= 0
                tmCrf = tgDuplCrf(ilIndex).tCrf
                mComputeTime True
                ilIndex = tgDuplCrf(ilIndex).iDuplIndex
            Loop
            DoEvents
        Next ilLoop
    End If
    pbcLbcRot_Paint
    pbclbcVehicle_Paint
    mSetCommands
    Screen.MousePointer = vbDefault
    imIgnoreVbcChg = False
End Sub
Private Sub ckcAll_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcCalDn_Click()
    If imDateBox = 1 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
        edcStartDate.SetFocus
    ElseIf imDateBox = 2 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
        edcEndDate.SetFocus
    ElseIf imDateBox = 3 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcTranDate.SelStart = 0
        edcTranDate.SelLength = Len(edcTranDate.Text)
        edcTranDate.SetFocus
    End If
End Sub
Private Sub cmcCalUp_Click()
    If imDateBox = 1 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
        edcStartDate.SetFocus
    ElseIf imDateBox = 2 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
        edcEndDate.SetFocus
    ElseIf imDateBox = 3 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcTranDate.SelStart = 0
        edcTranDate.SelLength = Len(edcTranDate.Text)
        edcTranDate.SetFocus
    End If
End Sub
Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcCopy_Click()
    StnFdCpy.Show vbModal
    plcDates.Visible = False
    plcDates.Visible = True
    If rbcInterface(0).Value Then
        plcSelect(0).Visible = False
        plcSelect(0).Visible = True
    Else
        plcSelect(1).Visible = False
        plcSelect(1).Visible = True
    End If
End Sub
Private Sub cmcCopy_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcEndDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
End Sub
Private Sub cmcEndDate_GotFocus()
    Dim slStr As String
    'If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
    imFirstFocus = False
    '    'Show branner
    '    mInitDDE
    '    mSendHelpMsg "BT"
    'End If
    If imDateBox <> 2 Then
        plcCalendar.Visible = False
        slStr = edcEndDate.Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    imDateBox = 2
    plcCalendar.Move plcDates.Left + edcEndDate.Left, plcDates.Top + edcEndDate.Top + edcEndDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcExport_Click()
    Dim slStr As String
    Dim slMissingCopyNames As String
    Dim slName As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilLink As Integer
    Dim ilVefCode As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slFYear As String
    Dim slFMonth As String
    Dim ilUpper As Integer
    Dim slFDay As String
    Dim slDate As String
    Dim slTime As String
    Dim ilFirstLastWk As Integer
    Dim ilWkNo As Integer
    Dim ilWem As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFound As Integer

    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    
    If (Not rbcInterface(0).Value) And (Not rbcInterface(1).Value) Then
        ''MsgBox "Station File Interface Must be Defined", vbExclamation, "Name Error"
        gAutomationAlertAndLogHandler "Station File Interface Must be Defined", vbExclamation, "Name Error"
        Exit Sub
    End If
    If (rbcInterface(0).Value) And (rbcGen(3).Value) Then
        smStationFile = ""
        smRunLetter = "A"
    Else
        smStationFile = Trim$(edcFrom.Text)
        If smStationFile = "" Then
            ''MsgBox "Station File Must be Defined", vbExclamation, "Name Error"
            gAutomationAlertAndLogHandler "Station File Must be Defined", vbExclamation, "Name Error"
            edcFrom.SetFocus
            Exit Sub
        End If
        smRunLetter = Trim$(edcRunLetter.Text)
        If smRunLetter = "" Then
            ''MsgBox "Run Letter Must be Defined", vbExclamation, "Name Error"
            gAutomationAlertAndLogHandler "Run Letter Must be Defined", vbExclamation, "Name Error"
            edcRunLetter.SetFocus
            Exit Sub
        End If
    End If
    If ((rbcInterface(0).Value) And ((rbcGen(0).Value) Or (rbcGen(3).Value))) Or (rbcInterface(1).Value) Then
        slStr = edcTranDate.Text
        If Not gValidDate(slStr) Then
            Beep
            edcTranDate.SetFocus
            Exit Sub
        End If
    End If
    slStr = Trim$(edcStartDate.Text)
    If (Not gValidDate(slStr)) Or (slStr = "") Then
        Beep
        edcStartDate.SetFocus
        Exit Sub
    End If
    If (rbcInterface(0).Value) And (rbcGen(0).Value) Or (rbcInterface(1).Value) Then
        slStr = gObtainPrevMonday(slStr)
    End If
    slStartDate = slStr
    lmInputStartDate = gDateValue(slStr)
    gObtainWkNo 0, slStr, ilWkNo, ilFirstLastWk
    smWeekNo = Trim$(str$(ilWkNo))
    Do While Len(smWeekNo) < 2
        smWeekNo = "0" & smWeekNo
    Loop
    slStr = Trim$(edcEndDate.Text)
    slEndDate = slStr
    If (Not gValidDate(slStr)) Or (slStr = "") Then
        Beep
        edcEndDate.SetFocus
        Exit Sub
    End If
    lmInputEndDate = gDateValue(slStr)
    If ((rbcInterface(0).Value) And (Not rbcGen(3).Value)) Or (rbcInterface(1).Value) Then
        If lmInputStartDate + 7 <= lmInputEndDate Then
            Beep
            edcEndDate.SetFocus
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    If Not mOpenMsgFile() Then
        Screen.MousePointer = vbDefault
        cmcCancel.SetFocus
        Exit Sub
    End If
    imExptPrevWeek = True
    If ((rbcInterface(0).Value) And (rbcGen(0).Value) Or (rbcGen(3).Value)) Or (rbcInterface(1).Value) Then
        slStr = edcTranDate.Text
        smTranDate = Format$(gDateValue(slStr), "mm/dd/yyyy")
        gObtainYearMonthDayStr smTranDate, True, slFYear, slFMonth, slFDay
        smFeedNo = right$(slFYear, 2) & slFMonth & slFDay
        If (rbcInterface(0).Value) And (rbcGen(3).Value) Then
            imExptPrevWeek = False
            'If Val(slFMonth) <= 9 Then
            '    slFMonth = right$(slFMonth, 1)
            'ElseIf Val(slFMonth) = 10 Then
            '    slFMonth = "A"
            'ElseIf Val(slFMonth) = 11 Then
            '    slFMonth = "B"
            'ElseIf Val(slFMonth) = 12 Then
            '    slFMonth = "C"
            'End If
            'smAllInstFileDate = right(slFYear, 2) & slFMonth & slFDay
            smAllInstFileDate = slFMonth & slFDay
        End If
        If ((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
            hmTxr = CBtrvTable(TEMPHANDLE) 'CBtrvObj
            ilRet = btrOpen(hmTxr, "", sgDBPath & "Txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                'Print #hmMsg, "Export Terminated as TXR.BTR could not be Opened, error #" & str$(ilRet)
                gAutomationAlertAndLogHandler "Export Terminated as TXR.BTR could not be Opened, error #" & str$(ilRet)
                Close #hmMsg
                On Error GoTo 0
                cmcCancel.SetFocus
                Exit Sub
            End If
            imTxrRecLen = Len(tmTxr)
            If imTxrRecLen <> btrRecordLength(hmTxr) Then
                Screen.MousePointer = vbDefault
                btrDestroy hmTxr
                'Print #hmMsg, "Export Terminated as TXR.BTR size is not matching, Internal" & str$(imTxrRecLen) & "vs External" & str$(btrRecordLength(hmTxr))
                gAutomationAlertAndLogHandler "Export Terminated as TXR.BTR size is not matching, Internal" & str$(imTxrRecLen) & "vs External" & str$(btrRecordLength(hmTxr))
                Close #hmMsg
                On Error GoTo 0
                cmcCancel.SetFocus
                Exit Sub
            End If
        End If
    Else
        smTranDate = Format$(gNow(), "mm/dd/yyyy")
        slStr = edcStartDate.Text
        slStr = Format$(gDateValue(slStr), "mm/dd/yyyy")
        gObtainYearMonthDayStr slStr, True, slFYear, slFMonth, slFDay
        smFeedNo = slFDay
    End If
    If (rbcInterface(0).Value) And (rbcGen(3).Value) Then
        ReDim tmStnInfo(0 To 0) As STNINFO
    Else
        lacProcessing.Caption = "Reading Station Information File"
        If Not mGetStnInfo(True) Then
            Screen.MousePointer = vbDefault
            If (rbcInterface(0).Value) And (rbcGen(1).Value) Then
                'MsgBox "Regional Spots not complete, check " & sgDBPath & "Messages\" & "ExpRgSpt.Txt", vbOkOnly + vbCritical + vbApplicationModal, "Station Error"
                gAutomationAlertAndLogHandler "Regional Spots not complete, check " & sgDBPath & "Messages\" & "ExpRgSpt.Txt", vbOkOnly + vbCritical + vbApplicationModal, "Station Error"
            ElseIf (rbcInterface(0).Value) And (rbcGen(2).Value) Then
                'MsgBox "All Spots not complete, check " & sgDBPath & "Messages\" & "ExpAlSpt.Txt", vbOkOnly + vbCritical + vbApplicationModal, "Station Error"
                gAutomationAlertAndLogHandler "All Spots not complete, check " & sgDBPath & "Messages\" & "ExpAlSpt.Txt", vbOkOnly + vbCritical + vbApplicationModal, "Station Error"
            Else
                'MsgBox "Station Feed not complete, check " & sgDBPath & "Messages\" & "ExpStnFd.Txt", vbOkOnly + vbCritical + vbApplicationModal, "Station Error"
                gAutomationAlertAndLogHandler "Station Feed not complete, check " & sgDBPath & "Messages\" & "ExpStnFd.Txt", vbOkOnly + vbCritical + vbApplicationModal, "Station Error"
                If ((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
                    ilRet = btrClose(hmTxr)
                    btrDestroy hmTxr
                End If
            End If
            'Print #hmMsg, "Export Terminated as Station Information is not Complete"
            gAutomationAlertAndLogHandler "Export Terminated as Station Information is not Complete"
            Close #hmMsg
            On Error GoTo 0
            lacProcessing.Caption = "See: " & sgDBPath & "Messages\" & "ExpStnFd.Txt" & " for Messages"
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    'Print #hmMsg, "   Transmission Date: " & smTranDate
    gAutomationAlertAndLogHandler "   Transmission Date: " & smTranDate
    'Print #hmMsg, "   Start Date: " & slStartDate & " End Date: " & slEndDate
    gAutomationAlertAndLogHandler "   Start Date: " & slStartDate & " End Date: " & slEndDate
    If ((rbcInterface(0).Value) And (rbcGen(0).Value)) Or (rbcInterface(1).Value) Then
        lacProcessing.Caption = "Checking Station Information File"
        If Not mCheckRegions() Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Copy Rotation Not Defined for All Regions(see Accessories->Messages->Station Feed - Export Station Feed), Continue Anyway", vbYesNo + vbQuestion, "Region Error")
            If ilRet = vbNo Then
                'Print #hmMsg, "Export Terminated as Copy Rotations missing for Regions"
                gAutomationAlertAndLogHandler "Export Terminated as Copy Rotations missing for Regions"
                Close #hmMsg
                If ((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
                    ilRet = btrClose(hmTxr)
                    btrDestroy hmTxr
                End If
                On Error GoTo 0
                cmcCancel.SetFocus
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
        End If
    End If
    If (rbcInterface(1).Value) Then
        lacProcessing.Caption = "Checking Vehicle Information"
        If Not mCheckVehicles() Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Vehicles defined in Station Information but not Selected and/or Vehicles Selected but not in Station Information (see Accessories->Messages->Station Feed - Export Station Feed), Continue Anyway", vbYesNo + vbQuestion, "Vehicle Information Error")
            If ilRet = vbNo Then
                'Print #hmMsg, "Export Terminated as Vehicles not matching"
                gAutomationAlertAndLogHandler "Export Terminated as Vehicles not matching"
                Close #hmMsg
                If ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
                    ilRet = btrClose(hmTxr)
                    btrDestroy hmTxr
                End If
                On Error GoTo 0
                cmcCancel.SetFocus
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
        End If
    End If
    'Set File Name for KenCast
    If (rbcInterface(1).Value) Then
        For ilLoop = 0 To UBound(tmStnInfo) - 1 Step 1
            If Trim$(tmStnInfo(ilLoop).sFileName) = "" Then
                tmStnInfo(ilLoop).sFileName = Trim$(tmStnInfo(ilLoop).sCallLetter) & "-" & Trim$(tmStnInfo(ilLoop).sBand)
            End If
        Next ilLoop
    End If
    lacProcessing.Caption = ""
    imExporting = True
    smGenTime = Format$(gNow(), "hh:mm")
    smGenDate = Format$(gNow(), "mm/dd/yyyy")
    'If rbcGen(1).Value Then
    If (rbcInterface(0).Value) And ((rbcGen(1).Value) Or (rbcGen(2).Value)) Then
'Moved below the If statement so rbcInterface(1) can use the same code
'        slDate = Format$(gNow(), "m/d/yy")
'        gPackDate slDate, imGenDate(0), imGenDate(1)
'        slTime = Format$(gNow(), "h:mm:ssAM/PM")
'        gPackTime slTime, imGenTime(0), imGenTime(1)
'        'Print #hmMsg, "** Storing Output into " & slToFile & " **"
'        slMissingCopyNames = ""
'        For ilLoop = 0 To lbcRegVeh.ListCount - 1 Step 1
'            If lbcRegVeh.Selected(ilLoop) Then
'                imCopyMissing = False
'                'slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
'                'ilRet = gParseItem(slNameCode, 1, "\", slName)
'                'ilRet = gParseItem(slName, 3, "|", slName)
'                'smVehName = Trim$(slName)
'                'Print #hmMsg, "** Generating Data for " & Trim$(slName) & " **"
'                'lacProcessing.Caption = "Generating Data for " & Trim$(slName)
'                'ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                'ilVefCode = Val(slCode)
'                'ilFound = False
'                'For ilVeh = 0 To UBound(imRegVefCode) - 1 Step 1
'                '    If imRegVefCode(ilVeh) = ilVefCode Then
'                '        ilFound = True
'                '        Exit For
'                '    End If
'                'Next ilVeh
'                ilVefCode = imRegVefCode(ilLoop)
'                'tmVefSrchKey.iCode = ilVefCode
'                'ilRet = btrGetEqual(hmVef, tmEVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'                ilRet = gBinarySearchVef(ilVefCode)
'                If ilRet <> -1 Then
'                    tmEVef = tgMVef(ilRet)
'                End If
'                If (ilRet <> -1) Then
'                    smVehName = Trim$(tmEVef.sName)
'                    slName = smVehName
'                    Print #hmMsg, "** Generating Data for " & Trim$(slName) & " **"
'                    lacProcessing.Caption = "Generating Data for " & Trim$(slName)
'                    If Not mExpSpots("O", "C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType()) Then
'                        mClearCPR
'                        Print #hmMsg, "** Terminated **"
'                        Close #hmMsg
'                        Close #hmTo
'                        imExporting = False
'                        'MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
'                        Screen.MousePointer = vbDefault
'                        cmcCancel.SetFocus
'                        Exit Sub
'                    End If
'                    Print #hmMsg, "** Completed " & Trim$(tmEVef.sName) & " **"
'                Else
'                End If
'            End If
'        Next ilLoop
'        ilRet = mCreateSchFile()
'        mClearCPR
'        Close #hmTo
'        Print #hmMsg, "** Completed Export Regional Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
'        Close #hmMsg
'        On Error GoTo 0
    Else
        gPackDate slDate, imGenDate(0), imGenDate(1)
        slTime = Format$(gNow(), "h:mm:ssAM/PM")
        ReDim tgAddCyf(0 To 0) As SENDCOPYINFO
        ReDim tmXRefCyf(0 To 0) As SENDCOPYINFO
        ReDim lmReadyCRF(0 To 0) As Long
        ReDim imNoTimesMod(0 To 0) As Integer
        'ReDim tgCartStnXRef(0 To 0) As CARTSTNXREF
        ReDim tgCartStnXRef(0 To 32000, 0 To 0) As CARTSTNXREF
        imCartStnXRef1 = 0
        If mNonRotFileNames() Then
            'If rbcFormat(0).Value Then
            '    gSwitchToPDF cdcSetup, 0
            'End If
            For ilLoop = 0 To UBound(tmStnInfo) - 1 Step 1
                If Trim$(tmStnInfo(ilLoop).sFileName) <> "" Then
                    mBuildExpTable tmStnInfo(ilLoop)
                    ilUpper = UBound(tmAddCyf)
                    If ilUpper > 0 Then
                        'ArraySortTyp fnAV(tgSort(),0), ilUpper, 0, LenB(tgSort(0)), 0, -9, 0
                        ArraySortTyp fnAV(tmAddCyf(), 0), ilUpper, 0, LenB(tmAddCyf(0)), 0, LenB(tmAddCyf(0).sKey), 0
                    End If
                    ilUpper = UBound(tmRotInfo)
                    If ilUpper > 0 Then
                        'ArraySortTyp fnAV(tgSort(),0), ilUpper, 0, LenB(tgSort(0)), 0, -9, 0
                        ArraySortTyp fnAV(tmRotInfo(), 0), ilUpper, 0, LenB(tmRotInfo(0)), 0, LenB(tmRotInfo(0).sKey), 0
                    End If
                    ilRet = mExpRot(tmStnInfo(ilLoop))
                    If Not ilRet Then
                        Screen.MousePointer = vbDefault
                        imExporting = False
                        'If rbcFormat(0).Value Then
                        '    gSwitchToPDF cdcSetup, 1
                        'End If
                        ''MsgBox "Station Feed Failed, Export terminated", vbOkOnly + vbCritical + vbApplicationModal, "Export"
                        gAutomationAlertAndLogHandler "Station Feed Failed, Export terminated", vbOkOnly + vbCritical + vbApplicationModal, "Export"
                        'Print #hmMsg, "Station Feed Failed, Export terminated"
                        gAutomationAlertAndLogHandler "Station Feed Failed, Export terminated"
                        Close #hmMsg
                        If ((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
                            ilRet = btrClose(hmTxr)
                            btrDestroy hmTxr
                        End If
                        On Error GoTo 0
                        cmcCancel.SetFocus
                        Exit Sub
                    End If
                    If (rbcInterface(0).Value) And (rbcGen(0).Value) Or (rbcInterface(1).Value) Then
                        mMakeCartStn tmStnInfo(ilLoop)
                        mMergeAddCyf
                        mMergeXRefCyf
                    End If
                End If
            Next ilLoop
            'If rbcFormat(0).Value Then
            '    gSwitchToPDF cdcSetup, 1
            'End If
            If (rbcInterface(0).Value) And (rbcGen(0).Value) Or (rbcInterface(1).Value) Then
                ilRet = mAddCyf()
                'Move carts into tmWemCyf so that PSA/Promo can be merged into it and tnXRefCyf is saved for mCreateCrossRef.
                'mCreateCrossRef can't be move prior to mCreateCartFile because it alters the structure of tmXRefCyf
                ReDim tmWemCyf(0 To UBound(tmXRefCyf)) As SENDCOPYINFO
                For ilWem = 0 To UBound(tmXRefCyf) Step 1
                    tmWemCyf(ilWem) = tmXRefCyf(ilWem)
                Next ilWem
                If (rbcInterface(0).Value) And (rbcGen(0).Value) Or (rbcInterface(1).Value) Then
                    mPSAPromoProcess
                End If
                ilRet = mCreateCartFile()
                ilRet = mCreateCrossRef()
                ilRet = mCreateEnvFile()
                ilRet = mCreateCartFile()
            End If
'            ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
'            mVehPop False
'            mRotPop
        Else
            Screen.MousePointer = vbDefault
            imExporting = False
            ''MsgBox "Station Feed already generated for this date, Export terminated", vbOkOnly + vbCritical + vbApplicationModal, "Export"
            gAutomationAlertAndLogHandler "Station Feed already generated for this date, Export terminated", vbOkOnly + vbCritical + vbApplicationModal, "Export"
            'Print #hmMsg, "Station Feed already generated for this date, Export terminated"
            gAutomationAlertAndLogHandler "Station Feed already generated for this date, Export terminated"
            Close #hmMsg
            If ((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
                ilRet = btrClose(hmTxr)
                btrDestroy hmTxr
            End If
            On Error GoTo 0
            cmcCancel.SetFocus
            Exit Sub
        End If
        If (rbcInterface(0).Value) Or ((rbcInterface(1).Value) And (ckcCmmlLog(0).Value = vbUnchecked) And (ckcCmmlLog(1).Value = vbUnchecked)) Then
            'Print #hmMsg, "** Export Station Feed Completed: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            gAutomationAlertAndLogHandler "** Export Station Feed Completed: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
        End If
        On Error GoTo 0
    End If
    If ((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
        ilRet = btrClose(hmTxr)
        btrDestroy hmTxr
    End If
    If ((rbcInterface(0).Value) And ((rbcGen(1).Value) Or (rbcGen(2).Value))) Or ((rbcInterface(1).Value) And ((ckcCmmlLog(0).Value = vbChecked) Or (ckcCmmlLog(1).Value = vbChecked))) Then
        ReDim imExpSpotVefCode(0 To 0) As Integer
        If rbcInterface(0).Value Then
            For ilLoop = 0 To lbcRegVeh.ListCount - 1 Step 1
                If lbcRegVeh.Selected(ilLoop) Then
                    imExpSpotVefCode(UBound(imExpSpotVefCode)) = imRegVefCode(ilLoop)
                    ReDim Preserve imExpSpotVefCode(0 To UBound(imExpSpotVefCode) + 1) As Integer
                End If
            Next ilLoop
        Else
            'Build vehicle array
            For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
                If lbcVeh.Selected(ilLoop) Then
                    slNameCode = tmVehCode(ilLoop).sKey    'Selling and conventional vehicles 'lbcVehCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                    ilRet = gBinarySearchVef(Val(slCode))
                    ilFound = False
                    If ilRet <> -1 Then
                        If tgMVef(ilRet).sType = "C" Then
                            For ilTest = 0 To UBound(imExpSpotVefCode) - 1 Step 1
                                If tgMVef(ilRet).iCode = imExpSpotVefCode(ilTest) Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilTest
                            If Not ilFound Then
                                imExpSpotVefCode(UBound(imExpSpotVefCode)) = tgMVef(ilRet).iCode
                                ReDim Preserve imExpSpotVefCode(0 To UBound(imExpSpotVefCode) + 1) As Integer
                            End If
                        ElseIf tgMVef(ilRet).sType = "S" Then
                            gBuildLinkArray hmVlf, tgMVef(ilRet), slStartDate, imAVefCode()
                            For ilLink = LBound(imAVefCode) To UBound(imAVefCode) - 1 Step 1
                                ilFound = False
                                For ilTest = 0 To UBound(imExpSpotVefCode) - 1 Step 1
                                    If imAVefCode(ilLink) = imExpSpotVefCode(ilTest) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilTest
                                If Not ilFound Then
                                    imExpSpotVefCode(UBound(imExpSpotVefCode)) = imAVefCode(ilLink)
                                    ReDim Preserve imExpSpotVefCode(0 To UBound(imExpSpotVefCode) + 1) As Integer
                                End If
                            Next ilLink
                        End If
                    End If
                End If
            Next ilLoop
        End If
        slDate = Format$(gNow(), "m/d/yy")
        gPackDate slDate, imGenDate(0), imGenDate(1)
        slTime = Format$(gNow(), "h:mm:ssAM/PM")
        gPackTime slTime, imGenTime(0), imGenTime(1)
        'Print #hmMsg, "** Storing Output into " & slToFile & " **"
        slMissingCopyNames = ""
'        For ilLoop = 0 To lbcRegVeh.ListCount - 1 Step 1
'            If lbcRegVeh.Selected(ilLoop) Then
        For ilLoop = 0 To UBound(imExpSpotVefCode) - 1 Step 1
                imCopyMissing = False
                ilVefCode = imExpSpotVefCode(ilLoop)    'imRegVefCode(ilLoop)
                ilRet = gBinarySearchVef(ilVefCode)
                If ilRet <> -1 Then
                    tmEVef = tgMVef(ilRet)
                End If
                If (ilRet <> -1) Then
                    smVehName = Trim$(tmEVef.sName)
                    slName = smVehName
                    'Print #hmMsg, "** Generating Data for " & Trim$(slName) & " **"
                    gAutomationAlertAndLogHandler "** Generating Data for " & Trim$(slName) & " **"
                    lacProcessing.Caption = "Generating Data for " & Trim$(slName)
                    If Not mExpSpots(0, "C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType()) Then
                        mClearCPR
                        'Print #hmMsg, "** Terminated **"
                        gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                        Close #hmMsg
                        'Close #hmTo
                        imExporting = False
                        'MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                        Screen.MousePointer = vbDefault
                        cmcCancel.SetFocus
                        Exit Sub
                    End If
                    If rbcInterface(1).Value Then
                        ilRet = mCreateCmmlLogFile(ilVefCode)
                        mClearCPR
                        If Not ilRet Then
                            'Print #hmMsg, "** Terminated **"
                            gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                            Close #hmMsg
                            'Close #hmTo
                            imExporting = False
                            'MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                            Screen.MousePointer = vbDefault
                            cmcCancel.SetFocus
                            Exit Sub
                        End If
                    End If
                    'Print #hmMsg, "** Completed " & Trim$(tmEVef.sName) & " **"
                    gAutomationAlertAndLogHandler "** Completed " & Trim$(tmEVef.sName) & " **"
                Else
                End If
'            End If
        Next ilLoop
        If rbcInterface(0).Value Then
            ilRet = mCreateSchFile()
        End If
        mClearCPR
        'Close #hmTo
        If rbcInterface(0).Value Then
            'Print #hmMsg, "** Completed Export Regional Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            gAutomationAlertAndLogHandler "** Completed Export Regional Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        Else
            'Print #hmMsg, "** Export Station Feed-KenCast Completed: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            gAutomationAlertAndLogHandler "** Export Station Feed-KenCast Completed: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        End If
        Close #hmMsg
        On Error GoTo 0
    End If
    If (rbcInterface(0).Value) And ((rbcGen(1).Value) Or (rbcGen(2).Value)) Then
    Else
        ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
        mVehPop False
        mRotPop
    End If
    If (rbcInterface(0).Value) And (rbcGen(1).Value) Then
        lacProcessing.Caption = "See: " & sgDBPath & "Messages\" & "ExpRgSpt.Txt" & " for Messages"
    ElseIf (rbcInterface(0).Value) And (rbcGen(2).Value) Then
        lacProcessing.Caption = "See: " & sgDBPath & "Messages\" & "ExpAlSpt.Txt" & " for Messages"
    ElseIf (rbcInterface(0).Value) And (rbcGen(3).Value) Then
        lacProcessing.Caption = "See: " & sgDBPath & "Messages\" & "ExpInst.Txt" & " for Messages"
    Else
        lacProcessing.Caption = "See: " & sgDBPath & "Messages\" & "ExpStnFd.Txt" & " for Messages"
    End If
    Screen.MousePointer = vbDefault
    imExporting = False
    cmcExport.Enabled = False
    cmcCancel.Caption = "Done"
    cmcCancel.SetFocus
    Exit Sub

ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)

End Sub
Private Sub cmcExport_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcExport_LostFocus()
    lacProcessing.Caption = ""
End Sub
Private Sub cmcFrom_Click()
    igBrowserType = 1
    Browser.Show vbModal
    plcDates.Visible = False
    plcDates.Visible = True
    If rbcInterface(0).Value Then
        plcSelect(0).Visible = False
        plcSelect(0).Visible = True
    Else
        plcSelect(1).Visible = False
        plcSelect(1).Visible = True
    End If
    If igBrowserReturn = 1 Then
        edcFrom.Text = sgBrowserFile
    End If
    DoEvents
    edcFrom.SetFocus
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
End Sub

Private Sub cmcGetRot_Click()
    tmcRot_Timer
End Sub

Private Sub cmcGetRot_GotFocus()
    plcCalendar.Visible = False
End Sub


Private Sub cmcReSend_Click()
    If rbcInterface(0).Value Then
        igSGOrKC = 0
    Else
        igSGOrKC = 1
    End If
    StnFdUnd.Show vbModal
    plcDates.Visible = False
    plcDates.Visible = True
    If rbcInterface(0).Value Then
        plcSelect(0).Visible = False
        plcSelect(0).Visible = True
    Else
        plcSelect(1).Visible = False
        plcSelect(1).Visible = True
    End If
    If igBFReturn = 1 Then
        Screen.MousePointer = vbHourglass
        ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
        mVehPop False
        mRotPop
        Screen.MousePointer = vbDefault
        mSetCommands
    End If
End Sub
Private Sub cmcReSend_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcStartDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcStartDate_GotFocus()
    Dim slStr As String
    'If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
    imFirstFocus = False
    '    'Show branner
    '    mInitDDE
    '    mSendHelpMsg "BT"
    'End If
    If imDateBox <> 1 Then
    plcCalendar.Visible = False
    slStr = edcStartDate.Text
    If gValidDate(slStr) Then
        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
        pbcCalendar_Paint   'mBoxCalDate called within paint
    Else
        lacDate.Visible = False
    End If
    End If
    imDateBox = 1
    plcCalendar.Move plcDates.Left + edcStartDate.Left, plcDates.Top + edcStartDate.Top + edcStartDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcSuppress_Click()
    Dim ilCrf As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer

    If lbcRot.ListCount <= 0 Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    For ilCrf = 0 To UBound(tgSortCrf) - 1 Step 1
        If tgSortCrf(ilCrf).iSelected Then
            ilIndex = tgSortCrf(ilCrf).iCombineIndex
            Do While ilIndex >= 0
                Do
                    ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgCombineCrf(ilIndex).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilRet = BTRV_ERR_NONE Then
                        'tmRec = tmCrf
                        'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                        'tmCrf = tmRec
                        If rbcInterface(0).Value Then
                            tmCrf.sAffFdStatus = "P" '"S"
                        Else
                            tmCrf.sKCFdStatus = "P"
                        End If
                        ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilIndex = tgCombineCrf(ilIndex).iCombineIndex
            Loop
            ilIndex = tgSortCrf(ilCrf).iDuplIndex
            Do While ilIndex >= 0
                Do
                    ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgDuplCrf(ilIndex).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilRet = BTRV_ERR_NONE Then
                        'tmRec = tmCrf
                        'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                        'tmCrf = tmRec
                        If rbcInterface(0).Value Then
                            tmCrf.sAffFdStatus = "P" '"S"
                        Else
                            tmCrf.sKCFdStatus = "P"
                        End If
                        ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilIndex = tgDuplCrf(ilIndex).iDuplIndex
            Loop
            Do
                ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgSortCrf(ilCrf).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet = BTRV_ERR_NONE Then
                    'tmRec = tmCrf
                    'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                    'tmCrf = tmRec
                    If rbcInterface(0).Value Then
                        tmCrf.sAffFdStatus = "P" '"S"
                    Else
                        tmCrf.sKCFdStatus = "P"
                    End If
                    ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
    Next ilCrf
    ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
    mVehPop False
    mRotPop
    Screen.MousePointer = vbDefault
    mSetCommands
End Sub
Private Sub cmcSuppress_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcTranDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcTranDate.SelStart = 0
    edcTranDate.SelLength = Len(edcTranDate.Text)
    edcTranDate.SetFocus
End Sub
Private Sub cmcTranDate_GotFocus()
    Dim slStr As String
    'If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
    '    'Show branner
    '    mInitDDE
    '    mSendHelpMsg "BT"
    'End If
    If imDateBox <> 3 Then
        plcCalendar.Visible = False
        slStr = edcTranDate.Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    imDateBox = 3
    plcCalendar.Move plcDates.Left + edcTranDate.Left, plcDates.Top + edcTranDate.Top + edcTranDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcEndDate_Change()
    Dim slStr As String
    Dim ilLoop As Integer
    slStr = edcEndDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    ReDim tgSortCrf(0 To 0) As SORTCRF
    lbcRot.Clear
    pbcLbcRot_Paint
    ReDim tgDuplCrf(0 To 0) As DUPLCRF
    ReDim tgCombineCrf(0 To 0) As COMBINECRF
    ckcAll.Value = vbUnchecked
    tmcRot.Enabled = False
    slStr = edcStartDate.Text
    If (gValidDate(slStr)) And (((rbcInterface(0).Value) And (rbcGen(0).Value)) Or (rbcInterface(1).Value)) Then
        For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
            If lbcVeh.Selected(ilLoop) Then
                'tmcRot.Enabled = False
                'tmcRot.Enabled = True
                Exit For
            End If
        Next ilLoop
    End If
    mSetCommands
End Sub
Private Sub edcEndDate_GotFocus()
    Dim slStr As String
    'If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
    imFirstFocus = False
    'Show branner
    '    mInitDDE
    '    mSendHelpMsg "BT"
    'End If
    If imDateBox <> 2 Then
        plcCalendar.Visible = False
    End If
    imDateBox = 2
    If edcEndDate.Text = "" Then
        slStr = edcStartDate.Text
        If slStr <> "" Then
            'If rbcGen(1).Value Then
            If (rbcInterface(0).Value) And ((rbcGen(1).Value) Or (rbcGen(2).Value)) Then
                edcEndDate.Text = slStr
            Else
                edcEndDate.Text = gObtainNextSunday(slStr)
            End If
        End If
    End If
    plcCalendar.Move plcDates.Left + edcEndDate.Left, plcDates.Top + edcEndDate.Top + edcEndDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcEndDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcEndDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
End Sub
Private Sub edcFrom_Change()
    Dim ilRet As Integer
    lbcRegVeh.Clear
    'If rbcGen(1).Value Then
    If (rbcInterface(0).Value) And ((rbcGen(1).Value) Or (rbcGen(2).Value)) Then
        Screen.MousePointer = vbHourglass
        smStationFile = Trim$(edcFrom.Text)
        ilRet = mGetStnInfo(False)
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcRunLetter_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcRunLetter_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub edcStartDate_Change()
    Dim slStr As String
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    edcEndDate.Text = ""
    ReDim tgSortCrf(0 To 0) As SORTCRF
    lbcRot.Clear
    pbcLbcRot_Paint
    ReDim tgDuplCrf(0 To 0) As DUPLCRF
    ReDim tgCombineCrf(0 To 0) As COMBINECRF
    ckcAll.Value = vbUnchecked
    'For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
    '    lbcVeh.Selected(ilLoop) = False
    '    tmcRot.Enabled = False
    'Next ilLoop
End Sub
Private Sub edcStartDate_GotFocus()
    'If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
    imFirstFocus = False
    'Show branner
    '    mInitDDE
    '    mSendHelpMsg "BT"
    'End If
    If imDateBox <> 1 Then
        plcCalendar.Visible = False
    End If
    imDateBox = 1
    plcCalendar.Move plcDates.Left + edcStartDate.Left, plcDates.Top + edcStartDate.Top + edcStartDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcStartDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcStartDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
End Sub
Private Sub edcTranDate_Change()
    Dim slStr As String
    slStr = edcTranDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    mSetCommands
End Sub
Private Sub edcTranDate_GotFocus()
    'If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    '    mInitDDE
    '    mSendHelpMsg "BT"
    'End If
    If imDateBox <> 3 Then
        plcCalendar.Visible = False
    End If
    imDateBox = 3
    plcCalendar.Move plcDates.Left + edcTranDate.Left, plcDates.Top + edcTranDate.Top + edcTranDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcTranDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTranDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcTranDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcTranDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcTranDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcTranDate.Text = slDate
            End If
        End If
        edcTranDate.SelStart = 0
        edcTranDate.SelLength = Len(edcTranDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcTranDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcTranDate.Text = slDate
            End If
        End If
        edcTranDate.SelStart = 0
        edcTranDate.SelLength = Len(edcTranDate.Text)
    End If
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False

'    Me.Visible = False
'    DoEvents    'Process events so pending keys are not sent to this
'    Me.Visible = True
'    If lbcVeh.Visible Then
'        lbcVeh.Visible = False
'        lbcVeh.Visible = True
'    End If
'    Me.KeyPreview = True
'    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
        plcDates.Visible = False
        plcDates.Visible = True
        If rbcInterface(0).Value Then
            plcSelect(0).Visible = False
            plcSelect(0).Visible = True
        Else
            plcSelect(1).Visible = False
            plcSelect(1).Visible = True
        End If
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase smRotComment
    Erase smTimeRestrictions
    Erase smDayRestrictions
    '6/14/06
    Erase tmDuplComment

    Erase imExpSpotVefCode
    Erase imAVefCode
    Erase tgAddCyf
    Erase tgCartStnXRef
    Erase tmSTCrf
    Erase tgSortCrf
    Erase tmSvSortCrf
    Erase lmReadyCRF
    Erase imNoTimesMod
    Erase imSelectPrevState
    Erase lmRotCodeBuild
    Erase tmCnfRot
    Erase tmCyfTest
    Erase tmVef
    Erase tmVpfInfo
    Erase tmLkVehInfo
    Erase tmSALink
    Erase tmVehTimes
    Erase tmRotInfo
    Erase tmAddCyf
    Erase tmXRefCyf
    Erase tgCombineCrf
    Erase tgDuplCrf
    Erase smFileNames
    Erase imVefCode
    Erase imDVefCode
    Erase tmStnInfo
    Erase tgSchSpotInfo
    Erase imRegVefCode
    Erase tmPSAPromoSortCrf
    Erase tmWemCyf
    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    ilRet = btrClose(hmDlf)
    btrDestroy hmDlf
    ilRet = btrClose(hmTzf)
    btrDestroy hmTzf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmCTSsf)
    btrDestroy hmCTSsf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmRsf)
    btrDestroy hmRsf
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    'ilRet = btrClose(hmVlf)
    'btrDestroy hmVlf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    ilRet = btrClose(hmCsf)
    btrDestroy hmCsf
    ilRet = btrClose(hmCpr)
    btrDestroy hmCpr
    ilRet = btrClose(hmBof)
    btrDestroy hmBof
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmCyf)
    btrDestroy hmCyf
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCnf)
    btrDestroy hmCnf
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    igJobShowing(STATIONFEEDJOB) = False
    Set ExpStnFd = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcRegVeh_Click()
    mSetCommands
End Sub
Private Sub lbcRegVeh_GotFocus()
    plcCalendar.Visible = False
    tmcRot.Enabled = False
End Sub
Private Sub lbcRot_Click()
    Dim ilStartIndex As Integer
    Dim ilEndIndex As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim llRg As Long
    Dim llRet As Long
    Dim ilListIndex As Integer
    Dim ilResetLast As Integer
    If imIgnoreVbcChg Then
        Exit Sub
    End If
    imIgnoreVbcChg = True
    Screen.MousePointer = vbHourglass
    If ckcAll.Value = vbChecked Then
        imIgnoreCkcAll = True
        ckcAll.Value = vbUnchecked
        DoEvents
        imIgnoreCkcAll = False
    End If
    ilResetLast = True
    ReDim imSelectPrevState(LBound(tgSortCrf) To UBound(tgSortCrf)) As Integer
    For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
        imSelectPrevState(ilLoop) = tgSortCrf(ilLoop).iSelected
    Next ilLoop
    ilListIndex = imCurrentIndex + vbcRot.Value
    ilStartIndex = vbcRot.Value
    ilEndIndex = ilStartIndex + vbcRot.LargeChange
    If ilEndIndex > UBound(tgSortCrf) - 1 Then
        ilEndIndex = UBound(tgSortCrf) - 1
    End If
    If ((imShiftKey And 1) = 1) And (imLastIndex >= 0) Then
        For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
            tgSortCrf(ilLoop).iSelected = False
        Next ilLoop
        If imLastIndex <= ilListIndex Then
            For ilLoop = imLastIndex To ilListIndex Step 1
                tgSortCrf(ilLoop).iSelected = True
            Next ilLoop
        Else
            For ilLoop = ilListIndex To imLastIndex Step 1
                tgSortCrf(ilLoop).iSelected = True
            Next ilLoop
        End If
        ilValue = False
        If UBound(tgSortCrf) < vbcRot.LargeChange + 1 Then
            llRg = CLng(UBound(tgSortCrf) - 1) * &H10000 Or 0
        Else
            llRg = CLng(vbcRot.LargeChange) * &H10000 Or 0
        End If
        llRet = SendMessageByNum(lbcRot.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        ilIndex = 0
        For ilLoop = ilStartIndex To ilEndIndex Step 1
            lbcRot.Selected(ilIndex) = tgSortCrf(ilLoop).iSelected
            ilIndex = ilIndex + 1
        Next ilLoop
        ilResetLast = False
    ElseIf ((imShiftKey And 2) = 2) Then    'Ctrl
        ilIndex = 0
        For ilLoop = ilStartIndex To ilEndIndex Step 1
            tgSortCrf(ilLoop).iSelected = lbcRot.Selected(ilIndex)
            ilIndex = ilIndex + 1
        Next ilLoop
    Else
        For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
            tgSortCrf(ilLoop).iSelected = False
        Next ilLoop
        ilIndex = 0
        For ilLoop = ilStartIndex To ilEndIndex Step 1
            tgSortCrf(ilLoop).iSelected = lbcRot.Selected(ilIndex)
            ilIndex = ilIndex + 1
        Next ilLoop
    End If
    DoEvents
    If ilResetLast Then
        imLastIndex = ilListIndex
    End If
    For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
        If imSelectPrevState(ilLoop) <> tgSortCrf(ilLoop).iSelected Then
            If imSelectPrevState(ilLoop) Then
                tmCrf = tgSortCrf(ilLoop).tCrf
                mComputeTime False
                ilIndex = tgSortCrf(ilLoop).iCombineIndex
                Do While ilIndex >= 0
                    tmCrf = tgCombineCrf(ilIndex).tCrf
                    mComputeTime False
                    ilIndex = tgCombineCrf(ilIndex).iCombineIndex
                Loop
                ilIndex = tgSortCrf(ilLoop).iDuplIndex
                Do While ilIndex >= 0
                    tmCrf = tgDuplCrf(ilIndex).tCrf
                    mComputeTime False
                    ilIndex = tgDuplCrf(ilIndex).iDuplIndex
                Loop
            Else
                tmCrf = tgSortCrf(ilLoop).tCrf
                mComputeTime True
                ilIndex = tgSortCrf(ilLoop).iCombineIndex
                Do While ilIndex >= 0
                    tmCrf = tgCombineCrf(ilIndex).tCrf
                    mComputeTime True
                    ilIndex = tgCombineCrf(ilIndex).iCombineIndex
                Loop
                ilIndex = tgSortCrf(ilLoop).iDuplIndex
                Do While ilIndex >= 0
                    tmCrf = tgDuplCrf(ilIndex).tCrf
                    mComputeTime True
                    ilIndex = tgDuplCrf(ilIndex).iDuplIndex
                Loop
            End If
            DoEvents
        End If
    Next ilLoop
    mSetCommands
    pbcLbcRot_Paint
    pbclbcVehicle_Paint
    Screen.MousePointer = vbDefault
    imIgnoreVbcChg = False
End Sub
Private Sub lbcRot_GotFocus()
    plcCalendar.Visible = False
    'If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    '    mInitDDE
    '    mSendHelpMsg "BT"
    'End If
End Sub
Private Sub lbcRot_KeyDown(KeyCode As Integer, Shift As Integer)
    imShiftKey = Shift
End Sub
Private Sub lbcRot_KeyUp(KeyCode As Integer, Shift As Integer)
    imShiftKey = Shift
End Sub
Private Sub lbcRot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imCurrentIndex = Y \ fgListHtArial825
    imButton = Button
    If Button = 2 Then  'Right Mouse
        imButtonIndex = imCurrentIndex + vbcRot.Value
        If (imButtonIndex >= 0) And (imButtonIndex <= UBound(tgSortCrf) - 1) Then
            imIgnoreRightMove = True
            mShowRotInfo
            imIgnoreRightMove = False
        End If
    End If
End Sub
Private Sub lbcRot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        If (Y < 0) Or (Y > lbcRot.Height) Then
            imButtonIndex = 0
            plcRotInfo.Visible = False
            Exit Sub
        End If
        If (X < 0) Or (X > lbcRot.Width) Then
            imButtonIndex = 0
            plcRotInfo.Visible = False
            Exit Sub
        End If
        If imButtonIndex <> (Y \ fgListHtArial825) + vbcRot.Value Then
            imIgnoreRightMove = True
            imButtonIndex = Y \ fgListHtArial825 + vbcRot.Value
            If (imButtonIndex >= 0) And (imButtonIndex <= UBound(tgSortCrf) - 1) Then
                mShowRotInfo
            Else
                plcRotInfo.Visible = False
            End If
            imIgnoreRightMove = False
        End If
    End If
End Sub
Private Sub lbcRot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        plcRotInfo.Visible = False
    End If
End Sub

Private Sub lbcRot_Scroll()
    pbcLbcRot_Paint
End Sub

Private Sub lbcVeh_Click()
    'tmcRot.Enabled = False
    'tmcRot.Enabled = True
    mSetCommands
End Sub
Private Sub lbcVeh_GotFocus()
    plcCalendar.Visible = False
    tmcRot.Enabled = False
End Sub

Private Sub lbcVeh_Scroll()
    If tmcRot.Enabled Then
        'tmcRot.Enabled = False
        'tmcRot.Enabled = True
    End If
End Sub

Private Sub lbcVehicle_Click()
    'ignore any selections
    'pbcLbcVehicle_Paint
End Sub

Private Sub lbcVehicle_GotFocus()
    plcCalendar.Visible = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAbortTrans                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Abort Transaction and remove   *
'*                      files created                  *
'*                                                     *
'*******************************************************
Private Sub mAbortTrans()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    ilRet = btrAbortTrans(hmCrf)
    On Error GoTo 0
    For ilLoop = 0 To UBound(smFileNames) - 1 Step 1
        Kill smFileNames(ilLoop)
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddCyf                         *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Update or Add the CYF records  *
'*                                                     *
'*******************************************************
Private Function mAddCyf() As Integer
    Dim ilLoop As Integer
    Dim ilLoop1 As Integer
    Dim ilRet As Integer
    Dim ilVIndex As Integer
    Dim tlCyf As CYF
    lacProcessing.Caption = "Updating Copy Inventory"
    mAddCyf = False
    DoEvents
    For ilLoop = 0 To UBound(tgAddCyf) - 1 Step 1
        'If transmitted- remove old record, then insert instead of updating
        Do
            tmCyfSrchKey.lCifCode = tgAddCyf(ilLoop).tCyf.lCifCode
            tmCyfSrchKey.iVefCode = tgAddCyf(ilLoop).tCyf.iVefCode
            If rbcInterface(0).Value Then
                tmCyfSrchKey.sSource = "S"
            Else
                tmCyfSrchKey.sSource = "K"
            End If
            If Trim$(tgAddCyf(ilLoop).tCyf.sTimeZone) <> "R" Then
                tmCyfSrchKey.sTimeZone = tgAddCyf(ilLoop).tCyf.sTimeZone
                tmCyfSrchKey.lRafCode = 0
            Else
                tmCyfSrchKey.sTimeZone = tgAddCyf(ilLoop).tCyf.sTimeZone
                tmCyfSrchKey.lRafCode = tgAddCyf(ilLoop).tCyf.lRafCode
            End If
            ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                If rbcInterface(0).Value Then
                    tgAddCyf(ilLoop).tCyf.sAffOrigXMitChar = tmCyf.sAffOrigXMitChar
                    tgAddCyf(ilLoop).tCyf.iAffOrigXMitDate(0) = tmCyf.iAffOrigXMitDate(0)
                    tgAddCyf(ilLoop).tCyf.iAffOrigXMitDate(1) = tmCyf.iAffOrigXMitDate(1)
                Else
                    tgAddCyf(ilLoop).tCyf.sKCOrigXMitChar = tmCyf.sKCOrigXMitChar
                    tgAddCyf(ilLoop).tCyf.iKCOrigXMitDate(0) = tmCyf.iKCOrigXMitDate(0)
                    tgAddCyf(ilLoop).tCyf.iKCOrigXMitDate(1) = tmCyf.iKCOrigXMitDate(1)
                End If
                ilRet = btrDelete(hmCyf)
            Else
                ilRet = BTRV_ERR_NONE
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            'Print #hmMsg, "Delete CYF Failed" & str$(ilRet) & " processing terminated"
            gAutomationAlertAndLogHandler "Delete CYF Failed" & str$(ilRet) & " processing terminated"
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("File in Use [Re-press Export], Delete Cyf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
            Exit Function
        End If
        tgAddCyf(ilLoop).tCyf.lCode = 0
        ilRet = btrInsert(hmCyf, tgAddCyf(ilLoop).tCyf, imCyfRecLen, INDEXKEY1)
        If ilRet <> BTRV_ERR_NONE Then
            'Print #hmMsg, "Insert CYF Failed" & str$(ilRet) & " processing terminated"
            gAutomationAlertAndLogHandler "Insert CYF Failed" & str$(ilRet) & " processing terminated"
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("File in Use [Re-press Export], Insert Cyf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
            Exit Function
        End If
        'Test if airing and group vehicles defined- if so insert for other
        'vehicles
        For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
            If tmVef(ilLoop1).iCode = tgAddCyf(ilLoop).tCyf.iVefCode Then
                If (tmVef(ilLoop1).sType = "A") Or (tmVef(ilLoop1).sType = "C") Then
                    'Update cyf for all vehicles
                    ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                    'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                        Do While ilVIndex >= 0
                            Do
                                tmCyfSrchKey.lCifCode = tgAddCyf(ilLoop).tCyf.lCifCode
                                tmCyfSrchKey.iVefCode = tmLkVehInfo(ilVIndex).iVefCode 'tmVpfInfo(ilVIndex).iVefLink(ilLoop2)
                                If rbcInterface(0).Value Then
                                    tmCyfSrchKey.sSource = "S"
                                Else
                                    tmCyfSrchKey.sSource = "K"
                                End If
                                If Trim$(tgAddCyf(ilLoop).tCyf.sTimeZone) <> "R" Then
                                    tmCyfSrchKey.sTimeZone = tgAddCyf(ilLoop).tCyf.sTimeZone
                                    tmCyfSrchKey.lRafCode = 0
                                Else
                                    tmCyfSrchKey.sTimeZone = tgAddCyf(ilLoop).tCyf.sTimeZone
                                    tmCyfSrchKey.lRafCode = tgAddCyf(ilLoop).tCyf.lRafCode
                                End If
                                ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                                    ilRet = btrDelete(hmCyf)
                                Else
                                    ilRet = BTRV_ERR_NONE
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If ilRet <> BTRV_ERR_NONE Then
                                'Print #hmMsg, "Delete Linked CYF Failed" & str$(ilRet) & " processing terminated"
                                gAutomationAlertAndLogHandler "Delete Linked CYF Failed" & str$(ilRet) & " processing terminated"
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("File in Use [Re-press Export], Delete Cyf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
                                Exit Function
                            End If
                            tlCyf = tgAddCyf(ilLoop).tCyf
                            tlCyf.iVefCode = tmLkVehInfo(ilVIndex).iVefCode 'tmVpfInfo(ilVIndex).iVefLink(ilLoop2)
                            tlCyf.lCode = 0
                            ilRet = btrInsert(hmCyf, tlCyf, imCyfRecLen, INDEXKEY1)
                            If ilRet <> BTRV_ERR_NONE Then
                                'Print #hmMsg, "Insert Linked CYF Failed" & str$(ilRet) & " processing terminated"
                                gAutomationAlertAndLogHandler "Insert Linked CYF Failed" & str$(ilRet) & " processing terminated"
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("File in Use [Re-press Export], Insert Cyf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
                                Exit Function
                            End If
                        'Next ilLoop2
                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                End If
            End If
        Next ilLoop1
    Next ilLoop
    mAddCyf = True
End Function
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
    If imDateBox = 1 Then
        slStr = edcStartDate.Text
    ElseIf imDateBox = 2 Then
        slStr = edcEndDate.Text
    ElseIf imDateBox = 3 Then
        slStr = edcTranDate.Text
    End If
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
'*      Procedure Name:mBuildExpTable                  *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build table from the selected   *
'*                     rotations for exporting         *
'*                                                     *
'*******************************************************
Private Sub mBuildExpTable(tlStnInfo As STNINFO)
'   mBuildExpTable
'
    Dim ilCrf As Integer
    Dim ilVeh As Integer
    Dim ilVehIndex As Integer
    Dim ilVpfIndex As Integer
    Dim slKey As String
    Dim ilVpf As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim llRotStartDate As Long
    Dim llRotEndDate As Long
    Dim llSTStartDate As Long
    Dim llSTEndDate As Long
    Dim llRotStartTime As Long
    Dim llRotEndTime As Long
    Dim llSTStartTime As Long
    Dim llSTEndTime As Long
    Dim ilDay As Integer
    Dim slFeedDate As String
    Dim llTranDate As Long
    Dim ilTranDate0 As Integer
    Dim ilTranDate1 As Integer
    Dim slProduct As String
    Dim ilName As Integer
    Dim ilFound As Integer
    Dim slStr As String
    Dim ilTransmit As Integer
    Dim ilVIndex As Integer
    Dim slCart As String
    Dim ilCombineIndex As Integer
    Dim ilDuplIndex As Integer
    Dim ilDone As Integer
    Dim ilSend As Integer
    Dim ilRot As Integer
    Dim ilVefCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFdDateNew As Integer
    Dim llRafCode As Long
    Dim tlCrf As CRF
    ReDim tmRotInfo(0 To 0) As SENDROTINFO
    ReDim tmAddCyf(0 To 0) As SENDCOPYINFO
    ReDim lmRotCodeBuild(0 To 0) As Long
    slDate = edcTranDate.Text
    llTranDate = gDateValue(slDate)
    gPackDate slDate, ilTranDate0, ilTranDate1
    For ilCrf = 0 To UBound(tgSortCrf) - 1 Step 1
        If tgSortCrf(ilCrf).iSelected Then
            tmCrf = tgSortCrf(ilCrf).tCrf
            ilSend = True
        Else
            ilSend = False
        End If
        If ilSend Then
            'Test if for vehicle
            ilSend = False
            For ilVeh = 0 To UBound(tmVef) - 1 Step 1
                If tmVef(ilVeh).iCode = tlStnInfo.iAirVeh Then
                    ilSend = True
                    ilFound = ilVeh
                    Exit For
                Else
                    ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                        Do While ilVIndex >= 0
                            If tlStnInfo.iAirVeh = tmLkVehInfo(ilVIndex).iVefCode Then
                                ilSend = True
                                ilFound = ilVeh
                                Exit For
                            End If
                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                End If
            Next ilVeh
        End If
        If ilSend Then
            'Test if for vehicle
            ilSend = False
            ilVeh = ilFound
            For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
                If lbcVeh.Selected(ilLoop) Then
                    slNameCode = tmVehCode(ilLoop).sKey    'Selling and conventional vehicles 'lbcVehCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                    If tmCrf.iVefCode = Val(slCode) Then
                        If tmVef(ilVeh).sType = "C" Then
                            If tlStnInfo.iAirVeh = Val(slCode) Then
                                ilSend = True
                                Exit For
                            End If
                            ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                            If ilVIndex >= 0 Then
                                ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                                Do While ilVIndex >= 0
                                    If tmLkVehInfo(ilVIndex).iVefCode = tlStnInfo.iAirVeh Then
                                        ilSend = True
                                        Exit Do
                                    End If
                                    ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                                Loop
                            End If
                        Else
                            ilVIndex = mFindVpfIndex(Val(slCode)) 'tmVef(ilVehIndex).iCode)
                            If ilVIndex >= 0 Then
                                ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                                Do While ilVIndex >= 0
                                    If tlStnInfo.iAirVeh = tmSALink(ilVIndex).iVefCode Then
                                        ilSend = True
                                        Exit Do
                                    End If
                                    ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                                Loop
                            End If
                        End If
                        Exit For
                    Else
                        ilVIndex = mFindVpfIndex(Val(slCode)) 'tmVef(ilVehIndex).iCode)
                        If ilVIndex >= 0 Then
                            ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                            Do While ilVIndex >= 0
                                If tmCrf.iVefCode = tmSALink(ilVIndex).iVefCode Then
                                    ilSend = True
                                    Exit Do
                                End If
                                ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                            Loop
                            If ilSend Then
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next ilLoop
        End If
        If ilSend Then
            'If tlStnInfo.lRafCode > 0 Then
            '    If Trim$(tmCrf.sZone) <> "R" Then
            '        ilSend = False
            '    Else
            '        If tlStnInfo.lRafCode <> tmCrf.lRafCode Then
            '            ilSend = False
            '            ilVIndex = tlStnInfo.iLkStnInfo
            '            Do While ilVIndex >= 0
            '                If tmStnInfo(ilVIndex).lRafCode = tmCrf.lRafCode Then
            '                    ilSend = True
            '                    Exit Do
            '                End If
            '                ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
            '            Loop
            '        End If
            '    End If
            'Else
            '    If Trim$(tmCrf.sZone) = "R" Then
            '        ilSend = False
            '    Else
            '        'Test if rotation exist that has region superseding this rotation
            '        If mRegionExist(tgSortCrf(ilCrf)) Then
            '            ilSend = False
            '        End If
            '    End If
            'End If
            If Trim$(tmCrf.sZone) = "R" Then
                'Is station in the region
                If tlStnInfo.lRafCode <> tmCrf.lRafCode Then
                    ilSend = False
                    ilVIndex = tlStnInfo.iLkStnInfo
                    Do While ilVIndex >= 0
                        If tmStnInfo(ilVIndex).lRafCode = tmCrf.lRafCode Then
                            ilSend = True
                            Exit Do
                        End If
                        ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
                    Loop
                End If
            Else
                'General sent to stations not in region
                llRafCode = tlStnInfo.lRafCode
                ilVIndex = tlStnInfo.iLkStnInfo
'                Do
'                    If llRafCode > 0 Then
'                        If mRegionExist(tgSortCrf(ilCrf), llRafCode) Then
'                            ilSend = False
'                            Exit Do
'                        End If
'                    End If
'                    If ilVIndex >= 0 Then
'                        ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
'                        If ilVIndex >= 0 Then
'                            llRafCode = tmStnInfo(ilVIndex).lRafCode
'                        End If
'                    End If
'                Loop While ilVIndex >= 0
                Do
                    If llRafCode > 0 Then
                        If mRegionExist(tgSortCrf(ilCrf), llRafCode) Then
                            ilSend = False
                            Exit Do
                        End If
                    End If
                    If ilVIndex >= 0 Then
                        llRafCode = tmStnInfo(ilVIndex).lRafCode
                        ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If ilSend Then
            tmCrf = tgSortCrf(ilCrf).tCrf
            gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
            llRotStartDate = gDateValue(slDate)
            gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
            llRotEndDate = gDateValue(slDate)
            gUnpackTimeLong tmCrf.iStartTime(0), tmCrf.iStartTime(1), False, llRotStartTime
            gUnpackTimeLong tmCrf.iEndTime(0), tmCrf.iEndTime(1), True, llRotEndTime
            ilCombineIndex = -1
            ilDuplIndex = -1
            ilDone = False
            ReDim ilVehSent(0 To 0) As Integer
            Do
                ilSend = True
                For ilRot = 0 To UBound(lmRotCodeBuild) - 1 Step 1
                    If tmCrf.lCode = lmRotCodeBuild(ilRot) Then
                        ilSend = False
                    End If
                Next ilRot
                If ilSend Then
                    lmRotCodeBuild(UBound(lmRotCodeBuild)) = tmCrf.lCode
                    ReDim Preserve lmRotCodeBuild(0 To UBound(lmRotCodeBuild) + 1) As Long
                    For ilVeh = 0 To UBound(tmVef) - 1 Step 1
                        If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
                            ilVehIndex = ilVeh
                            Exit For
                        End If
                    Next ilVeh
                    ilVpfIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
                    'Loop thru all airing vehicles associated with selling or conventional vehicle
                    'If tmVef(ilVehIndex).sType <> "S" Then
                    '    For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
                    '        tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) = 0
                    '    Next ilVeh
                    '    tmVpfInfo(ilVpfIndex).tVpf.iGLink(LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink)) = tmVef(ilVehIndex).iCode
                    'End If
                    'For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
                    '    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) <> 0 Then
                    ReDim imVefCode(0 To 0) As Integer
                    ilVefCode = tmVef(ilVehIndex).iCode
                    If tmVef(ilVehIndex).sType = "S" Then
                        'Find airing vehicle
                        If ilVpfIndex >= 0 Then
                            ilVpf = tmVpfInfo(ilVpfIndex).iFirstSALink
                            Do While ilVpf >= 0
                                imVefCode(UBound(imVefCode)) = tmSALink(ilVpf).iVefCode
                                ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
                                ilVpf = tmSALink(ilVpf).iNextLkVehInfo
                            Loop
                        End If
                    Else
                        ReDim imVefCode(0 To 1) As Integer
                        imVefCode(0) = ilVefCode
                    End If
                    For ilVeh = 0 To UBound(imVefCode) - 1 Step 1
                        ilVefCode = imVefCode(ilVeh)
                            ilSend = True
                            For ilRot = 0 To UBound(ilVehSent) - 1 Step 1
                                If ilVehSent(ilRot) = ilVefCode Then    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                    ilSend = False
                                    Exit For
                                End If
                            Next ilRot
                            If ilSend Then
                                ilVehSent(UBound(ilVehSent)) = ilVefCode    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                                ReDim Preserve ilVehSent(0 To UBound(ilVehSent) + 1) As Integer
                                ReDim Preserve tmRotInfo(0 To UBound(tmRotInfo) + 1) As SENDROTINFO
                                For ilName = 0 To UBound(tmVef) - 1 Step 1
                                    If tmVef(ilName).iCode = ilVefCode Then 'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                        slKey = tmVef(ilName).sName
                                        ilVpf = mFindVpfIndex(ilVefCode)    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh))
                                        'If ilVpf > 0 Then
                                        '    For ilLoop1 = LBound(tmVpfInfo(ilVpf).sVefName) To tmVpfInfo(ilVpf).iNoVefLinks - 1 Step 1
                                        '        slKey = Trim$(slKey) & " " & tmVpfInfo(ilVpf).sVefName(ilLoop1)
                                        '    Next ilLoop1
                                        'End If
                                        If ilVpf >= 0 Then
                                            ilVpf = tmVpfInfo(ilVpf).iFirstLkVehInfo
                                            Do While ilVpf >= 0
                                                slKey = Trim$(slKey) & " " & tmLkVehInfo(ilVpf).sVefName
                                                ilVpf = tmLkVehInfo(ilVpf).iNextLkVehInfo
                                            Loop
                                        End If
                                        Exit For
                                    End If
                                Next ilName
                                slKey = slKey & "|" & tgSortCrf(ilCrf).sCntrProd
                                slStr = Trim$(str$(llRotStartDate))
                                Do While Len(slStr) < 6
                                    slStr = "0" & slStr
                                Loop
                                slKey = slKey & "|" & slStr
                                For ilDay = 0 To 6 Step 1
                                    If tmCrf.sDay(ilDay) <> "N" Then
                                        slStr = Trim$(str$(ilDay))
                                        Exit For
                                    End If
                                Next ilDay
                                slKey = slKey & "|" & slStr
                                slStr = Trim$(str$(llRotStartTime))
                                Do While Len(slStr) < 6
                                    slStr = "0" & slStr
                                Loop
                                slKey = slKey & "|" & slStr
                                Select Case tmCrf.sZone
                                    Case "EST"
                                        slKey = slKey & "|1"
                                    Case "CST"
                                        slKey = slKey & "|2"
                                    Case "MST"
                                        slKey = slKey & "|3"
                                    Case "PST"
                                        slKey = slKey & "|4"
                                    Case Else
                                        slKey = slKey & "|5"
                                End Select
                                tmRotInfo(UBound(tmRotInfo) - 1).sKey = slKey
                                tmRotInfo(UBound(tmRotInfo) - 1).lCrfCode = tmCrf.lCode
                                tmRotInfo(UBound(tmRotInfo) - 1).iVefCode = ilVefCode   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                                tmRotInfo(UBound(tmRotInfo) - 1).iRevised = False
                                'If ilCombineIndex = -1 Then
                                    tmRotInfo(UBound(tmRotInfo) - 1).iSortCrfIndex = ilCrf
                                'Else
                                '    tmRotInfo(UBound(tmRotInfo) - 1).iSortCrfIndex = -1
                                'End If
                                'Determine rotation status (status = 1 = send; =2=duplicate- don't send)
                                tmRotInfo(UBound(tmRotInfo) - 1).iStatus = 1
                                'Determine if this is a revision
                                tmCrfSrchKey1.sRotType = tmCrf.sRotType
                                tmCrfSrchKey1.iEtfCode = tmCrf.iEtfCode
                                tmCrfSrchKey1.iEnfCode = tmCrf.iEnfCode
                                tmCrfSrchKey1.iAdfCode = tmCrf.iAdfCode
                                tmCrfSrchKey1.lChfCode = tmCrf.lChfCode
                                tmCrfSrchKey1.iVefCode = tmCrf.iVefCode
                                tmCrfSrchKey1.iRotNo = tmCrf.iRotNo
                                ilRet = btrGetGreaterOrEqual(hmCrf, tlCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                Do While (tlCrf.sRotType = tmCrf.sRotType) And (tlCrf.iEtfCode = tmCrf.iEtfCode) And (tlCrf.iEnfCode = tmCrf.iEnfCode) And (tlCrf.iAdfCode = tmCrf.iAdfCode) And (tlCrf.lChfCode = tmCrf.lChfCode) And (tlCrf.iVefCode = tmCrf.iVefCode) And (tlCrf.iRotNo < tmCrf.iRotNo)
                                    If (((rbcInterface(0).Value) And (tlCrf.sAffFdStatus = "S")) Or ((rbcInterface(1).Value) And (tlCrf.sKCFdStatus = "S"))) And (tlCrf.sInOut = tmCrf.sInOut) And (tlCrf.ianfCode = tmCrf.ianfCode) And (tlCrf.iLen = tmCrf.iLen) Then
                                        If tlCrf.sState <> "D" Then
                                            gUnpackDate tlCrf.iStartDate(0), tlCrf.iStartDate(1), slDate
                                            llSTStartDate = gDateValue(slDate)
                                            gUnpackDate tlCrf.iEndDate(0), tlCrf.iEndDate(1), slDate
                                            llSTEndDate = gDateValue(slDate)
                                            If (llRotEndDate <= llSTStartDate) And (llSTEndDate >= llRotStartDate) Then
                                                gUnpackTimeLong tlCrf.iStartTime(0), tlCrf.iStartTime(1), False, llSTStartTime
                                                gUnpackTimeLong tlCrf.iEndTime(0), tlCrf.iEndTime(1), True, llSTEndTime
                                                If (llRotEndTime <= llSTStartTime) And (llSTEndTime >= llRotStartTime) Then
                                                    tmRotInfo(UBound(tmRotInfo) - 1).iRevised = True
                                                End If
                                            End If
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmCrf, tlCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            End If
                    '    End If
                    Next ilVeh
                    tmCnfSrchKey.lCrfCode = tmCrf.lCode
                    tmCnfSrchKey.iInstrNo = 0
                    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                        'Loop thru all airing vehicle if selling, otherwise test conventional
                        'For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
                        '    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) <> 0 Then
                        For ilVeh = 0 To UBound(imVefCode) - 1 Step 1
                            ilVefCode = imVefCode(ilVeh)
                                ilTransmit = True
                                tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                                tmCyfSrchKey.iVefCode = ilVefCode   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                                If rbcInterface(0).Value Then
                                    tmCyfSrchKey.sSource = "S"
                                Else
                                    tmCyfSrchKey.sSource = "K"
                                End If
                                If Trim$(tmCrf.sZone) <> "R" Then
                                    tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                    tmCyfSrchKey.lRafCode = 0
                                Else
                                    tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                    tmCyfSrchKey.lRafCode = tmCrf.lRafCode
                                End If
                                ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                                    gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slFeedDate
                                    ilFdDateNew = False
                                    'If gDateValue(slFeedDate) + 90 > llTranDate Then
                                    '    'If airing- check if other vehicles has receive inventory
                                    '    For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
                                    '        If tmVef(ilLoop1).iCode = ilVefCode Then    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                    '            If (tmVef(ilLoop1).sType = "A") Or (tmVef(ilLoop1).sType = "C") Then
                                    '                ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                                    '                ilTransmit = False
                                    '                'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
                                    '                If ilVIndex >= 0 Then
                                    '                    ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                                    '                    Do While ilVIndex >= 0
                                    '                        tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                                    '                        tmCyfSrchKey.iVefCode = tmLkVehInfo(ilVIndex).iVefCode'tmVpfInfo(ilVIndex).iVefLink(ilLoop2)
                                    '                        tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                    '                        ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                    '                        If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                                    '                            gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slFeedDate
                                    '                            If gDateValue(slFeedDate) + 90 <= llTranDate Then
                                    '                                ilTransmit = True
                                    '                            End If
                                    '                        Else
                                    '                            ilTransmit = True
                                    '                            Exit For
                                    '                        End If
                                    '                    'Next ilLoop2
                                    '                        ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                                    '                    Loop
                                    '                End If
                                    '            Else
                                    '                ilTransmit = False
                                    '            End If
                                    '        End If
                                    '    Next ilLoop1
                                    'End If
                                Else
                                    ilTransmit = True
                                    slFeedDate = smTranDate
                                    ilFdDateNew = True
                                End If
                                If ilTransmit Then
                                    'Test for duplicates
                                    ilFound = False
                                    For ilLoop = LBound(tmAddCyf) To UBound(tmAddCyf) - 1 Step 1
                                        'Create records for each zone but only show cart for
                                        'one zone in the inventory area
                                        'in mComputeTime- the zone test is removed
                                        If (tmAddCyf(ilLoop).tCyf.lCifCode = tmCnf.lCifCode) And (tmAddCyf(ilLoop).tCyf.iVefCode = ilVefCode) Then   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) Then
                                            If llRotStartDate < tmAddCyf(ilLoop).lRotStartDate Then
                                                tmAddCyf(ilLoop).lRotStartDate = llRotStartDate
                                            End If
                                            If llRotEndDate > tmAddCyf(ilLoop).lRotEndDate Then
                                                tmAddCyf(ilLoop).lRotEndDate = llRotEndDate
                                            End If
                                            If Trim$(tmCrf.sZone) <> "R" Then
                                                If (tmAddCyf(ilLoop).tCyf.sTimeZone = tmCrf.sZone) And (tmAddCyf(ilLoop).tCyf.lRafCode = 0) Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Else
                                                If (tmAddCyf(ilLoop).tCyf.sTimeZone = tmCrf.sZone) And (tmAddCyf(ilLoop).tCyf.lRafCode = tmCrf.lRafCode) Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next ilLoop
                                    If Not ilFound Then
                                        For ilName = 0 To UBound(tmVef) - 1 Step 1
                                            If tmVef(ilName).iCode = ilVefCode Then 'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                                ilFound = True
                                                slKey = tmVef(ilName).sName
                                                ilVpf = mFindVpfIndex(ilVefCode)    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh))
                                                'If ilVpf > 0 Then
                                                '    For ilLoop1 = LBound(tmVpfInfo(ilVpf).sVefName) To tmVpfInfo(ilVpf).iNoVefLinks - 1 Step 1
                                                '        slKey = Trim$(slKey) & " " & tmVpfInfo(ilVpf).sVefName(ilLoop1)
                                                '    Next ilLoop1
                                                'End If
                                                If ilVpf >= 0 Then
                                                    ilVpf = tmVpfInfo(ilVpf).iFirstLkVehInfo
                                                    Do While ilVpf >= 0
                                                        slKey = Trim$(slKey) & " " & tmLkVehInfo(ilVpf).sVefName
                                                        ilVpf = tmLkVehInfo(ilVpf).iNextLkVehInfo
                                                    Loop
                                                End If
                                                Exit For
                                            End If
                                        Next ilName
                                        If ilFound Then
                                            ReDim Preserve tmAddCyf(0 To UBound(tmAddCyf) + 1) As SENDCOPYINFO
                                            slProduct = tgSortCrf(ilCrf).sCntrProd
                                            tmCifSrchKey.lCode = tmCnf.lCifCode
                                            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet = BTRV_ERR_NONE Then
                                                    'If Trim$(tmCpf.sName) <> "" Then
                                                    If (Trim$(tmCpf.sName) <> "") And (tgSpf.sUseProdSptScr <> "P") Then
                                                        slKey = slKey & "|" & tmCpf.sName
                                                        slProduct = tmCpf.sName
                                                    Else
                                                        slKey = slKey & "|" & tgSortCrf(ilCrf).sCntrProd
                                                    End If
                                                Else
                                                    slKey = slKey & "|" & tgSortCrf(ilCrf).sCntrProd
                                                    tmCpf.sISCI = ""
                                                    tmCpf.sCreative = ""
                                                End If
                                            Else
                                                slKey = slKey & "|" & tgSortCrf(ilCrf).sCntrProd
                                                tmCpf.sISCI = ""
                                                tmCpf.sCreative = ""
                                            End If
                                            If tmMcf.iCode <> tmCif.iMcfCode Then
                                                tmMcfSrchKey.iCode = tmCif.iMcfCode
                                                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    tmMcf.sName = ""
                                                End If
                                            End If
                                            slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                                            If (Len(Trim$(tmCif.sCut)) <> 0) Then
                                                slCart = slCart & "-" & tmCif.sCut
                                            End If
                                            'slStr = Trim$(Str$(tmCnf.lCifCode))
                                            'Do While Len(slStr) < 6
                                            '    slStr = "0" & slStr
                                            'Loop
                                            slKey = slKey & "|" & slCart    'slStr
                                            slStr = Trim$(str$(llRotStartDate))
                                            Do While Len(slStr) < 6
                                                slStr = "0" & slStr
                                            Loop
                                            slKey = slKey & "|" & slStr
                                            slStr = Trim$(str$(llRotEndDate))
                                            Do While Len(slStr) < 6
                                                slStr = "0" & slStr
                                            Loop
                                            slKey = slKey & "|" & slStr
                                            tmAddCyf(UBound(tmAddCyf) - 1).sKey = slKey
                                            tmAddCyf(UBound(tmAddCyf) - 1).sXFKey = slProduct & "|" & slCart & "|" & tmCpf.sISCI & "|" & tmCpf.sCreative
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.lCifCode = tmCnf.lCifCode
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.iVefCode = ilVefCode    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                                            If rbcInterface(0).Value Then
                                                tmAddCyf(UBound(tmAddCyf) - 1).tCyf.sSource = "S"
                                            Else
                                                tmAddCyf(UBound(tmAddCyf) - 1).tCyf.sSource = "K"
                                            End If
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.sTimeZone = tmCrf.sZone
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.iFeedDate(0) = ilTranDate0
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.iFeedDate(1) = ilTranDate1
                                            If Trim$(tmCrf.sZone) <> "R" Then
                                                tmAddCyf(UBound(tmAddCyf) - 1).tCyf.lRafCode = 0
                                            Else
                                                tmAddCyf(UBound(tmAddCyf) - 1).tCyf.lRafCode = tmCrf.lRafCode
                                            End If
                                            If rbcInterface(0).Value Then
                                                tmAddCyf(UBound(tmAddCyf) - 1).tCyf.sAffOrigXMitChar = smRunLetter
                                                gPackDateLong lmInputStartDate, tmAddCyf(UBound(tmAddCyf) - 1).tCyf.iAffOrigXMitDate(0), tmAddCyf(UBound(tmAddCyf) - 1).tCyf.iAffOrigXMitDate(1)
                                            Else
                                                tmAddCyf(UBound(tmAddCyf) - 1).tCyf.sKCOrigXMitChar = smRunLetter
                                                gPackDateLong lmInputStartDate, tmAddCyf(UBound(tmAddCyf) - 1).tCyf.iKCOrigXMitDate(0), tmAddCyf(UBound(tmAddCyf) - 1).tCyf.iKCOrigXMitDate(1)
                                            End If
                                            tmAddCyf(UBound(tmAddCyf) - 1).sChfProduct = Trim$(tgSortCrf(ilCrf).sCntrProd)
                                            tmAddCyf(UBound(tmAddCyf) - 1).lRotStartDate = llRotStartDate
                                            tmAddCyf(UBound(tmAddCyf) - 1).lRotEndDate = llRotEndDate
                                            tmAddCyf(UBound(tmAddCyf) - 1).lPrevFdDate = gDateValue(slFeedDate)
                                            tmAddCyf(UBound(tmAddCyf) - 1).iFdDateNew = ilFdDateNew
                                            tmAddCyf(UBound(tmAddCyf) - 1).iAdfCode = tmCrf.iAdfCode
                                            tmAddCyf(UBound(tmAddCyf) - 1).iLen = tmCrf.iLen
                                        End If
                                    End If
                                End If
                        '    End If
                        Next ilVeh
                        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                If ilCombineIndex = -1 Then
                    ilCombineIndex = tgSortCrf(ilCrf).iCombineIndex
                ElseIf ilCombineIndex >= 0 Then
                    ilCombineIndex = tgCombineCrf(ilCombineIndex).iCombineIndex
                End If
                If ilCombineIndex < 0 Then
                    ilCombineIndex = -2
                Else
                    tmCrf = tgCombineCrf(ilCombineIndex).tCrf
                End If
                If ilCombineIndex = -2 Then
                    If ilDuplIndex = -1 Then
                        ilDuplIndex = tgSortCrf(ilCrf).iDuplIndex
                    Else
                        ilDuplIndex = tgDuplCrf(ilDuplIndex).iDuplIndex
                    End If
                    If ilDuplIndex < 0 Then
                        ilDone = True
                    Else
                        tmCrf = tgDuplCrf(ilDuplIndex).tCrf
                    End If
                End If
            Loop While Not ilDone
        End If
    Next ilCrf
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCheckRegions                   *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Check That rotation defined for*
'*                      each region                    *
'*                                                     *
'*******************************************************
Private Function mCheckRegions() As Integer
    Dim ilStn As Integer
    Dim ilCrf As Integer
    Dim ilFound As Integer
    mCheckRegions = True
    For ilStn = 0 To UBound(tmStnInfo) - 1 Step 1
        If tmStnInfo(ilStn).lRegionCode > 0 Then
            ilFound = False
            For ilCrf = 0 To UBound(tgSortCrf) - 1 Step 1
                tmCrf = tgSortCrf(ilCrf).tCrf
                If Trim$(tmCrf.sZone) = "R" Then
                    If tmStnInfo(ilStn).lRafCode = tmCrf.lRafCode Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilCrf
            If Not ilFound Then
                'Print #hmMsg, "Rotation missing for Region ID:" & str$(tmStnInfo(ilStn).lRegionCode)
                gAutomationAlertAndLogHandler "Rotation missing for Region ID:" & str$(tmStnInfo(ilStn).lRegionCode)
                mCheckRegions = False
            End If
        End If
    Next ilStn
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCPR                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear CPR                      *
'*                                                     *
'*******************************************************
Private Sub mClearCPR()
    Dim ilRet As Integer
    Dim llGenTime As Long
    tmCprSrchKey.iGenDate(0) = imGenDate(0)
    tmCprSrchKey.iGenDate(1) = imGenDate(1)
    'tmCprSrchKey.iGenTime(0) = imGenTime(0)
    'tmCprSrchKey.iGenTime(1) = imGenTime(1)
    gUnpackTimeLong imGenTime(0), imGenTime(1), False, llGenTime
    tmCprSrchKey.lGenTime = llGenTime
    ilRet = btrGetGreaterOrEqual(hmCpr, tmCpr, imCprRecLen, tmCprSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmCpr.iGenDate(0) = imGenDate(0)) And (tmCpr.iGenDate(1) = imGenDate(1)) And (tmCpr.lGenTime = llGenTime)
        ilRet = btrDelete(hmCpr)
        tmCprSrchKey.iGenDate(0) = imGenDate(0)
        tmCprSrchKey.iGenDate(1) = imGenDate(1)
        'tmCprSrchKey.iGenTime(0) = imGenTime(0)
        'tmCprSrchKey.iGenTime(1) = imGenTime(1)
        tmCprSrchKey.lGenTime = llGenTime
        ilRet = btrGetGreaterOrEqual(hmCpr, tmCpr, imCprRecLen, tmCprSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Loop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearTxr                       *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove records from TXR        *
'*                                                     *
'*******************************************************
Private Sub mClearTxr()
    Dim ilRet As Integer
    Dim llGenTime As Long
    tmTxrSrchKey.iGenDate(0) = imPDFDate(0)
    tmTxrSrchKey.iGenDate(1) = imPDFDate(1)
    'tmTxrSrchKey.iGenTime(0) = imPDFTime(0)
    'tmTxrSrchKey.iGenTime(1) = imPDFTime(1)
    gUnpackTimeLong imPDFTime(0), imPDFTime(1), False, llGenTime
    tmTxr.lGenTime = llGenTime
    ilRet = btrGetGreaterOrEqual(hmTxr, tmTxr, imTxrRecLen, tmTxrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmTxr.iGenDate(0) = imPDFDate(0)) And (tmTxr.iGenDate(1) = imPDFDate(1)) And (tmTxr.lGenTime = llGenTime)
        ilRet = btrDelete(hmTxr)
        tmTxrSrchKey.iGenDate(0) = imPDFDate(0)
        tmTxrSrchKey.iGenDate(1) = imPDFDate(1)
        'tmTxrSrchKey.iGenTime(0) = imPDFTime(0)
        'tmTxrSrchKey.iGenTime(1) = imPDFTime(1)
        tmTxr.lGenTime = llGenTime
        ilRet = btrGetGreaterOrEqual(hmTxr, tmTxr, imTxrRecLen, tmTxrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Loop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mComputeTime                    *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Computer Inventory time         *
'*                                                     *
'*******************************************************
Private Sub mComputeTime(ilInc As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

'   mComputeTime ilInc
'   Where:
'       ilInc(I)- True =increment time; False=Remove time
'
'       tmCrf(I)- Rotation
'
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilLoop1 As Integer
    Dim ilVeh As Integer
    Dim ilVehIndex As Integer
    Dim ilVpfIndex As Integer
    Dim slDate As String
    Dim llTranDate As Long
    Dim ilTest As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slTime As String
    Dim slCode As String
    Dim ilPos As Integer
    Dim slMin As String
    Dim slSec As String
    Dim ilMin As Integer
    Dim ilSec As Integer
    Dim llTime As Long
    Dim ilVIndex As Integer
    Dim ilTransmit As Integer
    Dim ilVefCode As Integer
    Dim llIndex As Long
    Dim llLoop As Long

    ilVehIndex = -1
    For ilVeh = 0 To UBound(tmVef) - 1 Step 1
        If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
            ilVehIndex = ilVeh
            Exit For
        End If
    Next ilVeh
    If ilVehIndex = -1 Then
        For ilVeh = 0 To UBound(tmVef) - 1 Step 1
            If (tmVef(ilVeh).sType = "A") Or (tmVef(ilVeh).sType = "C") Then
                ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                If ilVIndex >= 0 Then
                    ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                    Do While ilVIndex >= 0
                        If tmCrf.iVefCode = tmLkVehInfo(ilVIndex).iVefCode Then
                            ilVehIndex = ilVeh
                            Exit For
                        End If
                        ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                    Loop
                End If
            Else
                ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                If ilVIndex >= 0 Then
                    ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                    Do While ilVIndex >= 0
                        If tmCrf.iVefCode = tmSALink(ilVIndex).iVefCode Then
                            ilVehIndex = ilVeh
                            Exit For
                        End If
                        ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                    Loop
                End If
            End If
        Next ilVeh
    End If
    If ilVehIndex = -1 Then
        Exit Sub
    End If
    ilVpfIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
    'iGLink will contain primary airing for group vehicles or airing vehicle
    'without groups
    'iGLink for Convention is zero- replace with conventional vehicle
    'If tmVef(ilVehIndex).sType <> "S" Then
    '    For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
    '        tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) = 0
    '    Next ilVeh
    '    tmVpfInfo(ilVpfIndex).tVpf.iGLink(LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink)) = tmVef(ilVehIndex).iCode
    'End If
    ReDim imVefCode(0 To 0) As Integer
    ilVefCode = tmVef(ilVehIndex).iCode
    If tmVef(ilVehIndex).sType = "S" Then
        'Find airing vehicle
        'ilSAGroupNo = tgVpf(gVpfFind(ExpStnFd, ilVefCode)).iSAGroupNo
        'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If (tgMVef(ilVeh).sType = "A") And (tgMVef(ilVeh).sState <> "D") Then
        '        If (ilSAGroupNo = tgVpf(gVpfFind(ExpStnFd, tgMVef(ilVeh).iCode)).iSAGroupNo) And (ilSAGroupNo <> 0) Then
        '            imVefCode(UBound(imVefCode)) = tgMVef(ilVeh).iCode
        '            ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
        '        End If
        '    End If
        'Next ilVeh
        If ilVpfIndex >= 0 Then
            ilVIndex = tmVpfInfo(ilVpfIndex).iFirstSALink
            Do While ilVIndex >= 0
                imVefCode(UBound(imVefCode)) = tmSALink(ilVIndex).iVefCode
                ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
                ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
            Loop
        End If
    Else
        ReDim imVefCode(0 To 1) As Integer
        imVefCode(0) = ilVefCode
    End If
    slDate = edcTranDate.Text
    llTranDate = gDateValue(slDate)
    tmCnfSrchKey.lCrfCode = tmCrf.lCode
    tmCnfSrchKey.iInstrNo = 0
    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
        'Loop thru all airing vehicle if selling, otherwise test conventional
        'For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
        '    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) <> 0 Then
        For ilVeh = LBound(imVefCode) To UBound(imVefCode) - 1 Step 1
            ilVefCode = imVefCode(ilVeh)
                ilTransmit = True
                'tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                'tmCyfSrchKey.iVefCode = ilVefCode   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                ''tmCyfSrchKey.sTimeZone = tmCrf.sZone
                'If Trim$(tmCrf.sZone) <> "R" Then
                '    tmCyfSrchKey.sTimeZone = tmCrf.sZone
                '    tmCyfSrchKey.lRafCode = 0
                'Else
                '    tmCyfSrchKey.sTimeZone = tmCrf.sZone
                '    tmCyfSrchKey.lRafCode = tmCrf.lRafCode
                'End If
                'ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                'If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                '    gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slFeedDate
                '    If gDateValue(slFeedDate) + 90 > llTranDate Then
                '        'If airing- check if other vehicles has receive inventory
                '        For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
                '            If tmVef(ilLoop1).iCode = ilVefCode Then    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                '                If (tmVef(ilLoop1).sType = "A") Or (tmVef(ilLoop1).sType = "C") Then
                '                    ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                '                    ilTransmit = False
                '                    'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
                '                    If ilVIndex >= 0 Then
                '                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                '                        Do While ilVIndex >= 0
                '                            tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                '                            tmCyfSrchKey.iVefCode = tmLkVehInfo(ilVIndex).iVefCode'tmVpfInfo(ilVIndex).iVefLink(ilLoop2)
                '                            tmCyfSrchKey.sTimeZone = tmCrf.sZone
                '                            ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                '                            If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                '                                gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slFeedDate
                '                                If gDateValue(slFeedDate) + 90 <= llTranDate Then
                '                                    ilTransmit = True
                '                                End If
                '                            Else
                '                                ilTransmit = True
                '                                Exit For
                '                            End If
                '                        'Next ilLoop2
                '                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                '                        Loop
                '                    End If
                '                Else
                '                    ilTransmit = False
                '                End If
                '            End If
                '        Next ilLoop1
                '    End If
                'End If
                If ilTransmit Then
                    'Test for duplicates
                    ilFound = False
                    For llLoop = LBound(tmCyfTest) To UBound(tmCyfTest) - 1 Step 1
                        ''If (tmCyfTest(ilLoop).lCifCode = tmCnf.lCifCode) And (tmCyfTest(ilLoop).iVefCode = tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) And (tmCyfTest(ilLoop).sTimeZone = tmCrf.sZone) Then
                        'If (tmCyfTest(ilLoop).lCifCode = tmCnf.lCifCode) And (tmCyfTest(ilLoop).iVefCode = ilVefCode) And (tmCyfTest(ilLoop).sTimeZone = tmCrf.sZone) Then
                        If (tmCyfTest(llLoop).lCifCode = tmCnf.lCifCode) And (tmCyfTest(llLoop).iVefCode = ilVefCode) And (tmCyfTest(llLoop).sTimeZone = tmCrf.sZone) And (tmCyfTest(llLoop).lRafCode = tmCrf.lRafCode) Then
                            'ilFound = True
                            If Not ilInc Then
                                For llIndex = ilLoop + 1 To UBound(tmCyfTest) - 1 Step 1
                                    'Ignore zone when counting time, when building
                                    'records in mBuildExpTable build one for each zone but one
                                    'show one piece of inventory without zone
                                    'If (tmCyfTest(ilIndex).lCifCode = tmCnf.lCifCode) And (tmCyfTest(ilIndex).iVefCode = tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) And (tmCyfTest(ilIndex).sTimeZone = tmCrf.sZone) Then
                                    If (tmCyfTest(llIndex).lCifCode = tmCnf.lCifCode) And (tmCyfTest(llIndex).iVefCode = ilVefCode) Then    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) Then
                                        ilFound = True  'Found a second time- retain time
                                    End If
                                    tmCyfTest(llIndex - 1).lCifCode = tmCyfTest(llIndex).lCifCode
                                    tmCyfTest(llIndex - 1).iVefCode = tmCyfTest(llIndex).iVefCode
                                    tmCyfTest(llIndex - 1).sSource = tmCyfTest(llIndex).sSource
                                    tmCyfTest(llIndex - 1).sTimeZone = tmCyfTest(llIndex).sTimeZone
                                    tmCyfTest(llIndex - 1).lRafCode = tmCyfTest(llIndex).lRafCode
                                Next llIndex
                                ReDim Preserve tmCyfTest(0 To UBound(tmCyfTest) - 1) As CYFTEST
                            Else
                                ilFound = True
                            End If
                            Exit For
                        Else
                            If (tmCyfTest(llLoop).lCifCode = tmCnf.lCifCode) And (tmCyfTest(llLoop).iVefCode = ilVefCode) Then  'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) Then
                                ilFound = True  'Found a second time- retain time
                            End If
                        End If
                    Next llLoop
                    If ilInc Then
                        ReDim Preserve tmCyfTest(0 To UBound(tmCyfTest) + 1) As CYFTEST
                        tmCyfTest(UBound(tmCyfTest) - 1).lCifCode = tmCnf.lCifCode
                        tmCyfTest(UBound(tmCyfTest) - 1).iVefCode = ilVefCode   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                        If rbcInterface(0).Value Then
                            tmCyfTest(UBound(tmCyfTest) - 1).sSource = "S"
                        Else
                            tmCyfTest(UBound(tmCyfTest) - 1).sSource = "K"
                        End If
                        tmCyfTest(UBound(tmCyfTest) - 1).sTimeZone = tmCrf.sZone
                        tmCyfTest(UBound(tmCyfTest) - 1).lRafCode = tmCrf.lRafCode
                    End If

                    If Not ilFound Then
                        'Adjust time
                        For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
                            If tmVef(ilLoop1).iCode = ilVefCode Then    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                                For ilTest = 0 To lbcVehicleCode.ListCount - 1 Step 1
                                    slNameCode = lbcVehicleCode.List(ilTest)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                    If ilVIndex = Val(slCode) Then
                                        'Obtain time, then adjust Vehicle|Time
                                        slNameCode = lbcVehicle.List(ilTest)
                                        ilRet = gParseItem(slNameCode, 1, "|", slName)
                                        ilRet = gParseItem(slNameCode, 2, "|", slTime)
                                        ilPos = InStr(slTime, ":")
                                        slMin = Left$(slTime, ilPos - 1)
                                        slSec = Mid$(slTime, ilPos + 1)
                                        llTime = 60 * Val(slMin) + Val(slSec)
                                        If ilInc Then
                                            llTime = llTime + tmCrf.iLen
                                        Else
                                            llTime = llTime - tmCrf.iLen
                                        End If
                                        ilSec = llTime Mod 60
                                        ilMin = llTime \ 60
                                        slTime = Trim$(str$(ilMin)) & ":" & Trim$(str$(ilSec))
                                        lbcVehicle.List(ilTest) = slName & "|" & slTime
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilTest
                                If ilFound Then
                                    Exit For
                                End If
                            End If
                        Next ilLoop1
                    End If
                End If
        '    End If
        Next ilVeh
        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateCartFile                 *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create cart to be sent file    *
'*                                                     *
'*******************************************************
Private Function mCreateCartFile() As Integer
    Dim slExportFile As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slKey As String
    Dim slProduct As String
    Dim slCart As String
    Dim slISCI As String
    Dim slCreative As String
    Dim slVehName As String
    Dim ilShowName As Integer
    Dim ilTest As Integer
    Dim ilLoop1 As Integer
    Dim ilVIndex As Integer
    Dim slBlank As String
    Dim slRecord As String
    Dim slBulkFeedGroup As String
    mCreateCartFile = False
    lacProcessing.Caption = "Generating Cart Reference"
    slBlank = " "
    If rbcInterface(0).Value Then
        slExportFile = sgExportPath & Mid$(smFeedNo, 3, 4) & Mid$(smFeedNo, 1, 2) & Left$(smGenTime, 2) & ".wem"
    Else
        slExportFile = sgExportPath & Format(smTranDate, "yyyy-mm-dd") & Left$(smGenTime, 2) & Mid$(smGenTime, 4, 2) & ".wem"
    End If
    DoEvents
    ilRet = 0
    'On Error GoTo mCreateCartFileErr:
    'hmExport = FreeFile
    ''Create file name based on vehicle name
    'Open slExportFile For Output As hmExport
    ilRet = gFileOpen(slExportFile, "Output", hmExport)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        ''MsgBox "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gAutomationAlertAndLogHandler "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Function
    End If
    For ilLoop = 0 To UBound(tmWemCyf) - 1 Step 1
        slKey = Trim$(tmWemCyf(ilLoop).sXFKey)
        ilRet = gParseItem(slKey, 1, "|", slProduct)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 2, "|", slCart)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 3, "|", slISCI)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 4, "|", slCreative)  'Obtain Index and code number
        slKey = Trim$(tmWemCyf(ilLoop).sKey)
        ilRet = gParseItem(slKey, 1, "|", slVehName)  'Obtain Index and code number
        ilShowName = False
        slBulkFeedGroup = ""
        'For ilTest = 0 To UBound(tgVpf) Step 1
        '    If tmWemCyf(ilLoop).tCyf.iVefCode = tgVpf(ilTest).iVefKCode Then
            ilTest = gBinarySearchVpf(tmWemCyf(ilLoop).tCyf.iVefCode)
            If ilTest <> -1 Then
                'If tgVpf(ilTest).sBulkXFer = "Y" Then
                '    ilShowName = True
                'End If
                If tgVpf(ilTest).sStnFdXRef = "Y" Then
                    ilShowName = True
                End If
                slBulkFeedGroup = tgVpf(ilTest).sGGroupNo
        '        Exit For
            End If
        'Next ilTest
        If Not ilShowName Then
            'Test if airing and group vehicles defined- if so test other
            'vehicles
            For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
                If tmVef(ilLoop1).iCode = tmWemCyf(ilLoop).tCyf.iVefCode Then
                    If (tmVef(ilLoop1).sType = "A") Or (tmVef(ilLoop1).sType = "C") Then
                        'Update cyf for all vehicles
                        ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                        'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
                        If ilVIndex >= 0 Then
                            ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                            Do While ilVIndex >= 0
                                'For ilTest = 0 To UBound(tgVpf) Step 1
                                '    'If tmVpfInfo(ilVIndex).iVefLink(ilLoop2) = tgVpf(ilTest).iVefKCode Then
                                '    If tmLkVehInfo(ilVIndex).iVefCode = tgVpf(ilTest).iVefKCode Then
                                    ilTest = gBinarySearchVpf(tmLkVehInfo(ilVIndex).iVefCode)
                                    If ilTest <> -1 Then
                                        'If tgVpf(ilTest).sBulkXFer = "Y" Then
                                        '    ilShowName = True
                                        'End If
                                        If tgVpf(ilTest).sStnFdXRef = "Y" Then
                                            ilShowName = True
                                        End If
                                '        Exit For
                                    End If
                                'Next ilTest
                                If ilShowName Then
                                    Exit For
                                End If
                            'Next ilLoop2
                                ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                            Loop
                        End If
                        If ilShowName Then
                            Exit For
                        End If
                    End If
                End If
            Next ilLoop1
        End If
        If ilShowName Then
            ''Cart #
            'slRecord = """" & slCart & """"
            ''ISCI Code
            'slRecord = slRecord & "," & """" & gFileNameFilter(slISCI) & """"
            ''Advertiser Name
            'tmCifSrchKey.lCode = tmWemCyf(ilLoop).tCyf.lCifCode
            'ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            'tmAdfSrchKey.iCode = tmCif.iAdfCode
            'ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            'slRecord = slRecord & "," & """" & Trim$(tmAdf.sName) & """"
            ''Short Title
            'slRecord = slRecord & "," & """" & slProduct & """"
            ''Airing Vehicle name
            'slRecord = slRecord & "," & """" & slVehName & """"
            'Cart #
            slRecord = slCart
            'ISCI Code
            slRecord = slRecord & "," & gFileNameFilter(slISCI)
            'Advertiser Name
            tmCifSrchKey.lCode = tmWemCyf(ilLoop).tCyf.lCifCode
            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            tmAdfSrchKey.iCode = tmCif.iAdfCode
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            slRecord = slRecord & "," & gAdvtNameFilter(Trim$(tmAdf.sName))
            'Short Title
            'slRecord = slRecord & "," & gAdvtNameFilter(slProduct)
            slRecord = slRecord & "," & gFileNameFilter(slProduct)  'gAdvtNameFilter(slProduct)
            'Airing Vehicle name
            slRecord = slRecord & "," & gAdvtNameFilter(slVehName)
            'Spot Length
            slRecord = slRecord & "," & Trim$(str$(tmCif.iLen))
            'Start Date
            slRecord = slRecord & "," & Format$(tmWemCyf(ilLoop).lRotStartDate, "mm/dd/yyyy")
            'End Date
            slRecord = slRecord & "," & Format$(tmWemCyf(ilLoop).lRotEndDate, "mm/dd/yyyy")
            'Regional copy flag
            If tmWemCyf(ilLoop).tCyf.lRafCode > 0 Then
                slRecord = slRecord & ",Y"
            Else
                slRecord = slRecord & ",N"
            End If
            'Bulk Feed Group Number
            slRecord = slRecord & "," & slBulkFeedGroup
            Print #hmExport, slRecord
        End If
    Next ilLoop
    Close hmExport
    mCreateCartFile = True
    Exit Function
'mCreateCartFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateCrossRef                 *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create Cross reference of carts*
'*                      to be sent                     *
'*                                                     *
'*******************************************************
Private Function mCreateCrossRef() As Integer
    Dim slStnCode As String
    Dim slXRefLetter As String
    Dim slExportFile As String
    Dim ilRet As Integer
    Dim slTimeStamp As String
    Dim ilLoop As Integer
    Dim slKey As String
    Dim ilUpper As Integer
    Dim ilPageNo As Integer
    Dim ilLineNo As Integer
    Dim slLine As String
    Dim slPrevProdISCITitle As String
    Dim slProdISCITitle As String
    Dim slProduct As String
    Dim slCart As String
    Dim slISCI As String
    Dim slCreative As String
    Dim slName As String
    Dim ilShowName As Integer
    Dim ilTest As Integer
    Dim ilLoop1 As Integer
    Dim ilVIndex As Integer
    Dim slBlank As String
    Dim slRecord As String
    Dim slPrevVehName As String
    Dim ilOldPageNo As Integer
    Dim ilPrtFirstXRef As Integer
    lacProcessing.Caption = "Generating Cross Reference"
    ilPrtFirstXRef = True
    slBlank = " "
    slStnCode = "X"
    'slXRefLetter = "A"
    slXRefLetter = smRunLetter
    If rbcInterface(0).Value Then
        Do
            slExportFile = sgExportPath & slStnCode & smFeedNo & slXRefLetter & ".xrf"
            ilRet = 0
            'On Error GoTo mCreateCrossRefErr:
            'slTimeStamp = FileDateTime(slExportFile)
            ilRet = gFileExist(slExportFile)
            If ilRet = 0 Then
                slXRefLetter = Chr$(Asc(slXRefLetter) + 1)
            End If
        Loop While ilRet = 0    'equal zero if file exist
    Else
        slExportFile = sgExportPath & Format(smTranDate, "yyyy-mm-dd") & Left$(smGenTime, 2) & Mid$(smGenTime, 4, 2) & ".xrf"
        'On Error GoTo mCreateCrossRefErr:
        'slTimeStamp = FileDateTime(slExportFile)
        ilRet = gFileExist(slExportFile)
        If ilRet = 0 Then
            Kill slExportFile
        End If
    End If
    DoEvents
    'Make key Short Title Cart # ISCI Creative Title, then vehicle name
    For ilLoop = 0 To UBound(tmXRefCyf) - 1 Step 1
        slKey = Trim$(tmXRefCyf(ilLoop).sKey)
        slKey = tmXRefCyf(ilLoop).sXFKey & "|" & slKey
        tmXRefCyf(ilLoop).sKey = slKey
    Next ilLoop
    ilUpper = UBound(tmXRefCyf)
    If ilUpper > 0 Then
        'ArraySortTyp fnAV(tgSort(),0), ilUpper, 0, LenB(tgSort(0)), 0, -9, 0
        ArraySortTyp fnAV(tmXRefCyf(), 0), ilUpper, 0, LenB(tmXRefCyf(0)), 0, LenB(tmXRefCyf(0).sKey), 0
    End If
    slStnCode = "X"
    If rbcInterface(0).Value Then
        slExportFile = sgExportPath & slStnCode & smFeedNo & slXRefLetter & ".xrf"
    Else
        slExportFile = sgExportPath & Format(smTranDate, "yyyy-mm-dd") & Left$(smGenTime, 2) & Mid$(smGenTime, 4, 2) & ".xrf"
    End If
    ilRet = 0
    'On Error GoTo mCreateCrossRefErr:
    'hmExport = FreeFile
    ''Create file name based on vehicle name
    'Open slExportFile For Output As hmExport
    ilRet = gFileOpen(slExportFile, "Output", hmExport)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        ''MsgBox "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gAutomationAlertAndLogHandler "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Function
    End If
    'Output new inventory
    ilPageNo = 0
    ilLineNo = 48
    slLine = ""
    slPrevProdISCITitle = " "
    For ilLoop = 0 To UBound(tmXRefCyf) - 1 Step 1
        slKey = Trim$(tmXRefCyf(ilLoop).sKey)
        ilRet = gParseItem(slKey, 1, "|", slProduct)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 2, "|", slCart)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 3, "|", slISCI)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 4, "|", slCreative)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 5, "|", slName)  'Obtain Index and code number
        ilShowName = False
        'For ilTest = 0 To UBound(tgVpf) Step 1
        '    If tmXRefCyf(ilLoop).tCyf.iVefCode = tgVpf(ilTest).iVefKCode Then
            ilTest = gBinarySearchVpf(tmXRefCyf(ilLoop).tCyf.iVefCode)
            If ilTest <> -1 Then
                'If tgVpf(ilTest).sBulkXFer = "Y" Then
                '    ilShowName = True
                'End If
                If tgVpf(ilTest).sStnFdXRef = "Y" Then
                    ilShowName = True
                End If
        '        Exit For
            End If
        'Next ilTest
        If Not ilShowName Then
            'Test if airing and group vehicles defined- if so test other
            'vehicles
            For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
                If tmVef(ilLoop1).iCode = tmXRefCyf(ilLoop).tCyf.iVefCode Then
                    If (tmVef(ilLoop1).sType = "A") Or (tmVef(ilLoop1).sType = "C") Then
                        'Update cyf for all vehicles
                        ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                        'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
                        If ilVIndex >= 0 Then
                            ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                            Do While ilVIndex >= 0
                                'For ilTest = 0 To UBound(tgVpf) Step 1
                                '    'If tmVpfInfo(ilVIndex).iVefLink(ilLoop2) = tgVpf(ilTest).iVefKCode Then
                                '    If tmLkVehInfo(ilVIndex).iVefCode = tgVpf(ilTest).iVefKCode Then
                                    ilTest = gBinarySearchVpf(tmLkVehInfo(ilVIndex).iVefCode)
                                    If ilTest <> -1 Then
                                        'If tgVpf(ilTest).sBulkXFer = "Y" Then
                                        '    ilShowName = True
                                        'End If
                                        If tgVpf(ilTest).sStnFdXRef = "Y" Then
                                            ilShowName = True
                                        End If
                                '        Exit For
                                    End If
                                'Next ilTest
                                If ilShowName Then
                                    Exit For
                                End If
                            'Next ilLoop2
                                ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                            Loop
                        End If
                        If ilShowName Then
                            Exit For
                        End If
                    End If
                End If
            Next ilLoop1
        End If
        If ilShowName Then
            Do While Len(slProduct) < 20
                slProduct = slProduct & " "
            Loop
            Do While Len(slCart) < 10
                slCart = slCart & " "
            Loop
            Do While Len(slISCI) < 20
                slISCI = slISCI & " "
            Loop
            Do While Len(slCreative) < 30
                slCreative = slCreative & " "
            Loop
            slProdISCITitle = slProduct & " " & slCart & " " & slISCI & " " & slCreative
            If slPrevProdISCITitle <> slProdISCITitle Then
                'First- output blank line
                If slPrevProdISCITitle <> " " Then
                    If Not mExportLine(slBlank, ilLineNo, -1) Then
                        Exit Function
                    End If
                End If
                slPrevProdISCITitle = slProdISCITitle
                slLine = slProdISCITitle
                slLine = slLine & " " & slName
                slPrevVehName = ""
            Else
                slLine = " "
                Do While Len(slLine) < Len(slProdISCITitle)
                    slLine = slLine & " "
                Loop
                slLine = slLine & " " & slName
            End If
            If slPrevVehName <> slName Then
                ilOldPageNo = ilPageNo
                '6/3/16: Replaced GoSub
                'GoSub cmcExportXFerHeader
                If Not mExportXFerHeader(ilPrtFirstXRef, ilLineNo, ilPageNo, slRecord) Then
                    Exit Function
                End If
                If (ilOldPageNo <> ilPageNo) And (ilOldPageNo > 0) Then
                    slLine = slProdISCITitle
                    slLine = slLine & " " & slName
                End If
                If Not mExportLine(slLine, ilLineNo, -1) Then
                    Exit Function
                End If
                slPrevVehName = slName
            End If
        End If
    Next ilLoop
    Close hmExport
    mCreateCrossRef = True
    Exit Function
'mCreateCrossRefErr:
'    ilRet = Err.Number
'    Resume Next
'cmcExportXFerHeader:
'    'If ilLineNo >= 48 Then
'    If ilPrtFirstXRef Then
'        ilPrtFirstXRef = False
'        If ilPageNo = 0 Then
'            slRecord = ""
'            If Not mExportLine(slRecord, ilLineNo, -1) Then
'                Exit Function
'            End If
'        Else
'            slRecord = Chr(12)  'Form Feed
'            If Not mExportLine(slRecord, ilLineNo, -1) Then
'                Exit Function
'            End If
'        End If
'        ilPageNo = ilPageNo + 1
'        ilLineNo = 0
'        slRecord = " "
'        Do While Len(slRecord) < 35
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & Trim$(tgSpf.sGClient)
'        If Not mExportLine(slRecord, ilLineNo, -1) Then
'            Exit Function
'        End If
'        slRecord = " "
'        Do While Len(slRecord) < 35
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & "Cross Reference"
'        If Not mExportLine(slRecord, ilLineNo, -1) Then
'            Exit Function
'        End If
'        slRecord = " "
'        Do While Len(slRecord) < 35
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & smTranDate & "  "
'        'slRecord = slRecord & "Page:"
'        'slStr = Trim$(Str$(ilPageNo))
'        'Do While Len(slStr) < 5
'        '    slStr = " " & slStr
'        'Loop
'        'slRecord = slRecord & slStr
'        If Not mExportLine(slRecord, ilLineNo, -1) Then
'            Exit Function
'        End If
'        slRecord = ""
'        If Not mExportLine(slRecord, ilLineNo, -1) Then
'            Exit Function
'        End If
'        If Not mExportLine(slRecord, ilLineNo, -1) Then
'            Exit Function
'        End If
'        slRecord = "Short Title"
'        Do While Len(slRecord) < 20
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & " Cart"
'        Do While Len(slRecord) < 31
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & " ISCI"
'        Do While Len(slRecord) < 52
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & " Creative Title"
'        Do While Len(slRecord) < 83
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & " Vehicle"
'        If Not mExportLine(slRecord, ilLineNo, -1) Then
'            Exit Function
'        End If
'    End If
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateEnvFile                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create envelop file            *
'*                                                     *
'*                      Structure                      *
'*                      ENVELOPE                       *
'*                      FILE: fileName,,,,,,           *
'*                      ADDR: transportal #, Serial #, *
'*                                                     *
'*******************************************************
Private Function mCreateEnvFile() As Integer
    Dim ilLoop As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim ilRet As Integer
    Dim ilFirstNon As Integer
    Dim ilGen As Integer
    Dim ilVeh As Integer
    Dim ilL1Count As Integer
    Dim ilL2Count As Integer
    Dim slChar1 As String * 1
    Dim slChar2 As String * 1
    Dim slChar3 As String * 1
    Dim hlEnv As Integer
    Dim slStr As String
    Dim llDays As Long
    Dim ilCIndex1 As Integer
    Dim ilCIndex2 As Integer
    Dim ilIndex1 As Integer
    Dim ilIndex As Integer
    Dim ilfirstTime As Integer
    Dim slVehName As String
    Dim ilVpfIndex As Integer
    Dim slIncludeRotation As String
    Dim ilTran As Integer
    Dim ilShowCartFile As Integer

    lacProcessing.Caption = "Generating Station Envelope"
    mCreateEnvFile = True
    ilFirstNon = UBound(tmStnInfo)
    For ilLoop = 0 To UBound(tmStnInfo) - 1 Step 1
        If (tmStnInfo(ilLoop).sType = "G") Then
            ilFirstNon = ilLoop
            Exit For
        End If
    Next ilLoop
    slChar1 = "A"
    slChar2 = "A"
    slChar3 = "A"
    ilL1Count = 1
    ilL2Count = 1
    For ilLoop = 0 To ilFirstNon - 1 Step 1
        'If (Trim$(tmStnInfo(ilLoop).sFileName) <> "") And (tmStnInfo(ilLoop).sType = "S") Then
        ilfirstTime = True
        For ilIndex = 0 To ilLoop - 1 Step 1
            If (tmStnInfo(ilIndex).sType = "S") Then
                If (StrComp(Trim$(tmStnInfo(ilLoop).sSiteID), Trim$(tmStnInfo(ilIndex).sSiteID), 1) = 0) Then
                    ilfirstTime = False
                    Exit For
                End If
            End If
        Next ilIndex
        If (tmStnInfo(ilLoop).sType = "S") And (ilfirstTime) Then
            ilRet = 0
            'On Error GoTo mCreateEnvFileErr:
            If rbcInterface(0).Value Then
                slToFile = sgExportPath & Trim$(tmStnInfo(ilLoop).sCallLetter) & Trim$(Left$(tmStnInfo(ilLoop).sBand, 1)) & smWeekNo & smRunLetter & ".env"
            Else
                slToFile = sgExportPath & Trim$(tmStnInfo(ilLoop).sCallLetter) & "-" & Trim$(Left$(tmStnInfo(ilLoop).sBand, 1)) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".fzt"
            End If
            'slDateTime = FileDateTime(slToFile)
            ilRet = gFileExist(slToFile)
            If ilRet = 0 Then
                Kill slToFile
            End If
            ilRet = 0
            'On Error GoTo mCreateEnvFileErr:
            'hlEnv = FreeFile
            'Open slToFile For Output As hlEnv
            ilRet = gFileOpen(slToFile, "Output", hlEnv)
            If ilRet <> 0 Then
                'Print #hmMsg, "** Error opening Envelope file: " & slToFile & " Error" & str(err)
                gAutomationAlertAndLogHandler "** Error opening Envelope file: " & slToFile & " Error" & str(err)
                mCreateEnvFile = False
                Exit Function
            End If
            Print #hlEnv, "ENVELOPE"
            'slStr = "SUBJ: " & """" & "Commercial Feed" & ": " & Trim$(tmStnInfo(ilLoop).sCallLetter) & "-" & Trim$(tmStnInfo(ilLoop).sBand) & " " & smGenTime & " " & smGenDate & """"
            slStr = "SUBJ: " & """" & "ABC Commercial Feed " & Trim$(tmStnInfo(ilLoop).sCallLetter) & "-" & Trim$(tmStnInfo(ilLoop).sBand) & " (" & smRunLetter & ") " & smGenDate & """"
            Print #hlEnv, slStr
            slStr = "DESC: " & """" & "ABC Radio Network's Commercial Feed for Station " & Trim$(tmStnInfo(ilLoop).sCallLetter) & "-" & Trim$(tmStnInfo(ilLoop).sBand) & ", created on " & smGenTime & " on " & smGenDate
            slStr = slStr & ", for the week beginning on Monday, " & Format$(lmInputStartDate, "dd mmm yyyy") & """"
            Print #hlEnv, slStr
            'slStr = "AD: " & """" & Trim$(tmStnInfo(ilLoop).sCallLetter) & "-" & Trim$(tmStnInfo(ilLoop).sBand) & " " & smGenTime & " " & smGenDate & """"
            'Print #hlEnv, slStr
            'llDays = 86400 * (lmInputEndDate - gDateValue("1/1/1970") + 28) '# Days in Seconds
            'ABC request to chjange it from 4 weeks to two weeks (requested: 11/14/02, Done: 1/22/03)
            llDays = 86400 * (lmInputEndDate - gDateValue("1/1/1970") + 14) '# Days in Seconds
            slStr = "LIFE: " & "0" & "," & Trim$(str$(llDays))
            Print #hlEnv, slStr
            'If (Trim$(tmStnInfo(ilLoop).sFileName) <> "") And (tmStnInfo(ilLoop).sType = "S") Then
            '    'Print #hlEnv, "FILE:" & Trim$(tmStnInfo(ilLoop).sFileName) & ",,," & Trim$(tmStnInfo(ilLoop).sSiteID) & ",,,"
            '    slStr = "FILE: " & """" & "TI-" & Trim$(tmStnInfo(ilLoop).sCallLetter) & "-" & Trim$(tmStnInfo(ilLoop).sBand) & " for Week " & smWeekNo & smRunLetter & ".Trf" & """"
            '    slStr = slStr & ", " & """" & Trim$(tmStnInfo(ilLoop).sCallLetter) & "'s traffic instructions for the airing week beginning Monday, " & Format$(lmInputStartDate, "dd mmm yyyy") & """"
            '    slStr = slStr & ", " & """" & "Please air using the flight data and rotation indicated in this document" & """" & "," & """" & """" & ", " & "reserved, reserved, reserved"
            '    Print #hlEnv, slStr
            'Else
            '    For ilGen = ilFirstNon To UBound(tmStnInfo) - 1 Step 1
            '        ilPrint = False
            '        ilSIndex = ilLoop
            '        Do
            '            If tmStnInfo(ilSIndex).iAirVeh = tmStnInfo(ilGen).iAirVeh Then
            '                'Print #hlEnv, "FILE:" & Trim$(tmStnInfo(ilGen).sFileName) & ",,," & Trim$(tmStnInfo(ilLoop).sSiteID) & ",,,"
            '                slStr = "FILE: " & """" & "TI-" & Trim$(tmStnInfo(ilGen).sFileName) & ".Trf" & """"
            '                slStr = slStr & ", " & """" & Trim$(tmStnInfo(ilSIndex).sCallLetter) & "'s traffic instructions for the airing week beginning Monday, " & Format$(lmInputStartDate, "dd mmm yyyy") & """"
            '                slStr = slStr & ", " & """" & "Please air using the flight data and rotation indicated in this document" & """" & "," & """" & """" & ", " & "reserved, reserved, reserved"
            '                Print #hlEnv, slStr
            '                ilPrint = True
            '            Else
            '                For ilVeh = 0 To UBound(tmVef) - 1 Step 1
            '                    If tmVef(ilVeh).iCode = tmStnInfo(ilGen).iAirVeh Then
            '                        ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
            '                        If ilVIndex >= 0 Then
            '                            ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
            '                            Do While ilVIndex >= 0
            '                                If tmStnInfo(ilSIndex).iAirVeh = tmLkVehInfo(ilVIndex).iVefCode Then
            '                                    'Print #hlEnv, "FILE:" & Trim$(tmStnInfo(ilGen).sFileName) & ".Txt" & ",,," & Trim$(tmStnInfo(ilLoop).sSiteID) & ",,,"
            '                                    slStr = "FILE: " & """" & "TI-" & """" & Trim$(tmStnInfo(ilGen).sFileName) & ".Trf" & """"
            '                                    slStr = slStr & ", " & """" & Trim$(tmStnInfo(ilSIndex).sCallLetter) & "'s traffic instructions for the airing week beginning Monday, " & Format$(lmInputStartDate, "dd mmm yyyy") & """"
            '                                    slStr = slStr & ", " & """" & "Please air using the flight data and rotation indicated in this document" & """" & "," & """" & """" & ", " & "reserved, reserved, reserved"
            '                                    Print #hlEnv, slStr
            '                                    ilPrint = True
            '                                    Exit Do
            '                                End If
            '                                ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
            '                            Loop
            '                        End If
            '                        Exit For
            '                    End If
            '                Next ilVeh
            '            End If
            '            ilSIndex = tmStnInfo(ilSIndex).iLkStnInfo
            '        Loop While (ilSIndex >= 0) And (Not ilPrint)
            '    Next ilGen
            'End If
            'ilCIndex = tmStnInfo(ilLoop).iLkCartInfo
            'Do While ilCIndex <> -1
            '    slStr = "File: " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & "-" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & ".mp2" & """"
            '    slStr = slStr & ", " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & "-" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & """"
            '    slStr = slStr & ", " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & """"
            '    slStr = slStr & ", " & """" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & """" & ", " & "reserved, reserved, reserved"
            '    Print #hlEnv, slStr
            '    ilCIndex = tgCartStnXRef(ilCIndex).iLkCartInfo
            'Loop
            For ilIndex = ilLoop To UBound(tmStnInfo) - 1 Step 1
                If (StrComp(Trim$(tmStnInfo(ilIndex).sSiteID), Trim$(tmStnInfo(ilLoop).sSiteID), 1) = 0) Then
                    If (Trim$(tmStnInfo(ilIndex).sFileName) <> "") And (tmStnInfo(ilIndex).sType = "S") Then
                        'slStr = "FILE: " & """" & "TI-" & """" & Trim$(tmStnInfo(ilIndex).sCallLetter) & "-" & Trim$(tmStnInfo(ilIndex).sBand) & " for Week " & smWeekNo & smRunLetter & ".Txt" & """"
                        'slStr = slStr & ", " & """" & Trim$(tmStnInfo(ilIndex).sCallLetter) & "'s traffic instructions for the airing week beginning Monday, " & Format$(lmInputStartDate, "dd mmm yyyy") & """"
                        'slStr = slStr & ", " & """" & "Please air using the flight data and rotation indicated in this document" & """" & "," & """" & """" & ", " & "reserved, reserved, reserved"
                        'Print #hlEnv, slStr
                        slVehName = ""
                        slIncludeRotation = "Y"
                        'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If tgMVef(ilVeh).iCode = tmStnInfo(ilIndex).iAirVeh Then
                            ilVeh = gBinarySearchVef(tmStnInfo(ilIndex).iAirVeh)
                            If ilVeh <> -1 Then
                                slVehName = Trim$(tgMVef(ilVeh).sName)
                        '        Exit For
                                ilVpfIndex = gBinarySearchVpf(tgMVef(ilVeh).iCode)
                                If ilVpfIndex <> -1 Then
                                    If Trim$(tgVpf(ilVpfIndex).sKCGenRot) <> "" Then
                                        slIncludeRotation = tgVpf(ilVpfIndex).sKCGenRot
                                    End If
                                End If
                            End If
                        'Next ilVeh
                        If rbcInterface(0).Value Then
                            slStr = "FILE: " & """" & "Traffic Instructions for " & Trim$(tmStnInfo(ilIndex).sCallLetter) & "-" & Trim$(tmStnInfo(ilIndex).sBand) & " for Week " & smWeekNo & smRunLetter & " (" & slVehName & ")" & ".Txt" & """"
                            'slStr = slStr & "," & """" & Trim$(tmStnInfo(ilIndex).sCallLetter) & "-" & Trim$(tmStnInfo(ilIndex).sBand) & "'s traffic instructions for the " & slVehName & " network for the airing week beginning Monday, " & Format$(lmInputStartDate, "dd mmm yyyy") & """"
                            slStr = slStr & "," & """" & "ABC Traffic instructions for " & Trim$(tmStnInfo(ilIndex).sCallLetter) & "-" & Trim$(tmStnInfo(ilIndex).sBand) & " (" & slVehName & ")(" & smRunLetter & ")" & "-" & Format$(lmInputStartDate, "mm/dd/yy") & """"
                            slStr = slStr & "," & """" & "Please air using the flight data and rotation indicated in this document" & """" & "," & """" & """" & "," & "reserved,reserved,reserved"
                            Print #hlEnv, slStr
                        Else
                            If slIncludeRotation <> "N" Then
                                '11/8/06:  Change extension to be pdf if creating rotation in pdf format
                                'slStr = "FILE: " & """" & Trim$(tmStnInfo(ilIndex).sCallLetter) & "-" & Trim$(tmStnInfo(ilIndex).sBand) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Traffic_Instructions(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Txt" & """"
                                If rbcFormat(2).Value Then
                                    slStr = "FILE: " & """" & Trim$(tmStnInfo(ilIndex).sCallLetter) & "-" & Trim$(tmStnInfo(ilIndex).sBand) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Traffic_Instructions(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Pdf" & """"
                                Else
                                    slStr = "FILE: " & """" & Trim$(tmStnInfo(ilIndex).sCallLetter) & "-" & Trim$(tmStnInfo(ilIndex).sBand) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Traffic_Instructions(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Txt" & """"
                                End If
                                slStr = slStr & "," & """" & "ABC Traffic instructions for " & Trim$(tmStnInfo(ilIndex).sCallLetter) & "-" & Trim$(tmStnInfo(ilIndex).sBand) & " (" & slVehName & ")(" & smRunLetter & ")" & "-" & Format$(lmInputStartDate, "mm/dd/yy") & """"
                                slStr = slStr & "," & """" & "Please air using the flight data and rotation indicated in this document" & """" & ",,,," & Trim$(tmStnInfo(ilIndex).sKCNo) & "," & Format(lmInputStartDate, "mm-dd-yyyy")
                                Print #hlEnv, slStr
                            End If
                        End If
                    End If
                End If
            Next ilIndex
            'For ilGen = ilFirstNon To UBound(tmStnInfo) - 1 Step 1
            '    If (StrComp(Trim$(tmStnInfo(ilGen).sSiteID), Trim$(tmStnInfo(ilLoop).sSiteID), 1) = 0) Then
            '        If (Trim$(tmStnInfo(ilGen).sFileName) <> "") Then
            '            slStr = "FILE: " & """" & "TI-" & """" & Trim$(tmStnInfo(ilGen).sFileName) & ".Txt" & """"
            '            slStr = slStr & ", " & """" & Trim$(tmStnInfo(ilSIndex).sCallLetter) & "'s traffic instructions for the airing week beginning Monday, " & Format$(lmInputStartDate, "dd mmm yyyy") & """"
            '            slStr = slStr & ", " & """" & "Please air using the flight data and rotation indicated in this document" & """" & "," & """" & """" & ", " & "reserved, reserved, reserved"
            '            Print #hlEnv, slStr
            '        End If
            '    End If
            'Next ilGen
            For ilIndex = ilLoop To UBound(tmStnInfo) - 1 Step 1
                If (StrComp(Trim$(tmStnInfo(ilIndex).sSiteID), Trim$(tmStnInfo(ilLoop).sSiteID), 1) = 0) Then
                    If (Trim$(tmStnInfo(ilIndex).sFileName) = "") And (tmStnInfo(ilIndex).sType = "S") Then
                        For ilGen = ilFirstNon To UBound(tmStnInfo) - 1 Step 1
                            If tmStnInfo(ilGen).iAirVeh = tmStnInfo(ilIndex).iAirVeh Then
                                If (Trim$(tmStnInfo(ilGen).sFileName) <> "") Then
                                    'slStr = "FILE: " & """" & "TI-" & """" & Trim$(tmStnInfo(ilGen).sFileName) & ".Txt" & """"
                                    'slStr = slStr & ", " & """" & Trim$(tmStnInfo(ilIndex).sCallLetter) & "'s traffic instructions for the airing week beginning Monday, " & Format$(lmInputStartDate, "dd mmm yyyy") & """"
                                    'slStr = slStr & ", " & """" & "Please air using the flight data and rotation indicated in this document" & """" & "," & """" & """" & ", " & "reserved, reserved, reserved"
                                    'Print #hlEnv, slStr
                                    slVehName = ""
                                    slIncludeRotation = "Y"
                                    'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                    '    If tgMVef(ilVeh).iCode = tmStnInfo(ilGen).iAirVeh Then
                                        ilVeh = gBinarySearchVef(tmStnInfo(ilGen).iAirVeh)
                                        If ilVeh <> -1 Then
                                            slVehName = Trim$(tgMVef(ilVeh).sName)
                                    '        Exit For
                                            ilVpfIndex = gBinarySearchVpf(tgMVef(ilVeh).iCode)
                                            If ilVpfIndex <> -1 Then
                                                If Trim$(tgVpf(ilVpfIndex).sKCGenRot) <> "" Then
                                                    slIncludeRotation = tgVpf(ilVpfIndex).sKCGenRot
                                                End If
                                            End If
                                        End If
                                    'Next ilVeh
                                    If rbcInterface(0).Value Then
                                        slStr = "FILE: " & """" & "General Traffic Instructions for the " & slVehName & " Network for Week " & smWeekNo & smRunLetter & ".Txt" & """"
                                        'slStr = slStr & "," & """" & "Traffic Instructions for the " & slVehName & " network for the airing week beginning Monday, " & Format$(lmInputStartDate, "dd mmm yyyy") & """"
                                        slStr = slStr & "," & """" & "ABC Traffic instructions for- " & slVehName & " (" & smRunLetter & ")" & "-" & Format$(lmInputStartDate, "mm/dd/yy") & """"
                                        slStr = slStr & "," & """" & "Please air using the flight data and rotation indicated in this document" & """" & "," & """" & """" & "," & "reserved,reserved,reserved"
                                        Print #hlEnv, slStr
                                    Else
                                        If slIncludeRotation <> "N" Then
                                            '11/8/06:  Change extension to be pdf if creating rotation in pdf format
                                            'slStr = "FILE: " & """" & "General_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Traffic_Instructions(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Txt" & """"
                                            If rbcFormat(2).Value Then
                                                slStr = "FILE: " & """" & "General_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Traffic_Instructions(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Pdf" & """"
                                            Else
                                                slStr = "FILE: " & """" & "General_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Traffic_Instructions(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Txt" & """"
                                            End If
                                            'slStr = slStr & "," & """" & "Traffic Instructions for the " & slVehName & " network for the airing week beginning Monday, " & Format$(lmInputStartDate, "dd mmm yyyy") & """"
                                            slStr = slStr & "," & """" & "ABC Traffic instructions for- " & slVehName & " (" & smRunLetter & ")" & "-" & Format$(lmInputStartDate, "mm/dd/yy") & """"
                                            slStr = slStr & "," & """" & "Please air using the flight data and rotation indicated in this document" & """" & ",,,," & Trim$(tmStnInfo(ilIndex).sKCNo) & "," & Format$(lmInputStartDate, "mm/dd/yy")
                                            Print #hlEnv, slStr
                                        End If
                                    End If
                                End If
                            End If
                        Next ilGen
                    End If
                End If
            Next ilIndex
            If rbcInterface(1).Value Then
                If ckcCmmlLog(0).Value = vbChecked Then
                    For ilIndex = ilLoop To UBound(tmStnInfo) - 1 Step 1
                        If (StrComp(Trim$(tmStnInfo(ilIndex).sSiteID), Trim$(tmStnInfo(ilLoop).sSiteID), 1) = 0) Then
                            If (Trim$(tmStnInfo(ilIndex).sFileName) <> "") And (tmStnInfo(ilIndex).sType = "S") And (tmStnInfo(ilIndex).sCmmlLogReq = "L") Then
                                slVehName = ""
                                'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                '    If tgMVef(ilVeh).iCode = tmStnInfo(ilIndex).iAirVeh Then
                                    ilVeh = gBinarySearchVef(tmStnInfo(ilIndex).iAirVeh)
                                    If ilVeh <> -1 Then
                                        slVehName = Trim$(tgMVef(ilVeh).sName)
                                '        Exit For
                                    End If
                                'Next ilVeh
                                slStr = "FILE: " & """" & Trim$(tmStnInfo(ilIndex).sFileName) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Commercial_Log(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Pdf" & """"
                                slStr = slStr & "," & """" & "ABC Commercial instructions for " & Trim$(tmStnInfo(ilIndex).sCallLetter) & "-" & Trim$(tmStnInfo(ilIndex).sBand) & " (" & slVehName & ")(" & smRunLetter & ")" & "-" & Format$(lmInputStartDate, "mm/dd/yy") & """"
                                slStr = slStr & "," & """" & "Please air using the flight data and copy indicated in this document" & """" & ",,,," & Trim$(tmStnInfo(ilIndex).sKCNo) & "," & Format(lmInputStartDate, "mm-dd-yyyy")
                                Print #hlEnv, slStr
                            End If
                        End If
                    Next ilIndex
                End If
            End If
            For ilIndex = ilLoop To UBound(tmStnInfo) - 1 Step 1
                If (StrComp(Trim$(tmStnInfo(ilIndex).sSiteID), Trim$(tmStnInfo(ilLoop).sSiteID), 1) = 0) Then
                    If (Trim$(tmStnInfo(ilIndex).sFileName) <> "") And (tmStnInfo(ilIndex).sType = "S") Then
                        'ilCIndex = tmStnInfo(ilIndex).iLkCartInfo
                        'Do While ilCIndex <> -1
                        '    For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '        If tgMVef(ilVeh).iCode = tmStnInfo(ilIndex).iAirVeh Then
                        '            'slStr = "File: " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & ";" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & ".mp2" & """"
                        '            slStr = "File: " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & "(" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & ")" & ".mp2" & """"
                        '            'slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & """"
                        '            slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & "(" & Trim$(Str$(tgCartStnXRef(ilCIndex).iLen)) & ")" & """"
                        '            slStr = slStr & "," & """" & Trim$(tgMVef(ilVeh).sName) & """"
                        '            'slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & """" & "," & "reserved,reserved,reserved"
                        '            slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & """"
                        '            tmAdfSrchKey.iCode = tgCartStnXRef(ilCIndex).iAdfCode
                        '            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        '            If ilRet = BTRV_ERR_NONE Then
                        '                slStr = slStr & "," & """" & Trim$(tmAdf.sName) & """" & "," & "reserved,reserved"
                        '            Else
                        '                slStr = slStr & "," & "reserved,reserved,reserved"
                        '            End If
                        '            Print #hlEnv, slStr
                        '            Exit For
                        '        End If
                        '    Next ilVeh
                        '    ilCIndex = tgCartStnXRef(ilCIndex).iLkCartInfo
                        'Loop
                        ilCIndex1 = tmStnInfo(ilIndex).iLkCartInfo1
                        ilCIndex2 = tmStnInfo(ilIndex).iLkCartInfo2
                        Do While ilCIndex1 <> -1
                            'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            '    If tgMVef(ilVeh).iCode = tmStnInfo(ilIndex).iAirVeh Then
                                ilVeh = gBinarySearchVef(tmStnInfo(ilIndex).iAirVeh)
                                If ilVeh <> -1 Then
                                    If rbcInterface(0).Value Then
                                        ilShowCartFile = True
                                    Else
                                        If rbcEnv(1).Value Then
                                            'Show New Only
                                            If tgCartStnXRef(ilCIndex1, ilCIndex2).iFdDateNew Then
                                                ilShowCartFile = True
                                            Else
                                                ilShowCartFile = False
                                            End If
                                        ElseIf rbcEnv(2).Value Then
                                            'Use Vantive to determine setting
                                            If tmStnInfo(ilIndex).sKCEnvCopy = "A" Then
                                                ilShowCartFile = True
                                            Else
                                                If tgCartStnXRef(ilCIndex1, ilCIndex2).iFdDateNew Then
                                                    ilShowCartFile = True
                                                Else
                                                    ilShowCartFile = False
                                                End If
                                            End If
                                        Else
                                            'Show All
                                            ilShowCartFile = True
                                        End If
                                    End If
                                    If ilShowCartFile Then
                                        'slStr = "File: " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & ";" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & ".mp2" & """"
                                        slStr = "File: " & """" & Trim$(tgCartStnXRef(ilCIndex1, ilCIndex2).sShortTitle) & "(" & Trim$(tgCartStnXRef(ilCIndex1, ilCIndex2).sISCI) & ")" & ".mp2" & """"
                                        'slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & """"
                                        slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex1, ilCIndex2).sShortTitle) & "(" & Trim$(str$(tgCartStnXRef(ilCIndex1, ilCIndex2).iLen)) & ")" & """"
                                        slStr = slStr & "," & """" & Trim$(tgMVef(ilVeh).sName) & """"
                                        'slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & """" & "," & "reserved,reserved,reserved"
                                        slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex1, ilCIndex2).sISCI) & """"
                                        If tmAdf.iCode <> tgCartStnXRef(ilCIndex1, ilCIndex2).iAdfCode Then
                                            tmAdfSrchKey.iCode = tgCartStnXRef(ilCIndex1, ilCIndex2).iAdfCode
                                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        Else
                                            ilRet = BTRV_ERR_NONE
                                        End If
                                        If rbcInterface(0).Value Then
                                            If ilRet = BTRV_ERR_NONE Then
                                                slStr = slStr & "," & """" & Trim$(tmAdf.sName) & """" & "," & "reserved,reserved"
                                            Else
                                                slStr = slStr & "," & "reserved,reserved,reserved"
                                            End If
                                        Else
                                            If ilRet = BTRV_ERR_NONE Then
                                                slStr = slStr & "," & """" & Trim$(tmAdf.sName) & """"
                                            Else
                                                slStr = slStr & ","
                                            End If
                                            slStr = slStr & "," & Trim$(str$(tgCartStnXRef(ilCIndex1, ilCIndex2).iLen)) & "," & Trim$(tmStnInfo(ilIndex).sKCNo) & "," & Format$(lmInputStartDate, "mm/dd/yy")
                                        End If
                                        Print #hlEnv, slStr
                                    End If
                            '        Exit For
                                End If
                            'Next ilVeh
                            ilIndex1 = ilCIndex1
                            ilCIndex1 = tgCartStnXRef(ilIndex1, ilCIndex2).iLkCartInfo1
                            ilCIndex2 = tgCartStnXRef(ilIndex1, ilCIndex2).iLkCartInfo2
                        Loop
                    End If
                End If
            Next ilIndex
            'For ilGen = ilFirstNon To UBound(tmStnInfo) - 1 Step 1
            '    If (StrComp(Trim$(tmStnInfo(ilGen).sSiteID), Trim$(tmStnInfo(ilLoop).sSiteID), 1) = 0) Then
            '        If (Trim$(tmStnInfo(ilGen).sFileName) <> "") Then
            '            ilCIndex = tmStnInfo(ilGen).iLkCartInfo
            '            For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '                If tgMVef(ilVeh).iCode = tmStnInfo(ilGen).iAirVeh Then
            '                    slStr = "File: " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & "-" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & ".mp2" & """"
            '                    slStr = slStr & ", " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & """"
            '                    slStr = slStr & ", " & """" & Trim$(tgMVef(ilVeh).sName) & """"
            '                    slStr = slStr & ", " & """" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & """" & ", " & "reserved, reserved, reserved"
            '                    Print #hlEnv, slStr
            '                    Exit For
            '                End If
            '            Next ilVeh
            '            ilCIndex = tgCartStnXRef(ilCIndex).iLkCartInfo
            '        End If
            '    End If
            'Next ilGen
            For ilIndex = ilLoop To UBound(tmStnInfo) - 1 Step 1
                If (StrComp(Trim$(tmStnInfo(ilIndex).sSiteID), Trim$(tmStnInfo(ilLoop).sSiteID), 1) = 0) Then
                    If (Trim$(tmStnInfo(ilIndex).sFileName) = "") And (tmStnInfo(ilIndex).sType = "S") Then
                        For ilGen = ilFirstNon To UBound(tmStnInfo) - 1 Step 1
                            If tmStnInfo(ilGen).iAirVeh = tmStnInfo(ilIndex).iAirVeh Then
                                If (Trim$(tmStnInfo(ilGen).sFileName) <> "") Then
                                    'ilCIndex = tmStnInfo(ilGen).iLkCartInfo
                                    'Do While ilCIndex <> -1
                                    '    For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                    '        If tgMVef(ilVeh).iCode = tmStnInfo(ilGen).iAirVeh Then
                                    '            'slStr = "File: " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & ";" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & ".mp2" & """"
                                    '            slStr = "File: " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & "(" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & ")" & ".mp2" & """"
                                    '            'slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & """"
                                    '            slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & "(" & Trim$(Str$(tgCartStnXRef(ilCIndex).iLen)) & ")" & """"
                                    '            slStr = slStr & "," & """" & Trim$(tgMVef(ilVeh).sName) & """"
                                    '            'slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & """" & "," & "reserved,reserved,reserved"
                                    '            slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & """"
                                    '            tmAdfSrchKey.iCode = tgCartStnXRef(ilCIndex).iAdfCode
                                    '            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    '            If ilRet = BTRV_ERR_NONE Then
                                    '                slStr = slStr & "," & """" & Trim$(tmAdf.sName) & """" & "," & "reserved,reserved"
                                    '            Else
                                    '                slStr = slStr & "," & "reserved,reserved,reserved"
                                    '            End If
                                    '            Print #hlEnv, slStr
                                    '            Exit For
                                    '        End If
                                    '    Next ilVeh
                                    '    ilCIndex = tgCartStnXRef(ilCIndex).iLkCartInfo
                                    'Loop
                                    ilCIndex1 = tmStnInfo(ilGen).iLkCartInfo1
                                    ilCIndex2 = tmStnInfo(ilGen).iLkCartInfo2
                                    Do While ilCIndex1 <> -1
                                        'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                        '    If tgMVef(ilVeh).iCode = tmStnInfo(ilGen).iAirVeh Then
                                            ilVeh = gBinarySearchVef(tmStnInfo(ilGen).iAirVeh)
                                            If ilVeh <> -1 Then
                                                If rbcInterface(0).Value Then
                                                    ilShowCartFile = True
                                                Else
                                                    If rbcEnv(1).Value Then
                                                        'Show New Only
                                                        If tgCartStnXRef(ilCIndex1, ilCIndex2).iFdDateNew Then
                                                            ilShowCartFile = True
                                                        Else
                                                            ilShowCartFile = False
                                                        End If
                                                    ElseIf rbcEnv(2).Value Then
                                                        'Use Vantive to determine setting
                                                        If tmStnInfo(ilGen).sKCEnvCopy = "A" Then
                                                            ilShowCartFile = True
                                                        Else
                                                            If tgCartStnXRef(ilCIndex1, ilCIndex2).iFdDateNew Then
                                                                ilShowCartFile = True
                                                            Else
                                                                ilShowCartFile = False
                                                            End If
                                                        End If
                                                    Else
                                                        'Show All
                                                        ilShowCartFile = True
                                                    End If
                                                End If
                                                If ilShowCartFile Then
                                                    'slStr = "File: " & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & ";" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & ".mp2" & """"
                                                    slStr = "File: " & """" & Trim$(tgCartStnXRef(ilCIndex1, ilCIndex2).sShortTitle) & "(" & Trim$(tgCartStnXRef(ilCIndex1, ilCIndex2).sISCI) & ")" & ".mp2" & """"
                                                    'slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sShortTitle) & """"
                                                    slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex1, ilCIndex2).sShortTitle) & "(" & Trim$(str$(tgCartStnXRef(ilCIndex1, ilCIndex2).iLen)) & ")" & """"
                                                    slStr = slStr & "," & """" & Trim$(tgMVef(ilVeh).sName) & """"
                                                    'slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex).sISCI) & """" & "," & "reserved,reserved,reserved"
                                                    slStr = slStr & "," & """" & Trim$(tgCartStnXRef(ilCIndex1, ilCIndex2).sISCI) & """"
                                                    If tmAdf.iCode <> tgCartStnXRef(ilCIndex1, ilCIndex2).iAdfCode Then
                                                        tmAdfSrchKey.iCode = tgCartStnXRef(ilCIndex1, ilCIndex2).iAdfCode
                                                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    Else
                                                        ilRet = BTRV_ERR_NONE
                                                    End If
                                                    If rbcInterface(0).Value Then
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            slStr = slStr & "," & """" & Trim$(tmAdf.sName) & """" & "," & "reserved,reserved"
                                                        Else
                                                            slStr = slStr & "," & "reserved,reserved,reserved"
                                                        End If
                                                    Else
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            slStr = slStr & "," & """" & Trim$(tmAdf.sName) & """"
                                                        Else
                                                            slStr = slStr & ","
                                                        End If
                                                        slStr = slStr & "," & Trim$(str$(tgCartStnXRef(ilCIndex1, ilCIndex2).iLen)) & "," & Trim$(tmStnInfo(ilLoop).sKCNo) & "," & Format$(lmInputStartDate, "mm/dd/yy")    'Trim$(smWeekNo)
                                                    End If
                                                    Print #hlEnv, slStr
                                                End If
                                        '        Exit For
                                            End If
                                        'Next ilVeh
                                        ilIndex1 = ilCIndex1
                                        ilCIndex1 = tgCartStnXRef(ilIndex1, ilCIndex2).iLkCartInfo1
                                        ilCIndex2 = tgCartStnXRef(ilIndex1, ilCIndex2).iLkCartInfo2
                                    Loop
                                End If
                            End If
                        Next ilGen
                    End If
                End If
            Next ilIndex
            'Print #hlEnv, "ADDR: " & Trim$(tmStnInfo(ilLoop).sTransportal) & ", " & Trim$(tmStnInfo(ilLoop).sCallLetter) & "-" & Trim$(tmStnInfo(ilLoop).sBand)
            If rbcInterface(0).Value Then
                For ilTran = LBound(tmStnInfo(ilLoop).sTransportal) To UBound(tmStnInfo(ilLoop).sTransportal) Step 1
                    If Trim$(tmStnInfo(ilLoop).sTransportal(ilTran)) <> "" Then
                        Print #hlEnv, "ADDR: " & Trim$(tmStnInfo(ilLoop).sTransportal(ilTran)) & "," & """" & Trim$(tmStnInfo(ilLoop).sCallLetter) & "-" & Trim$(tmStnInfo(ilLoop).sBand) & """"
                    End If
                Next ilTran
            End If
            Close #hlEnv
        End If
    Next ilLoop
    Exit Function
'mCreateEnvFileErr:
'    ilRet = 1
'    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateCmmlLogFile                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create envelop file            *
'*                                                     *
'*                      Structure                      *
'*                      SCHEDULE                       *
'*                      EVENT: Date Time, Window, Relay, Carts *
'*                      ADDR: EDAS serail #            *
'*                                                     *
'*******************************************************
Private Function mCreateCmmlLogFile(ilVefCode As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilL1Count                     ilL2Count                     ilVeh                     *
'*                                                                                        *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim ilVIndex As Integer
    Dim slExportFile As String
    Dim slPDFFileName As String
    Dim ilRet As Integer
    Dim hlSch As Integer
    Dim ilInclude As Integer
    Dim slDate As String
    Dim slTime As String
    Dim llTime As Long
    Dim llRunTime As Long
    Dim slSpotDate As String
    Dim llSpotTime As Long
    Dim slSeqNo As String
    Dim ilUpper As Integer
    Dim ilSpot As Integer
    Dim ilIndex As Integer
    Dim ilSIndex As Integer
    Dim ilEIndex As Integer
    Dim ilFound As Integer
    Dim ilLen As Integer
    Dim slCartNo As String
    Dim slRecord As String
    Dim slPDFRecord As String
    Dim slDateTime As String
    Dim slShortTitle As String
    Dim slTShortTitle As String
    Dim slISCI As String
    Dim slAirPlays As String
    Dim llRunDate As Long
    Dim llTstDate As Long
    Dim llGenTime As Long
    Dim ilRetCopy As Integer
    Dim slOutput As String
    Dim ilPledge As Integer
    Dim slStr As String
    Dim ilRecordType As Integer
    Dim slSTime As String
    Dim slETime As String
    Dim ilRdf As Integer
    Dim ilTest As Integer
    Dim ilShowMGTime As Integer
    Dim slRotStartDate As String
    Dim slRotEndDate As String
    Dim slRotComment As String
    Dim slTimeRestrictions As String
    Dim slDayRestrictions As String
    '5/31/06:
    Dim llPrevDate As Long
    '5/31/06
    '6/30/06: Handle long comments
    Dim llCsfCode As Long
    '6/30/06: End of Change

    lacProcessing.Caption = "Generating Commercial Log File"

    mCreateCmmlLogFile = True
    For ilLoop = 0 To UBound(tmStnInfo) - 1 Step 1
        If (tmStnInfo(ilLoop).sType = "S") And (tmStnInfo(ilLoop).sCmmlLogReq = "L") Then
            ilInclude = False
            If ilVefCode = tmStnInfo(ilLoop).iAirVeh Then
                ilInclude = True
            Else
                ilVIndex = tmStnInfo(ilLoop).iLkStnInfo
                Do While ilVIndex >= 0
                    If tmStnInfo(ilVIndex).iAirVeh = ilVefCode Then
                        ilInclude = True
                        Exit Do
                    End If
                    ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
                Loop
            End If
            If ilInclude Then
                ilRet = gBinarySearchVef(ilVefCode)
                ilRet = 0
                'On Error GoTo mCreateCmmlLogFileErr:
                slExportFile = sgExportPath & Trim$(tmStnInfo(ilLoop).sFileName) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(smVehName) & "_Commercial_Log(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Csv"
                slPDFFileName = Trim$(tmStnInfo(ilLoop).sFileName) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(smVehName) & "_Commercial_Log(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Pdf"
                slStr = Trim$(tmStnInfo(ilLoop).sCallLetter) & "-" & Trim$(tmStnInfo(ilLoop).sBand) ' & "," & "Commercial Log: " & smVehName & "," & "Week " & smWeekNo & " " & Format(lmInputStartDate, "mm-dd-yyyy") & "," & "(Run " & smRunLetter & " " & Format(smTranDate, "mm-dd-yyyy") & ")"
                Do While Len(slStr) < 7
                    slStr = slStr & " "
                Loop
                slPDFRecord = slStr
                slStr = "Commercial Log: " & smVehName
                Do While Len(slStr) < 60
                    slStr = slStr & " "
                Loop
                slPDFRecord = slPDFRecord & slStr
                slStr = "Week " & smWeekNo & " " & Format(lmInputStartDate, "mm-dd-yyyy")
                Do While Len(slStr) < 20
                    slStr = slStr & " "
                Loop
                slPDFRecord = slPDFRecord & slStr
                slStr = "(Run " & smRunLetter & " " & Format(smTranDate, "mm-dd-yyyy") & ")"
                Do While Len(slStr) < 20
                    slStr = slStr & " "
                Loop
                slPDFRecord = slPDFRecord & slStr
                If ckcCmmlLog(1).Value = vbChecked Then
                    'slDateTime = FileDateTime(slExportFile)
                    ilRet = gFileExist(slExportFile)
                    If ilRet = 0 Then
                        Kill slExportFile
                    End If
                    ilRet = 0
                    'On Error GoTo mCreateCmmlLogFileErr:
                    'hlSch = FreeFile
                    'Open slExportFile For Output As hlSch
                    ilRet = gFileOpen(slExportFile, "Output", hlSch)
                    If ilRet <> 0 Then
                        'Print #hmMsg, "** Error opening Commercial Log file: " & slExportFile & " Error" & str(err)
                        gAutomationAlertAndLogHandler "** Error opening Commercial Log file: " & slExportFile & " Error" & str(err)
                        mCreateCmmlLogFile = False
                        Exit Function
                    End If
                    Print #hlSch, slStr
                    '5/31/06:  Output title in csv file
                    slStr = "Date" & "," & "Short Title" & "," & "Copy Start" & "," & "Copy End" & "," & "ISCI" & "," & "Len" & "," & "# Air" & "," & "Time"
                    Print #hlSch, slStr
                    slStr = "," & "," & "Date" & "," & "Date" & "," & "," & "," & "Plays" & ","
                    Print #hlSch, slStr
                    '5/31/06
                End If
                If ckcCmmlLog(0).Value = vbChecked Then
                    smPDFDate = Format$(gNow(), "m/d/yy")
                    gPackDate smPDFDate, imPDFDate(0), imPDFDate(1)
                    smPDFTime = Format$(gNow(), "h:mm:ssAM/PM")
                    gPackTime smPDFTime, imPDFTime(0), imPDFTime(1)
                    lmPDFSeqNo = 0
                    hmTxr = CBtrvTable(TEMPHANDLE) 'CBtrvObj
                    ilRet = btrOpen(hmTxr, "", sgDBPath & "Txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                    If ilRet <> BTRV_ERR_NONE Then
                        'Print #hmMsg, "Export Terminated as TXR.BTR could not be Opened, error #" & str$(ilRet)
                        gAutomationAlertAndLogHandler "Export Terminated as TXR.BTR could not be Opened, error #" & str$(ilRet)
                        mCreateCmmlLogFile = False
                        Exit Function
                    End If
                    imTxrRecLen = Len(tmTxr)
                    If imTxrRecLen <> btrRecordLength(hmTxr) Then
                        'Print #hmMsg, "Export Terminated as TXR.BTR size is not matching, Internal" & str$(imTxrRecLen) & "vs External" & str$(btrRecordLength(hmTxr))
                        gAutomationAlertAndLogHandler "Export Terminated as TXR.BTR size is not matching, Internal" & str$(imTxrRecLen) & "vs External" & str$(btrRecordLength(hmTxr))
                        mCreateCmmlLogFile = False
                        Exit Function
                    End If
                    gPackDate smPDFDate, tmTxr.iGenDate(0), tmTxr.iGenDate(1)
                    tmTxr.lGenTime = gTimeToLong(smPDFTime, False)
                    lmPDFSeqNo = lmPDFSeqNo + 1
                    tmTxr.lSeqNo = lmPDFSeqNo
                    tmTxr.iType = 0
                    tmTxr.sText = slPDFRecord
                    tmTxr.lCsfCode = 0
                    ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
                    If ilRet <> BTRV_ERR_NONE Then
                        'Print #hmMsg, "Insert TXR Failed" & str$(ilRet) & " processing terminated"
                        gAutomationAlertAndLogHandler "Insert TXR Failed" & str$(ilRet) & " processing terminated"
                        If ckcCmmlLog(1).Value = vbChecked Then
                            Close #hlSch
                        End If
                        mClearTxr
                        btrDestroy hmTxr
                        mCreateCmmlLogFile = False
                        Exit Function
                    End If
                End If
                ReDim tgSchSpotInfo(0 To 0) As SCHSPOTINFO
                'Build array of spot for station
                tmCprSrchKey.iGenDate(0) = imGenDate(0)
                tmCprSrchKey.iGenDate(1) = imGenDate(1)
                gUnpackTimeLong imGenTime(0), imGenTime(1), False, llGenTime
                tmCprSrchKey.lGenTime = llGenTime
                ilRet = btrGetGreaterOrEqual(hmCpr, tmCpr, imCprRecLen, tmCprSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmCpr.iGenDate(0) = imGenDate(0)) And (tmCpr.iGenDate(1) = imGenDate(1)) And (tmCpr.lGenTime = llGenTime)
                    ilInclude = True
                    If StrComp(Trim$(tmStnInfo(ilLoop).sFdZone), Trim$(tmCpr.sZone), 1) <> 0 Then
                        If Trim$(tmCpr.sZone) <> "" Then
                            ilInclude = False
                        End If
                    End If
                    If ilInclude Then
                        ilInclude = False
                        If tmCpr.iVefCode = tmStnInfo(ilLoop).iAirVeh Then
                            ilInclude = True
                        Else
                            ilVIndex = tmStnInfo(ilLoop).iLkStnInfo
                            Do While ilVIndex >= 0
                                If tmStnInfo(ilVIndex).iAirVeh = tmCpr.iVefCode Then
                                    ilInclude = True
                                    Exit Do
                                End If
                                ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
                            Loop
                        End If
                    End If
                    If ilInclude Then
                        'DL: 5/3/05 Translate local time to EST transmit time
                        gUnpackDate tmCpr.iSpotDate(0), tmCpr.iSpotDate(1), slSpotDate
                        gUnpackTimeLong tmCpr.iSpotTime(0), tmCpr.iSpotTime(1), False, llTime
                        'Translate time based on zone
                        Select Case UCase$(Trim$(tmStnInfo(ilLoop).sFdZone))
                            Case "EST"
                                llSpotTime = llTime
                            Case "CST"
                                llSpotTime = llTime + 3600
                            Case "MST"
                                llSpotTime = llTime + 2 * 3600
                            Case "PST"
                                llSpotTime = llTime + 3 * 3600
                            Case Else
                                llSpotTime = llTime
                        End Select
                        If (llSpotTime >= 24 * CLng(3600)) Then
                            'Adjust date
                            If gWeekDayStr(slSpotDate) = 6 Then
                                slSpotDate = gObtainPrevMonday(slSpotDate)
                            Else
                                slSpotDate = gIncOneDay(slSpotDate)
                            End If
                            llSpotTime = llSpotTime - 24 * CLng(3600)
                            gPackDate slSpotDate, tmCpr.iSpotDate(0), tmCpr.iSpotDate(1)
                            gPackTimeLong llSpotTime, tmCpr.iSpotTime(0), tmCpr.iSpotTime(1)
                        Else
                            gPackTimeLong llSpotTime, tmCpr.iSpotTime(0), tmCpr.iSpotTime(1)
                        End If
                        gUnpackDateForSort tmCpr.iSpotDate(0), tmCpr.iSpotDate(1), slDate
                        gUnpackTimeLong tmCpr.iSpotTime(0), tmCpr.iSpotTime(1), False, llTime

                        slTime = Trim$(str$(llTime))
                        Do While Len(slTime) < 6
                            slTime = "0" & slTime
                        Loop
                        slSeqNo = Trim$(str$(tmCpr.iLineNo))
                        Do While Len(slSeqNo) < 5
                            slSeqNo = "0" & slSeqNo
                        Loop
                        tgSchSpotInfo(UBound(tgSchSpotInfo)).sKey = slDate & slTime & slSeqNo
                        tgSchSpotInfo(UBound(tgSchSpotInfo)).tCpr = tmCpr
                        ReDim Preserve tgSchSpotInfo(0 To UBound(tgSchSpotInfo) + 1) As SCHSPOTINFO
                    End If
                    ilRet = btrGetNext(hmCpr, tmCpr, imCprRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                'Sort spots
                ilUpper = UBound(tgSchSpotInfo)
                If ilUpper > 0 Then
                    ArraySortTyp fnAV(tgSchSpotInfo(), 0), ilUpper, 0, LenB(tgSchSpotInfo(0)), 0, LenB(tgSchSpotInfo(0).sKey), 0
                End If
                ilEIndex = -1
                'Output spots in same avail and adjacant avails if spot has regional copy
                '5/31/06:  Output first occurrance of the comment
                ReDim smRotComment(0 To 0) As String
                ReDim smTimeRestrictions(0 To 0) As String
                ReDim smDayRestrictions(0 To 0) As String
                ReDim tmDuplComment(0 To 0) As DUPLCOMMENT
                llPrevDate = 0
                '5/31/06
                ilSpot = 0
                Do While ilSpot < UBound(tgSchSpotInfo)
                    ilSIndex = ilSpot
                    ilEIndex = ilSIndex
                    tmCpr = tgSchSpotInfo(ilSpot).tCpr
                    gUnpackDateLong tmCpr.iSpotDate(0), tmCpr.iSpotDate(1), llRunDate
                    gUnpackTimeLong tmCpr.iSpotTime(0), tmCpr.iSpotTime(1), False, llRunTime
                    llRunTime = llRunTime + tmCpr.iLen
                    For ilIndex = ilSIndex + 1 To UBound(tgSchSpotInfo) - 1 Step 1
                        gUnpackDateLong tgSchSpotInfo(ilIndex).tCpr.iSpotDate(0), tgSchSpotInfo(ilIndex).tCpr.iSpotDate(1), llTstDate
                        gUnpackTimeLong tgSchSpotInfo(ilIndex).tCpr.iSpotTime(0), tgSchSpotInfo(ilIndex).tCpr.iSpotTime(1), False, llTime
                        'If llTime <= llRunTime Then
                        If (llTime <= llRunTime) And (llRunDate = llTstDate) Then
                            ilEIndex = ilIndex
                            llRunTime = llRunTime + tgSchSpotInfo(ilIndex).tCpr.iLen
                        Else
                            Exit For
                        End If
                    Next ilIndex
                    tmCpr = tgSchSpotInfo(ilSIndex).tCpr
                    gUnpackDate tmCpr.iSpotDate(0), tmCpr.iSpotDate(1), slDate
                    slDate = gAdjYear(slDate)
                    gUnpackTime tmCpr.iSpotTime(0), tmCpr.iSpotTime(1), "A", "1", slTime
                    '5/31/06:
                    If ckcCmmlLog(1).Value = vbChecked Then
                        If (llPrevDate <> gDateValue(slDate)) And (llPrevDate > 0) Then
                            Print #hlSch, ","   'Force a blank line
                        End If
                        llPrevDate = gDateValue(slDate)
                    End If
                    '5/31/06
                    'ABC request 1/17/06:  Output date in mm-dd-yyyy
                    'slDate = Format$(slDate, "yyyy-mm-dd")
                    slDate = Format$(slDate, "mm-dd-yyyy")
                    slTime = Format$(slTime, "hh:mm:ss")
                    For ilIndex = ilSIndex To ilEIndex Step 1
                        ilLen = ilLen + tgSchSpotInfo(ilIndex).tCpr.iLen
                    Next ilIndex
                    'Output spots
                    For ilIndex = ilSIndex To ilEIndex Step 1
                        slRecord = slDate & ","
                        tmSdfSrchKey3.lCode = tgSchSpotInfo(ilIndex).tCpr.lCntrNo
                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            tmChfSrchKey.lCode = tmSdf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            If ilRet = BTRV_ERR_NONE Then
                                If tmAdf.iCode <> tmChf.iAdfCode Then
                                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                Else
                                    ilRet = BTRV_ERR_NONE
                                End If
                                If ilRet = BTRV_ERR_NONE Then
                                    slShortTitle = gFileNameFilter(Trim$(gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)))
                                    ilFound = False
                                    tmRsfSrchKey1.lCode = tmSdf.lCode
                                    ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                    tmRsfSrchKey1.lCode = tmSdf.lCode
                                    ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
                                        If tmRsf.sType = "B" Then
                                            If tmRsf.iBVefCode = tmStnInfo(ilLoop).iAirVeh Then
                                                ilFound = True
                                            Else
                                                ilVIndex = tmStnInfo(ilLoop).iLkStnInfo
                                                Do While ilVIndex >= 0
                                                    If tmRsf.iBVefCode = tmStnInfo(ilVIndex).iAirVeh Then
                                                        ilFound = True
                                                        Exit Do
                                                    End If
                                                    ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
                                                Loop
                                            End If
                                            If ilFound Then
                                                Exit Do
                                            End If
                                        End If
                                        ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                    Loop
                                    If ilFound Then
                                        'Get regional copy
                                        tmSdf.sPtType = tmRsf.sPtType
                                        tmSdf.lCopyCode = tmRsf.lCopyCode
                                        ilRetCopy = mObtainCopy("", slTShortTitle)
                                        mGetRotInfo Trim$(tmCpf.sISCI), slRotStartDate, slRotEndDate, slRotComment, slTimeRestrictions, slDayRestrictions, llCsfCode
                                    Else
                                        'Get copy
                                        ilRetCopy = mObtainCopy(tmStnInfo(ilLoop).sFdZone, slTShortTitle)
                                        mGetRotInfo Trim$(tmCpf.sISCI), slRotStartDate, slRotEndDate, slRotComment, slTimeRestrictions, slDayRestrictions, llCsfCode
                                    End If
                                    If (ilFound) And (tgSpf.sUseProdSptScr = "P") Then
                                        tmBofSrchKey.lCode = tmRsf.lRBofCode
                                        ilRet = btrGetEqual(hmBof, tmBof, imBofRecLen, tmBofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        If ilRet = BTRV_ERR_NONE Then
                                            tmSifSrchKey.lCode = tmBof.lSifCode
                                            ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet = BTRV_ERR_NONE Then
                                                slShortTitle = gFileNameFilter(Trim$(tmSif.sName))
                                            Else
                                                slShortTitle = gFileNameFilter(slTShortTitle)
                                            End If
                                        Else
                                            slShortTitle = gFileNameFilter(slTShortTitle)
                                        End If
                                        If Trim$(slShortTitle) = "" Then
                                            tmChfSrchKey.lCode = tmRsf.lRChfCode
                                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet = BTRV_ERR_NONE Then
                                                tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet = BTRV_ERR_NONE Then
                                                    tmSdf.lChfCode = tmRsf.lRChfCode
                                                    tmSdf.iLineNo = 0
                                                    slShortTitle = gFileNameFilter(Trim$(gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)))
                                                End If
                                            End If
                                        End If
                                    End If
                                    If ilRetCopy Then
                                        If Trim$(tmCif.sCut) = "" Then
                                            slCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName)
                                        Else
                                            slCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut)
                                        End If
                                        slISCI = gFileNameFilter(Trim$(tmCpf.sISCI))
                                    Else
                                        slCartNo = ""
                                        slISCI = ""
                                    End If
                                    If tmStnInfo(ilLoop).sCmmlLogCart <> "C" Then
                                        slCartNo = ""
                                    End If
                                    'ABC Request 1/17/06:  Include air plays between len and dayparts
                                    slAirPlays = Trim$(str$(tmStnInfo(ilLoop).iAirPlays))
                                    'slRecord = slDate & "," & """" & slShortTitle & """" & "," & """" & slCartNo & """" & "," & """" & slISCI & """" & "," & Trim$(Str$(tmCpr.iLen))
                                    '3/9/06:  Replace cart # with Start and End Rotation Dates
                                    'slRecord = slDate & "," & """" & slShortTitle & """" & "," & """" & slCartNo & """" & "," & """" & slISCI & """" & "," & Trim$(Str$(tmCpr.iLen)) & "," & slAirPlays
                                    slRecord = slDate & "," & """" & slShortTitle & """" & "," & slRotStartDate & "," & slRotEndDate & "," & """" & slISCI & """" & "," & Trim$(str$(tmCpr.iLen)) & "," & slAirPlays
                                    slStr = slDate
                                    Do While Len(slStr) < 10
                                        slStr = slStr & " "
                                    Loop
                                    slPDFRecord = slStr
                                    slStr = slShortTitle
                                    Do While Len(slStr) < 15
                                        slStr = slStr & " "
                                    Loop
                                    slPDFRecord = slPDFRecord & slStr
                                    '3/9/06:  Replace Cart with Start and End Rotation Dates
                                    'slStr = slCartNo
                                    'Do While Len(slStr) < 10
                                    '    slStr = slStr & " "
                                    'Loop
                                    'slPDFRecord = slPDFRecord & slStr
                                    slStr = slRotStartDate
                                    Do While Len(slStr) < 10
                                        slStr = slStr & " "
                                    Loop
                                    slPDFRecord = slPDFRecord & slStr
                                    slStr = slRotEndDate
                                    Do While Len(slStr) < 10
                                        slStr = slStr & " "
                                    Loop
                                    slPDFRecord = slPDFRecord & slStr
                                    slStr = slISCI
                                    Do While Len(slStr) < 20
                                        slStr = slStr & " "
                                    Loop
                                    slPDFRecord = slPDFRecord & slStr
                                    slStr = Trim$(str$(tmCpr.iLen))
                                    Do While Len(slStr) < 5
                                        slStr = slStr & " "
                                    Loop
                                    slPDFRecord = slPDFRecord & slStr
                                    'ABC Request 1/17/06:  Include air plays between length and dayparts
                                    slStr = slAirPlays
                                    Do While Len(slStr) < 4
                                        slStr = " " & slStr
                                    Loop
                                    slPDFRecord = slPDFRecord & slStr
                                    'Append time
                                    ilRecordType = 1
                                    If tmStnInfo(ilLoop).sCmmlLogDPType = "S" Then
                                        'Get RDF
                                        tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                        tmClfSrchKey.iLine = tmSdf.iLineNo
                                        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                                        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                                            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                        Loop
                                        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                                            ilRdf = gBinarySearchRdf(tmClf.iRdfCode)
                                            If ilRdf <> -1 Then
                                                If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                                                    ilShowMGTime = True
                                                    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                                                    For ilTest = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                                        If (tgMRdf(ilRdf).iStartTime(0, ilTest) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilTest) <> 0) Then
                                                            gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilTest), tgMRdf(ilRdf).iStartTime(1, ilTest), "A", "1", slSTime
                                                            gUnpackTime tgMRdf(ilRdf).iEndTime(0, ilTest), tgMRdf(ilRdf).iEndTime(1, ilTest), "A", "1", slETime
                                                            If (llTime >= gTimeToLong(slSTime, False)) And (llTime <= gTimeToLong(slETime, True)) Then
                                                                ilShowMGTime = False
                                                            End If
                                                        End If
                                                    Next ilTest
                                                Else
                                                    ilShowMGTime = False
                                                End If
                                                If Not ilShowMGTime Then
                                                    For ilTest = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                                        If (tgMRdf(ilRdf).iStartTime(0, ilTest) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilTest) <> 0) Then
                                                            gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilTest), tgMRdf(ilRdf).iStartTime(1, ilTest), "A", "1", slSTime
                                                            gUnpackTime tgMRdf(ilRdf).iEndTime(0, ilTest), tgMRdf(ilRdf).iEndTime(1, ilTest), "A", "1", slETime
                                                            slStr = slRecord & "," & slSTime & "-" & slETime
                                                            If ckcCmmlLog(1).Value = vbChecked Then
                                                                Print #hlSch, slStr
                                                            End If
                                                            slStr = slSTime & "-" & slETime
                                                            Do While Len(slStr) < 30
                                                                slStr = slStr & " "
                                                            Loop
                                                            slPDFRecord = slPDFRecord & slStr
                                                            If ckcCmmlLog(0).Value = vbChecked Then
                                                                gPackDate smPDFDate, tmTxr.iGenDate(0), tmTxr.iGenDate(1)
                                                                tmTxr.lGenTime = gTimeToLong(smPDFTime, False)
                                                                lmPDFSeqNo = lmPDFSeqNo + 1
                                                                tmTxr.lSeqNo = lmPDFSeqNo
                                                                tmTxr.iType = ilRecordType
                                                                tmTxr.sText = slPDFRecord
                                                                tmTxr.lCsfCode = 0
                                                                ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
                                                                If ilRet <> BTRV_ERR_NONE Then
                                                                    'Print #hmMsg, "Insert TXR Failed" & str$(ilRet) & " processing terminated"
                                                                    gAutomationAlertAndLogHandler "Insert TXR Failed" & str$(ilRet) & " processing terminated"
                                                                    If ckcCmmlLog(1).Value = vbChecked Then
                                                                        Close #hlSch
                                                                    End If
                                                                    mClearTxr
                                                                    btrDestroy hmTxr
                                                                    mCreateCmmlLogFile = False
                                                                    Exit Function
                                                                End If
                                                            End If
                                                            'Include Air Plays
                                                            'slRecord = "" & "," & "" & "," & "" & "," & "" & "," & ""
                                                            'slPDFRecord = String(60, " ")
                                                            slRecord = "" & "," & "" & "," & "" & "," & "" & "," & "" & "," & "" & "," & ""
                                                            slPDFRecord = String(74, " ")
                                                            ilRecordType = 2
                                                        End If
                                                    Next ilTest
                                                Else
                                                    gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slSTime
                                                    slSTime = Format$(slSTime, "hAM/PM")
                                                    llTime = gTimeToLong(slSTime, False) + 3600
                                                    slETime = gFormatTimeLong(llTime, "A", "1")
                                                    slETime = Format$(slETime, "hAM/PM")
                                                    slStr = slRecord & "," & slSTime & "-" & slETime
                                                    If ckcCmmlLog(1).Value = vbChecked Then
                                                        Print #hlSch, slStr
                                                    End If
                                                    slStr = slSTime & "-" & slETime
                                                    Do While Len(slStr) < 30
                                                        slStr = slStr & " "
                                                    Loop
                                                    slPDFRecord = slPDFRecord & slStr
                                                    If ckcCmmlLog(0).Value = vbChecked Then
                                                        gPackDate smPDFDate, tmTxr.iGenDate(0), tmTxr.iGenDate(1)
                                                        tmTxr.lGenTime = gTimeToLong(smPDFTime, False)
                                                        lmPDFSeqNo = lmPDFSeqNo + 1
                                                        tmTxr.lSeqNo = lmPDFSeqNo
                                                        tmTxr.iType = ilRecordType
                                                        tmTxr.sText = slPDFRecord
                                                        tmTxr.lCsfCode = 0
                                                        ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            'Print #hmMsg, "Insert TXR Failed" & str$(ilRet) & " processing terminated"
                                                            gAutomationAlertAndLogHandler "Insert TXR Failed" & str$(ilRet) & " processing terminated"
                                                            If ckcCmmlLog(1).Value = vbChecked Then
                                                                Close #hlSch
                                                            End If
                                                            mClearTxr
                                                            btrDestroy hmTxr
                                                            mCreateCmmlLogFile = False
                                                            Exit Function
                                                        End If
                                                    End If
                                                End If
                                                If Not mWriteRestrictions(hlSch, slRotComment, llCsfCode) Then
                                                    mCreateCmmlLogFile = False
                                                    Exit Function
                                                End If
                                                If Not mWriteRestrictions(hlSch, slTimeRestrictions, 0) Then
                                                    mCreateCmmlLogFile = False
                                                    Exit Function
                                                End If
                                                If Not mWriteRestrictions(hlSch, slDayRestrictions, 0) Then
                                                    mCreateCmmlLogFile = False
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                    Else
                                        For ilPledge = LBound(tmStnInfo(ilLoop).sCmmlLogPledge) To UBound(tmStnInfo(ilLoop).sCmmlLogPledge) Step 1
                                            If Trim$(tmStnInfo(ilLoop).sCmmlLogPledge(ilPledge)) <> "" Then
                                                slStr = slRecord & "," & tmStnInfo(ilLoop).sCmmlLogPledge(ilPledge)
                                                If ckcCmmlLog(1).Value = vbChecked Then
                                                    Print #hlSch, slStr
                                                End If
                                                slStr = tmStnInfo(ilLoop).sCmmlLogPledge(ilPledge)
                                                Do While Len(slStr) < 30
                                                    slStr = slStr & " "
                                                Loop
                                                slPDFRecord = slPDFRecord & slStr
                                                If ckcCmmlLog(0).Value = vbChecked Then
                                                    gPackDate smPDFDate, tmTxr.iGenDate(0), tmTxr.iGenDate(1)
                                                    tmTxr.lGenTime = gTimeToLong(smPDFTime, False)
                                                    lmPDFSeqNo = lmPDFSeqNo + 1
                                                    tmTxr.lSeqNo = lmPDFSeqNo
                                                    tmTxr.iType = ilRecordType
                                                    tmTxr.sText = slPDFRecord
                                                    tmTxr.lCsfCode = 0
                                                    ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
                                                    If ilRet <> BTRV_ERR_NONE Then
                                                        'Print #hmMsg, "Insert TXR Failed" & str$(ilRet) & " processing terminated"
                                                        gAutomationAlertAndLogHandler "Insert TXR Failed" & str$(ilRet) & " processing terminated"
                                                        mClearTxr
                                                        If ckcCmmlLog(1).Value = vbChecked Then
                                                            Close #hlSch
                                                        End If
                                                        btrDestroy hmTxr
                                                        mCreateCmmlLogFile = False
                                                        Exit Function
                                                    End If
                                                End If
                                                'Include Air Plays
                                                'slRecord = "" & "," & "" & "," & "" & "," & "" & "," & ""
                                                'slPDFRecord = String(60, " ")
                                                slRecord = "" & "," & "" & "," & "" & "," & "" & "," & "" & "," & "" & "," & ""
                                                slPDFRecord = String(74, " ")
                                                ilRecordType = 2
                                            End If
                                        Next ilPledge
                                        If Not mWriteRestrictions(hlSch, slRotComment, llCsfCode) Then
                                            mCreateCmmlLogFile = False
                                            Exit Function
                                        End If
                                        If Not mWriteRestrictions(hlSch, slTimeRestrictions, 0) Then
                                            mCreateCmmlLogFile = False
                                            Exit Function
                                        End If
                                        If Not mWriteRestrictions(hlSch, slDayRestrictions, 0) Then
                                            mCreateCmmlLogFile = False
                                            Exit Function
                                        End If
                                    End If
                                Else
                                End If
                            Else
                            End If
                        Else
                        End If
                    Next ilIndex
                    ilSpot = ilEIndex + 1
                Loop
                If ckcCmmlLog(1).Value = vbChecked Then
                    Close #hlSch
                End If
                If ckcCmmlLog(0).Value = vbChecked Then
                    igRptCallType = STATIONFEEDJOB
                    igRptType = 2
                    slOutput = "2"
                    If (Not igStdAloneMode) And (imShowHelpMsg) Then
                        If igTestSystem Then
                            slStr = "ExpStnFd^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & smPDFDate & "\" & smPDFTime
                        Else
                            slStr = "ExpStnFd^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & smPDFDate & "\" & smPDFTime
                        End If
                    Else
                        If igTestSystem Then
                            slStr = "ExpStnFd^Test^NOHELP\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & smPDFDate & "\" & smPDFTime
                        Else
                            slStr = "ExpStnFd^Prod^NOHELP\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & smPDFDate & "\" & smPDFTime
                        End If
                    End If
                    sgCommandStr = slStr
                    RptSelTx.Show vbModal
                    mClearTxr
                    btrDestroy hmTxr
                End If
            End If
        End If
    Next ilLoop
    Exit Function
'mCreateCmmlLogFileErr:
'    ilRet = 1
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateSchFile                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create envelop file            *
'*                                                     *
'*                      Structure                      *
'*                      SCHEDULE                       *
'*                      EVENT: Date Time, Window, Relay, Carts *
'*                      ADDR: EDAS serail #            *
'*                                                     *
'*******************************************************
Private Function mCreateSchFile() As Integer
    Dim ilLoop As Integer
    Dim ilVIndex As Integer
    Dim slToFile As String
    Dim ilRet As Integer
    Dim hlSch As Integer
    Dim ilL1Count As Integer
    Dim ilL2Count As Integer
    Dim ilInclude As Integer
    Dim slChar1 As String * 1
    Dim slChar2 As String * 1
    Dim slChar3 As String * 1
    Dim slDate As String
    Dim slTime As String
    Dim llTime As Long
    Dim llRunTime As Long
    Dim slSpotDate As String
    Dim llSpotTime As Long
    Dim slSeqNo As String
    Dim ilUpper As Integer
    Dim ilSpot As Integer
    Dim ilIndex As Integer
    Dim ilSIndex As Integer
    Dim ilEIndex As Integer
    Dim ilFound As Integer
    Dim ilLen As Integer
    Dim slCartNo As String
    Dim slRecord As String
    Dim slDateTime As String
    Dim slShortTitle As String
    Dim slTShortTitle As String
    Dim slISCI As String
    Dim llRunDate As Long
    Dim llTstDate As Long
    Dim llGenTime As Long
    Dim ilVeh As Integer
    Dim ilRetCopy As Integer
    Dim ilEDAS As Integer

    lacProcessing.Caption = "Generating Schedule File"

    mCreateSchFile = True
    slChar1 = "A"
    slChar2 = "A"
    slChar3 = "A"
    ilL1Count = 1
    ilL2Count = 1
    For ilLoop = 0 To UBound(tmStnInfo) - 1 Step 1

        ilInclude = True
        If rbcGen(2).Value Then
            ilInclude = False
            For ilVeh = 0 To lbcRegVeh.ListCount - 1 Step 1
                If lbcRegVeh.Selected(ilVeh) Then
                    For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If StrComp(Trim$(tgMVef(ilIndex).sName), Trim$(lbcRegVeh.List(ilVeh)), 1) = 0 Then
                            If tgMVef(ilIndex).iCode = tmStnInfo(ilLoop).iAirVeh Then
                                ilInclude = True
                            Else
                                ilVIndex = tmStnInfo(ilLoop).iLkStnInfo
                                Do While ilVIndex >= 0
                                    If tmStnInfo(ilVIndex).iAirVeh = tgMVef(ilIndex).iCode Then
                                        ilInclude = True
                                        Exit Do
                                    End If
                                    ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
                                Loop
                            End If
                            If ilInclude Then
                                Exit For
                            End If
                        End If
                    Next ilIndex
                End If
            Next ilVeh
        End If

        'If (tmStnInfo(ilLoop).sType = "S") And (tmStnInfo(ilLoop).lRafCode > 0) Then 'Ignore linked records- handled within
        If ((tmStnInfo(ilLoop).sType = "S") And (tmStnInfo(ilLoop).lRafCode > 0)) Or (rbcGen(2).Value And (tmStnInfo(ilLoop).sType = "S") And ilInclude) Then  'Ignore linked records- handled within
            ilRet = 0
            'On Error GoTo mCreateSchFileErr:
            'slToFile = sgExportPath & "Env" & Trim$(tmStnInfo(ilLoop).sSiteID) & smFeedNo & ".txt"
            'slToFile = sgExportPath & "Sch" & slChar3 & slChar2 & slChar1 & Right$(smFeedNo, 2) & ".txt"
            slToFile = sgExportPath & Trim$(tmStnInfo(ilLoop).sFileName) & ".sch"
            ''ilL1Count = ilL1Count + 1
            ''If ilL1Count >= 27 Then
            ''    slChar1 = "A"
            ''    ilL1Count = 1
            ''    ilL2Count = ilL2Count + 1
            ''    If ilL2Count >= 27 Then
            ''        slChar2 = "A"
            ''        ilL2Count = 1
            ''        slChar3 = Chr$(Asc(slChar3) + 1)
            ''    Else
            ''        slChar2 = Chr$(Asc(slChar2) + 1)
            ''    End If
            ''Else
            ''    slChar1 = Chr$(Asc(slChar1) + 1)
            ''End If
            'slDateTime = FileDateTime(slToFile)
            ilRet = gFileExist(slToFile)
            If ilRet = 0 Then
                Kill slToFile
            End If
            ilRet = 0
            'On Error GoTo mCreateSchFileErr:
            'hlSch = FreeFile
            'Open slToFile For Output As hlSch
            ilRet = gFileOpen(slToFile, "Output", hlSch)
            If ilRet <> 0 Then
                'Print #hmMsg, "** Error opening Schedule file: " & slToFile & " Error" & str(err)
                gAutomationAlertAndLogHandler "** Error opening Schedule file: " & slToFile & " Error" & str(err)
                mCreateSchFile = False
                Exit Function
            End If
            Print #hlSch, "SCHEDULE"
            ReDim tgSchSpotInfo(0 To 0) As SCHSPOTINFO
            'Build array of spot for station
            tmCprSrchKey.iGenDate(0) = imGenDate(0)
            tmCprSrchKey.iGenDate(1) = imGenDate(1)
            'tmCprSrchKey.iGenTime(0) = imGenTime(0)
            'tmCprSrchKey.iGenTime(1) = imGenTime(1)
            gUnpackTimeLong imGenTime(0), imGenTime(1), False, llGenTime
            tmCprSrchKey.lGenTime = llGenTime
            ilRet = btrGetGreaterOrEqual(hmCpr, tmCpr, imCprRecLen, tmCprSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmCpr.iGenDate(0) = imGenDate(0)) And (tmCpr.iGenDate(1) = imGenDate(1)) And (tmCpr.lGenTime = llGenTime)
                ilInclude = True
                'If StrComp(Trim$(tmStnInfo(ilLoop).sFdZone), Trim$(tmCpr.sZone), 1) <> 0 Then
                'Later- add a site option: 'Transmit all spots as if in EST zone'
                'changed 3/19/01- abc request
                'If StrComp(Trim$("EST"), Trim$(tmCpr.sZone), 1) <> 0 Then
                If StrComp(Trim$(tmStnInfo(ilLoop).sFdZone), Trim$(tmCpr.sZone), 1) <> 0 Then
                    If Trim$(tmCpr.sZone) <> "" Then
                        ilInclude = False
                    End If
                End If
                If ilInclude Then
                    ilInclude = False
                    If tmCpr.iVefCode = tmStnInfo(ilLoop).iAirVeh Then
                        ilInclude = True
                    Else
                        ilVIndex = tmStnInfo(ilLoop).iLkStnInfo
                        Do While ilVIndex >= 0
                            If tmStnInfo(ilVIndex).iAirVeh = tmCpr.iVefCode Then
                                ilInclude = True
                                Exit Do
                            End If
                            ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
                        Loop
                    End If
                End If
                If ilInclude Then
                    'DL: 5/3/05 Translate local time to EST transmit time
                    gUnpackDate tmCpr.iSpotDate(0), tmCpr.iSpotDate(1), slSpotDate
                    gUnpackTimeLong tmCpr.iSpotTime(0), tmCpr.iSpotTime(1), False, llTime
                    'Translate time based on zone
                    Select Case UCase$(Trim$(tmStnInfo(ilLoop).sFdZone))
                        Case "EST"
                            llSpotTime = llTime
                        Case "CST"
                            llSpotTime = llTime + 3600
                        Case "MST"
                            llSpotTime = llTime + 2 * 3600
                        Case "PST"
                            llSpotTime = llTime + 3 * 3600
                        Case Else
                            llSpotTime = llTime
                    End Select
                    If (llSpotTime >= 24 * CLng(3600)) Then
                        'Adjust date
                        If gWeekDayStr(slSpotDate) = 6 Then
                            'slSpotDate = gObtainPrevMonday(slSpotDate)
                            'Delivery link local and affiliate times alway match i.e. 9p local is 9p pst
                            'Therefore date is not adjust in expspot.  move date forward here
                            slSpotDate = gIncOneDay(slSpotDate)
                        Else
                            slSpotDate = gIncOneDay(slSpotDate)
                        End If
                        llSpotTime = llSpotTime - 24 * CLng(3600)
                        gPackDate slSpotDate, tmCpr.iSpotDate(0), tmCpr.iSpotDate(1)
                        gPackTimeLong llSpotTime, tmCpr.iSpotTime(0), tmCpr.iSpotTime(1)
                    Else
                        gPackTimeLong llSpotTime, tmCpr.iSpotTime(0), tmCpr.iSpotTime(1)
                    End If
                    gUnpackDateForSort tmCpr.iSpotDate(0), tmCpr.iSpotDate(1), slDate
                    gUnpackTimeLong tmCpr.iSpotTime(0), tmCpr.iSpotTime(1), False, llTime

                    slTime = Trim$(str$(llTime))
                    Do While Len(slTime) < 6
                        slTime = "0" & slTime
                    Loop
                    slSeqNo = Trim$(str$(tmCpr.iLineNo))
                    Do While Len(slSeqNo) < 5
                        slSeqNo = "0" & slSeqNo
                    Loop
                    tgSchSpotInfo(UBound(tgSchSpotInfo)).sKey = slDate & slTime & slSeqNo
                    tgSchSpotInfo(UBound(tgSchSpotInfo)).tCpr = tmCpr
                    ReDim Preserve tgSchSpotInfo(0 To UBound(tgSchSpotInfo) + 1) As SCHSPOTINFO
                End If
                ilRet = btrGetNext(hmCpr, tmCpr, imCprRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            'Sort spots
            ilUpper = UBound(tgSchSpotInfo)
            If ilUpper > 0 Then
                ArraySortTyp fnAV(tgSchSpotInfo(), 0), ilUpper, 0, LenB(tgSchSpotInfo(0)), 0, LenB(tgSchSpotInfo(0).sKey), 0
            End If
            ilEIndex = -1
            'Output spots in same avail and adjacant avails if spot has regional copy
            ilSpot = 0
            Do While ilSpot < UBound(tgSchSpotInfo)
                ilSIndex = ilSpot
                ilEIndex = ilSIndex
                tmCpr = tgSchSpotInfo(ilSpot).tCpr
                gUnpackDateLong tmCpr.iSpotDate(0), tmCpr.iSpotDate(1), llRunDate
                gUnpackTimeLong tmCpr.iSpotTime(0), tmCpr.iSpotTime(1), False, llRunTime
                llRunTime = llRunTime + tmCpr.iLen
                For ilIndex = ilSIndex + 1 To UBound(tgSchSpotInfo) - 1 Step 1
                    gUnpackDateLong tgSchSpotInfo(ilIndex).tCpr.iSpotDate(0), tgSchSpotInfo(ilIndex).tCpr.iSpotDate(1), llTstDate
                    gUnpackTimeLong tgSchSpotInfo(ilIndex).tCpr.iSpotTime(0), tgSchSpotInfo(ilIndex).tCpr.iSpotTime(1), False, llTime
                    'If llTime <= llRunTime Then
                    If (llTime <= llRunTime) And (llRunDate = llTstDate) Then
                        ilEIndex = ilIndex
                        llRunTime = llRunTime + tgSchSpotInfo(ilIndex).tCpr.iLen
                    Else
                        Exit For
                    End If
                Next ilIndex
                'Test if any spot has regional copy
                If rbcGen(1).Value Then
                    ilFound = False
                    For ilIndex = ilSIndex To ilEIndex Step 1
                        tmRsfSrchKey1.lCode = tgSchSpotInfo(ilIndex).tCpr.lCntrNo
                        ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tgSchSpotInfo(ilIndex).tCpr.lCntrNo)
                            If tmRsf.lRafCode = tmStnInfo(ilLoop).lRafCode Then
                                ilFound = True
                            Else
                                ilVIndex = tmStnInfo(ilLoop).iLkStnInfo
                                Do While ilVIndex >= 0
                                    If tmRsf.lRafCode = tmStnInfo(ilVIndex).lRafCode Then
                                        ilFound = True
                                        Exit Do
                                    End If
                                    ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
                                Loop
                            End If
                            If ilFound Then
                                Exit For
                            End If
                            ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        Loop
                    Next ilIndex
                Else
                    ilFound = True
                End If

                If ilFound Then
                    tmCpr = tgSchSpotInfo(ilSIndex).tCpr
                    gUnpackDate tmCpr.iSpotDate(0), tmCpr.iSpotDate(1), slDate
                    slDate = gAdjYear(slDate)
                    gUnpackTime tmCpr.iSpotTime(0), tmCpr.iSpotTime(1), "A", "1", slTime
                    slDate = Format$(slDate, "yyyy-mm-dd")
                    slTime = Format$(slTime, "hh:mm:ss")
                    If slTime = "12M" Then
                        slTime = "00:00:00"
                    End If
                    slRecord = "EVENT: " & slDate & " " & slTime
                    'Window
                    '3/8/06-  Replace 400 with value from Vehicle option file.  This value is set in mExtSpot
                    'slRecord = slRecord & "," & "400"
                    slRecord = slRecord & "," & Trim$(str$(tmCpr.lHd1CefCode))
                    'Relay #
                    ilLen = 0
                    For ilIndex = ilSIndex To ilEIndex Step 1
                        ilLen = ilLen + tgSchSpotInfo(ilIndex).tCpr.iLen
                    Next ilIndex
                    'If ilLen < 60 Then
                    '    slRecord = slRecord & ", 02"
                    'ElseIf ilLen > 60 Then
                    '    slRecord = slRecord & ", 03"
                    'Else
                    '    slRecord = slRecord & ", 01"
                    'End If
                    If ilLen = 30 Then
                        slRecord = slRecord & ",0004"
                    ElseIf ilLen = 60 Then
                        slRecord = slRecord & ",0002"
                    Else
                        slRecord = slRecord & ",0001"
                    End If
                    'Output spots
                    For ilIndex = ilSIndex To ilEIndex Step 1
                        tmSdfSrchKey3.lCode = tgSchSpotInfo(ilIndex).tCpr.lCntrNo
                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            tmChfSrchKey.lCode = tmSdf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            If ilRet = BTRV_ERR_NONE Then
                                If tmAdf.iCode <> tmChf.iAdfCode Then
                                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                Else
                                    ilRet = BTRV_ERR_NONE
                                End If
                                If ilRet = BTRV_ERR_NONE Then
                                    slShortTitle = gFileNameFilter(Trim$(gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)))
                                    ilFound = False
                                    If Not rbcGen(2).Value Then
                                        tmRsfSrchKey1.lCode = tmSdf.lCode
                                        ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
                                            If (tmRsf.sType <> "B") And (tmRsf.sType <> "A") Then
                                                If tmRsf.lRafCode = tmStnInfo(ilLoop).lRafCode Then
                                                    ilFound = True
                                                Else
                                                    ilVIndex = tmStnInfo(ilLoop).iLkStnInfo
                                                    Do While ilVIndex >= 0
                                                        If tmRsf.lRafCode = tmStnInfo(ilVIndex).lRafCode Then
                                                            ilFound = True
                                                            Exit Do
                                                        End If
                                                        ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
                                                    Loop
                                                End If
                                                If ilFound Then
                                                    Exit Do
                                                End If
                                            End If
                                            ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                        Loop
                                    Else
                                        tmRsfSrchKey1.lCode = tmSdf.lCode
                                        ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
                                            If tmRsf.sType = "B" Then
                                                If tmRsf.iBVefCode = tmStnInfo(ilLoop).iAirVeh Then
                                                    ilFound = True
                                                Else
                                                    ilVIndex = tmStnInfo(ilLoop).iLkStnInfo
                                                    Do While ilVIndex >= 0
                                                        If tmRsf.iBVefCode = tmStnInfo(ilVIndex).iAirVeh Then
                                                            ilFound = True
                                                            Exit Do
                                                        End If
                                                        ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
                                                    Loop
                                                End If
                                                If ilFound Then
                                                    Exit Do
                                                End If
                                            End If
                                            ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                        Loop
                                    End If
                                    If ilFound Then
                                        'Get regional copy
                                        tmSdf.sPtType = tmRsf.sPtType
                                        tmSdf.lCopyCode = tmRsf.lCopyCode
                                        ilRetCopy = mObtainCopy("", slTShortTitle)
                                    Else
                                        'Get copy
                                        ilRetCopy = mObtainCopy(tmStnInfo(ilLoop).sFdZone, slTShortTitle)
                                    End If
                                    If (rbcGen(2).Value) And (ilFound) And (tgSpf.sUseProdSptScr = "P") Then
                                        tmBofSrchKey.lCode = tmRsf.lRBofCode
                                        ilRet = btrGetEqual(hmBof, tmBof, imBofRecLen, tmBofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        If ilRet = BTRV_ERR_NONE Then
                                            tmSifSrchKey.lCode = tmBof.lSifCode
                                            ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet = BTRV_ERR_NONE Then
                                                slShortTitle = gFileNameFilter(Trim$(tmSif.sName))
                                            Else
                                                slShortTitle = gFileNameFilter(slTShortTitle)
                                            End If
                                        Else
                                            slShortTitle = gFileNameFilter(slTShortTitle)
                                        End If
                                        If Trim$(slShortTitle) = "" Then
                                            tmChfSrchKey.lCode = tmRsf.lRChfCode
                                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet = BTRV_ERR_NONE Then
                                                tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet = BTRV_ERR_NONE Then
                                                    tmSdf.lChfCode = tmRsf.lRChfCode
                                                    tmSdf.iLineNo = 0
                                                    slShortTitle = gFileNameFilter(Trim$(gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)))
                                                End If
                                            End If
                                        End If
                                    End If
                                    'If ilRet Then
                                    '    If Trim$(tmCif.sCut) = "" Then
                                    '        slCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName)
                                    '    Else
                                    '        slCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut)
                                    '    End If
                                    'Else
                                    '    slCartNo = ""
                                    'End If
                                    'slRecord = slRecord & ", " & slCartNo
                                    If ilRetCopy Then
                                        If Trim$(tmCif.sCut) = "" Then
                                            slCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName)
                                        Else
                                            slCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut)
                                        End If
                                        slISCI = gFileNameFilter(Trim$(tmCpf.sISCI))
                                    Else
                                        slCartNo = ""
                                        slISCI = ""
                                    End If
                                    'slRecord = slRecord & "," & """" & slShortTitle & ";" & slISCI & ".mp2" & """"
                                    slRecord = slRecord & "," & """" & slShortTitle & "(" & slISCI & ")" & ".mp2" & """"
                                Else
                                End If
                            Else
                            End If
                        Else
                        End If
                    Next ilIndex
                    Print #hlSch, slRecord
                End If
                ilSpot = ilEIndex + 1
            Loop
            For ilEDAS = LBound(tmStnInfo(ilLoop).sEDAS) To UBound(tmStnInfo(ilLoop).sEDAS) Step 1
                If Trim$(tmStnInfo(ilLoop).sEDAS(ilEDAS)) <> "" Then
                    Print #hlSch, "ADDR: " & Trim$(tmStnInfo(ilLoop).sEDAS(ilEDAS))
                End If
            Next ilEDAS
            Close #hlSch
        End If
    Next ilLoop
    Exit Function
'mCreateSchFileErr:
'    ilRet = 1
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mDuplRotation                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if this is a duplicate*
'*                     rotations                       *
'*                                                     *
'*******************************************************
Private Function mDuplRotation(ilVpfIndex As Integer, llRecPos As Long) As Integer
'   ilDupl = mDuplRotation()
'   Where:
'
'       ilVpfIndex(O)- Vpf Index
'       ilDupl(O)- 1 if this rotation should be combined with a selected rotation
'                  2 if this rotation matches a selected rotation
'                  0 if this rotation does not match a selected rotation
'
'       tmCrf(I)- Rotation to be checked
'       tmChf(I)- Contract associated with tmCrf
'
'   Note: you can't have tmCrf be "Combined" or "Matched" with more then
'         one Crf (they would have been "combined" or "Matched" previously)
'
    Dim ilRet As Integer
    Dim ilCrf As Integer
    Dim tlCrf As CRF
    Dim ilVeh As Integer
    Dim ilVehIndex As Integer
    Dim ilDVeh As Integer
    Dim ilDVehIndex As Integer
    Dim ilDVpfIndex As Integer
    Dim ilMatch As Integer
    Dim ilDay As Integer
    Dim ilTest As Integer
    Dim ilUpper As Integer
    Dim llSifCode As Long
    Dim ilVsf As Integer
    Dim slStr As String
    Dim ilDuplIndex As Integer
    Dim ilCombIndex As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim ilVefCode As Integer
    Dim ilVpf As Integer
    Dim ilVIndex As Integer
    ReDim llCifCode(0 To 0) As Long

    'Restoring bypassing of duplictae test 2/27/04 because rotation were not
    'being sent to vehicles within same group with other vehicles

    'Remove that duplicate test request made by ABC on 2/5/02- done on 7/12/02
    'Re-instated on 9/17/03
    mDuplRotation = 0
    Exit Function
    'Retain code in case we need to re-add

    ilVehIndex = -1
    For ilVeh = 0 To UBound(tmVef) - 1 Step 1
        If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
            ilVehIndex = ilVeh
            Exit For
        End If
    Next ilVeh
    If ilVehIndex = -1 Then
        For ilVeh = 0 To UBound(tmVef) - 1 Step 1
            If (tmVef(ilVeh).sType = "A") Or (tmVef(ilVeh).sType = "C") Then
                ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                If ilVIndex >= 0 Then
                    ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                    Do While ilVIndex >= 0
                        If tmCrf.iVefCode = tmLkVehInfo(ilVIndex).iVefCode Then
                            ilVehIndex = ilVeh
                            Exit For
                        End If
                        ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                    Loop
                End If
            Else
                ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                If ilVIndex >= 0 Then
                    ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                    Do While ilVIndex >= 0
                        If tmCrf.iVefCode = tmSALink(ilVIndex).iVefCode Then
                            ilVehIndex = ilVeh
                            Exit For
                        End If
                        ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                    Loop
                End If
            End If
        Next ilVeh
    End If
    If ilVehIndex = -1 Then
        mDuplRotation = 0
        Exit Function
    End If
    ilVpfIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
    'If tmVef(ilVehIndex).sType <> "S" Then
    '    For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
    '        tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) = 0
    '    Next ilVeh
    '    tmVpfInfo(ilVpfIndex).tVpf.iGLink(LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink)) = tmVef(ilVehIndex).iCode
    'End If
    ReDim imVefCode(0 To 0) As Integer
    ilVefCode = tmVef(ilVehIndex).iCode
    If tmVef(ilVehIndex).sType = "S" Then
        'Find airing vehicle
        'ilSAGroupNo = tgVpf(gVpfFind(ExpStnFd, ilVefCode)).iSAGroupNo
        'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If (tgMVef(ilVeh).sType = "A") And (tgMVef(ilVeh).sState <> "D") Then
        '        If (ilSAGroupNo = tgVpf(gVpfFind(ExpStnFd, tgMVef(ilVeh).iCode)).iSAGroupNo) And (ilSAGroupNo <> 0) Then
        '            imVefCode(UBound(imVefCode)) = tgMVef(ilVeh).iCode
        '            ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
        '        End If
        '    End If
        'Next ilVeh
        If ilVpfIndex >= 0 Then
            ilVpf = tmVpfInfo(ilVpfIndex).iFirstSALink
            Do While ilVpf >= 0
                imVefCode(UBound(imVefCode)) = tmSALink(ilVpf).iVefCode
                ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
                ilVpf = tmSALink(ilVpf).iNextLkVehInfo
            Loop
        End If
    Else
        ReDim imVefCode(0 To 1) As Integer
        imVefCode(0) = ilVefCode
    End If
    ilDuplIndex = -1
    ilCombIndex = -1
    'Get instructions
    ilUpper = 0
    tmCnfSrchKey.lCrfCode = tmCrf.lCode
    tmCnfSrchKey.iInstrNo = 0
    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
        llCifCode(ilUpper) = tmCnf.lCifCode
        ilUpper = ilUpper + 1
        ReDim Preserve llCifCode(0 To ilUpper)
        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mDuplRotation = 0
    'For ilCrf = 0 To UBound(tgSortCrf) - 1 Step 1
    ilCrf = 0
    Do While ilCrf <= UBound(tgSortCrf) - 1
        tlCrf = tgSortCrf(ilCrf).tCrf
        If tlCrf.lCode <> tmCrf.lCode Then
            'Determine if sent to same vehicle
            ilMatch = True
            If tmAdf.iCode <> tmChf.iAdfCode Then
                tmAdfSrchKey.iCode = tmChf.iAdfCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
            llSifCode = 0
            If tmChf.lVefCode < 0 Then
                tmVsfSrchKey.lCode = -tmChf.lVefCode
                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Do While ilRet = BTRV_ERR_NONE
                    For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                        If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
                            If tmVsf.lFSComm(ilVsf) > 0 Then
                                llSifCode = tmVsf.lFSComm(ilVsf)
                            End If
                            Exit For
                        End If
                    Next ilVsf
                    If llSifCode <> 0 Then
                        Exit Do
                    End If
                    If tmVsf.lLkVsfCode <= 0 Then
                        Exit Do
                    End If
                    tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
            slStr = gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf)
            'If StrComp(Trim$(tmChf.sProduct), Trim$(tgSortCrf(ilCrf).sCntrProd), 1) <> 0 Then
            If StrComp(Trim$(slStr), Trim$(tgSortCrf(ilCrf).sCntrProd), 1) <> 0 Then
                ilMatch = False
            End If
            If tmCrf.iAdfCode <> tlCrf.iAdfCode Then
                ilMatch = False
            End If
            If tmCrf.sRotType <> tlCrf.sRotType Then
                ilMatch = False
            End If
            'Determine if Dates, times, days,... are the same
            If (tmCrf.iStartDate(0) <> tlCrf.iStartDate(0)) Or (tmCrf.iStartDate(1) <> tlCrf.iStartDate(1)) Then
                ilMatch = False
            End If
            If (tmCrf.iEndDate(0) <> tlCrf.iEndDate(0)) Or (tmCrf.iEndDate(1) <> tlCrf.iEndDate(1)) Then
                ilMatch = False
            End If
            If (tmCrf.iStartTime(0) <> tlCrf.iStartTime(0)) Or (tmCrf.iStartTime(1) <> tlCrf.iStartTime(1)) Then
                ilMatch = False
            End If
            If (tmCrf.iEndTime(0) <> tlCrf.iEndTime(0)) Or (tmCrf.iEndTime(1) <> tlCrf.iEndTime(1)) Then
                ilMatch = False
            End If
            For ilDay = 0 To 6 Step 1
                If (tmCrf.sDay(ilDay) <> tlCrf.sDay(ilDay)) Then
                    ilMatch = False
                    Exit For
                End If
            Next ilDay
            If tmCrf.sZone <> tlCrf.sZone Then
                ilMatch = False
            Else
                If Trim$(tmCrf.sZone) = "R" Then
                    If tmCrf.lRafCode <> tlCrf.lRafCode Then
                        ilMatch = False
                    End If
                End If
            End If
            If tmCrf.iLen <> tlCrf.iLen Then
                ilMatch = False
            End If
            If tmCrf.sInOut <> tlCrf.sInOut Then
                ilMatch = False
            End If
            If tmCrf.ianfCode <> tlCrf.ianfCode Then
                ilMatch = False
            End If
            If ilMatch Then
                If tmCrf.iVefCode <> tlCrf.iVefCode Then
                    For ilDVeh = 0 To UBound(tmVef) - 1 Step 1
                        If tmVef(ilDVeh).iCode = tlCrf.iVefCode Then
                            ilDVehIndex = ilDVeh
                            Exit For
                        End If
                    Next ilDVeh
                    ilDVpfIndex = mFindVpfIndex(tmVef(ilDVehIndex).iCode)
                    'If tmVef(ilDVehIndex).sType <> "S" Then
                    '    For ilDVeh = LBound(tmVpfInfo(ilDVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilDVpfIndex).tVpf.iGLink) Step 1
                    '        tmVpfInfo(ilDVpfIndex).tVpf.iGLink(ilDVeh) = 0
                    '    Next ilDVeh
                    '    tmVpfInfo(ilDVpfIndex).tVpf.iGLink(LBound(tmVpfInfo(ilDVpfIndex).tVpf.iGLink)) = tmVef(ilDVehIndex).iCode
                    'End If
                    ReDim imDVefCode(0 To 0) As Integer
                    If tmVef(ilDVehIndex).sType = "S" Then
                        'Find airing vehicle
                        'ilSAGroupNo = tgVpf(gVpfFind(ExpStnFd, tmVef(ilDVehIndex).iCode)).iSAGroupNo
                        'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If (tgMVef(ilVeh).sType = "A") And (tgMVef(ilVeh).sState <> "D") Then
                        '        If (ilSAGroupNo = tgVpf(gVpfFind(ExpStnFd, tgMVef(ilVeh).iCode)).iSAGroupNo) And (ilSAGroupNo <> 0) Then
                        '            imDVefCode(UBound(imDVefCode)) = tgMVef(ilVeh).iCode
                        '            ReDim Preserve imDVefCode(0 To UBound(imDVefCode) + 1) As Integer
                        '        End If
                        '    End If
                        'Next ilVeh
                        If ilDVpfIndex >= 0 Then
                            ilVpf = tmVpfInfo(ilDVpfIndex).iFirstSALink
                            Do While ilVpf >= 0
                                imDVefCode(UBound(imDVefCode)) = tmSALink(ilVpf).iVefCode
                                ReDim Preserve imDVefCode(0 To UBound(imDVefCode) + 1) As Integer
                                ilVpf = tmSALink(ilVpf).iNextLkVehInfo
                            Loop
                        End If
                    Else
                        ReDim imDVefCode(0 To 1) As Integer
                        imDVefCode(0) = tmVef(ilDVehIndex).iCode
                    End If
                    'Only one vehicle has to match
                    ilMatch = False
                    'For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
                    '    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) <> 0 Then
                    '        For ilDVeh = LBound(tmVpfInfo(ilDVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilDVpfIndex).tVpf.iGLink) Step 1
                    '            If tmVpfInfo(ilDVpfIndex).tVpf.iGLink(ilDVeh) <> 0 Then
                    '                If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) = tmVpfInfo(ilDVpfIndex).tVpf.iGLink(ilDVeh) Then
                    '                ilMatch = True
                    '                Exit For
                    '                End If
                    '            End If
                    '        Next ilDVeh
                    '        If ilMatch Then
                    '            Exit For
                    '        End If
                    '    End If
                    'Next ilVeh
                    For ilVeh = LBound(imVefCode) To UBound(imVefCode) - 1 Step 1
                        For ilDVeh = LBound(imDVefCode) To UBound(imDVefCode) - 1 Step 1
                            If imVefCode(ilVeh) = imDVefCode(ilDVeh) Then
                                ilMatch = True
                                Exit For
                            End If
                        Next ilDVeh
                        If ilMatch Then
                            Exit For
                        End If
                    Next ilVeh

                End If
            End If
            If ilMatch Then
            'Test if same instructions
                ilTest = 0
                tmCnfSrchKey.lCrfCode = tlCrf.lCode
                tmCnfSrchKey.iInstrNo = 0
                ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tlCrf.lCode)
                    If ilTest >= UBound(llCifCode) Then
                        ilMatch = False
                        Exit Do
                    End If
                    If tmCnf.lCifCode <> llCifCode(ilTest) Then
                        ilMatch = False
                        Exit Do
                    End If
                    ilTest = ilTest + 1
                    ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If ilMatch Then
                    If ilDuplIndex = -1 Then
                        ilDuplIndex = ilCrf
                        If tgSortCrf(ilCrf).iDuplIndex = -1 Then
                            tgSortCrf(ilCrf).iDuplIndex = UBound(tgDuplCrf)
                        Else
                            'Test if combine or superseded
                            'Assume combine
                            ilTest = tgSortCrf(ilCrf).iDuplIndex
                            Do While tgDuplCrf(ilTest).iDuplIndex >= 0
                                ilTest = tgDuplCrf(ilTest).iDuplIndex
                            Loop
                            tgDuplCrf(ilTest).iDuplIndex = UBound(tgDuplCrf)
                        End If
                        'System can't have combine and superseded
                        mDuplRotation = 2
                        tgDuplCrf(UBound(tgDuplCrf)).lCntrNo = tmChf.lCntrNo
                        tgDuplCrf(UBound(tgDuplCrf)).sVehName = tmVef(ilVehIndex).sName
                        tgDuplCrf(UBound(tgDuplCrf)).tCrf = tmCrf
                        tgDuplCrf(UBound(tgDuplCrf)).lCrfRecPos = llRecPos
                        tgDuplCrf(UBound(tgDuplCrf)).iDuplIndex = -1
                        tgDuplCrf(UBound(tgDuplCrf)).iVpfIndex = ilVpfIndex
                        ReDim Preserve tgDuplCrf(0 To UBound(tgDuplCrf) + 1) As DUPLCRF
                    Else
                        'Merge
                        If tgSortCrf(ilDuplIndex).iDuplIndex = -1 Then
                            tgSortCrf(ilDuplIndex).iDuplIndex = UBound(tgDuplCrf)
                        Else
                            'Test if combine or superseded
                            'Assume combine
                            ilTest = tgSortCrf(ilDuplIndex).iDuplIndex
                            Do While tgDuplCrf(ilTest).iDuplIndex >= 0
                                ilTest = tgDuplCrf(ilTest).iDuplIndex
                            Loop
                            tgDuplCrf(ilTest).iDuplIndex = UBound(tgDuplCrf)
                        End If
                        'System can't have combine and superseded
                        mDuplRotation = 2
                        tgDuplCrf(UBound(tgDuplCrf)).lCntrNo = tgSortCrf(ilCrf).lCntrNo
                        slNameCode = tgSortCrf(ilCrf).sKey
                        ilRet = gParseItem(slNameCode, 3, "|", slName)
                        tgDuplCrf(UBound(tgDuplCrf)).sVehName = slName
                        tgDuplCrf(UBound(tgDuplCrf)).tCrf = tgSortCrf(ilCrf).tCrf
                        tgDuplCrf(UBound(tgDuplCrf)).lCrfRecPos = tgSortCrf(ilCrf).lCrfRecPos
                        tgDuplCrf(UBound(tgDuplCrf)).iDuplIndex = tgSortCrf(ilCrf).iDuplIndex
                        tgDuplCrf(UBound(tgDuplCrf)).iVpfIndex = tgSortCrf(ilCrf).iVpfIndex
                        ReDim Preserve tgDuplCrf(0 To UBound(tgDuplCrf) + 1) As DUPLCRF
                        For ilTest = ilCrf To UBound(tgSortCrf) - 1 Step 1
                            tgSortCrf(ilTest) = tgSortCrf(ilTest + 1)
                        Next ilTest
                        ReDim Preserve tgSortCrf(0 To UBound(tgSortCrf) - 1) As SORTCRF
                        ilCrf = ilCrf - 1   'Test next which has been moved
                    End If
                    'Exit Function
                Else
                    'Set combine flag
                    If tmVef(ilVehIndex).sType = "S" Then
                        If ilCombIndex = -1 Then
                            ilCombIndex = ilCrf
                            If tgSortCrf(ilCrf).iCombineIndex = -1 Then
                                tgSortCrf(ilCrf).iCombineIndex = UBound(tgCombineCrf)
                            Else
                                'Test if combine or superseded
                                'Assume combine
                                ilTest = tgSortCrf(ilCrf).iCombineIndex
                                Do While tgCombineCrf(ilTest).iCombineIndex >= 0
                                    ilTest = tgCombineCrf(ilTest).iCombineIndex
                                Loop
                                tgCombineCrf(ilTest).iCombineIndex = UBound(tgCombineCrf)
                            End If
                            mDuplRotation = 1
                            tgCombineCrf(UBound(tgCombineCrf)).lCntrNo = tmChf.lCntrNo
                            slNameCode = tgSortCrf(ilCrf).sKey
                            ilRet = gParseItem(slNameCode, 3, "|", slName)
                            tgCombineCrf(UBound(tgCombineCrf)).sVehName = tmVef(ilVehIndex).sName 'slName
                            tgCombineCrf(UBound(tgCombineCrf)).tCrf = tmCrf
                            tgCombineCrf(UBound(tgCombineCrf)).lCrfRecPos = llRecPos
                            tgCombineCrf(UBound(tgCombineCrf)).iCombineIndex = -1
                            tgCombineCrf(UBound(tgCombineCrf)).iVpfIndex = ilVpfIndex
                            ReDim Preserve tgCombineCrf(0 To UBound(tgCombineCrf) + 1) As COMBINECRF
                            'Exit Function
                        Else
                            'Merge
                            If tgSortCrf(ilCombIndex).iCombineIndex = -1 Then
                                tgSortCrf(ilCombIndex).iCombineIndex = UBound(tgCombineCrf)
                            Else
                                'Test if combine or superseded
                                'Assume combine
                                ilTest = tgSortCrf(ilCombIndex).iCombineIndex
                                Do While tgCombineCrf(ilTest).iCombineIndex >= 0
                                    ilTest = tgCombineCrf(ilTest).iCombineIndex
                                Loop
                                tgCombineCrf(ilTest).iCombineIndex = UBound(tgCombineCrf)
                            End If
                            'System can't have combine and superseded
                            mDuplRotation = 1
                            tgCombineCrf(UBound(tgCombineCrf)).lCntrNo = tgSortCrf(ilCrf).lCntrNo
                            tgCombineCrf(UBound(tgCombineCrf)).sVehName = tmVef(ilDVehIndex).sName
                            tgCombineCrf(UBound(tgCombineCrf)).tCrf = tgSortCrf(ilCrf).tCrf
                            tgCombineCrf(UBound(tgCombineCrf)).lCrfRecPos = tgSortCrf(ilCrf).lCrfRecPos
                            tgCombineCrf(UBound(tgCombineCrf)).iCombineIndex = tgSortCrf(ilCrf).iCombineIndex
                            tgCombineCrf(UBound(tgCombineCrf)).iVpfIndex = tgSortCrf(ilCrf).iVpfIndex
                            ReDim Preserve tgCombineCrf(0 To UBound(tgCombineCrf) + 1) As COMBINECRF
                            For ilTest = ilCrf To UBound(tgSortCrf) - 1 Step 1
                                tgSortCrf(ilTest) = tgSortCrf(ilTest + 1)
                            Next ilTest
                            ReDim Preserve tgSortCrf(0 To UBound(tgSortCrf) - 1) As SORTCRF
                            ilCrf = ilCrf - 1   'Test next which has been moved
                        End If
                    End If
                End If
            End If
        End If
        ilCrf = ilCrf + 1
    Loop
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mExportLine                     *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Send line to output             *
'*                                                     *
'*******************************************************
Private Function mExportLine(slRecord As String, ilLineNo As Integer, ilRecordType As Integer) As Integer
    Dim ilRet As Integer
    On Error GoTo mExportLineErr
    ilRet = 0
    If (((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value))) And (ilRecordType >= 0) Then
        gPackDate smPDFDate, tmTxr.iGenDate(0), tmTxr.iGenDate(1)
        'gPackTime smPDFTime, tmTXR.iGenTime(0), tmTXR.iGenTime(1)
        tmTxr.lGenTime = gTimeToLong(smPDFTime, False)
        lmPDFSeqNo = lmPDFSeqNo + 1
        tmTxr.lSeqNo = lmPDFSeqNo
        tmTxr.iType = ilRecordType
        tmTxr.sText = slRecord
        tmTxr.lCsfCode = 0
        ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            'Print #hmMsg, "Insert TXR Failed" & str$(ilRet) & " processing terminated"
            gAutomationAlertAndLogHandler "Insert TXR Failed" & str$(ilRet) & " processing terminated"
            mClearTxr
            imExporting = False
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("File in Use [Re-press Export], Insert Txr" & str(ilRet), vbOkOnly + vbExclamation, "Export")
            cmcCancel.SetFocus
            mExportLine = False
            Exit Function
        End If
    Else
        Print #hmExport, slRecord
        If ilRet <> 0 Then
            'Print #hmMsg, "Error writing to Export file" & str$(ilRet) & " processing terminated"
            gAutomationAlertAndLogHandler "Error writing to Export file" & str$(ilRet) & " processing terminated"
            imExporting = False
            Close #hmExport
            Screen.MousePointer = vbDefault
            ''MsgBox "Error writing to file" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            gAutomationAlertAndLogHandler "Error writing to file" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            cmcCancel.SetFocus
            mExportLine = False
            Exit Function
        End If
    End If
    ilLineNo = ilLineNo + 1
    mExportLine = True
    Exit Function
mExportLineErr:
    ilRet = err.Number
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mExpRot                         *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Export Copy and Rotation        *
'*                     instructions                    *
'*                                                     *
'*******************************************************
Private Function mExpRot(tlStnInfo As STNINFO) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVsf                         llSifCode                                               *
'******************************************************************************************

    Dim slExportFile As String
    Dim slMsgFile As String
    Dim slMsgFileName As String
    Dim ilMsgType As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slNameTime As String
    Dim slVehName As String
    Dim slCallLettersBand As String
    Dim slStr As String
    Dim ilVeh As Integer
    Dim ilVehIndex As Integer
    Dim ilVefSelected As Integer
    Dim ilPos As Integer
    Dim ilPageNo As Integer
    Dim ilLineNo As Integer
    Dim ilCopyNo As Integer
    Dim slRecord As String
    Dim ilTotalTime As Integer
    Dim ilDay As Integer
    Dim ilDayReq As Integer
    Dim llRotStartDate As Long
    Dim llRotEndDate As Long
    Dim llDate As Long
    Dim ilCheckInv As Integer
    Dim slComment As String
    'Dim slNowDate As String
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilTotalNoInstr As Integer
    'Dim ilLen As Integer
    Dim slPrevNewInv As String
    Dim slBlank As String
    Dim ilAnyTransmitted As Integer
    Dim ilCombineIndex As Integer
    'Dim ilDuplIndex As Integer
    'Dim ilCrfIndex As Integer
    'Dim ilLoop2 As Integer
    Dim ilVIndex As Integer
    Dim ilLastInvIndex As Integer
    Dim slCart As String
    Dim ilShowCart As Integer
    'Dim ilCombineOk As Integer
    Dim ilVpfIndex As Integer
    ReDim ilDayOn(0 To 6) As Integer
    Dim ilBeginMsgSent As Integer
    Dim slMsgLine As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilSAGroupNo As Integer
    Dim ilPrtFirstCopy As Integer
    Dim ilPrtFirstRot As Integer
    Dim ilNewHdRot As Integer
    Dim ilNoRotSent As Integer
    Dim hlMsg As Integer
    Dim slOutput As String
    Dim slPDFFileName As String
    Dim ilNoTimesMod As Integer
    Dim ilIncludeNewMessage As Integer
    Dim slShortTitle As String

    slBlank = ""
    mExpRot = False
    ReDim smFileNames(0 To 0) As String 'Use if abort issued
    For ilVeh = 0 To lbcVehicle.ListCount - 1 Step 1
        ilVefSelected = False
        slNameTime = lbcVehicle.List(ilVeh)
        ilPos = InStr(slNameTime, "|")
        slVehName = Left$(slNameTime, ilPos - 1)
        For ilLoop = 0 To UBound(tmVef) - 1 Step 1
            'ilLen = Len(Trim$(tmVef(ilLoop).sName))
            'If (Trim$(tmVef(ilLoop).sName) = Left$(slVehName, ilLen)) Then
            If tmVef(ilLoop).iCode = lbcVehicle.ItemData(ilVeh) Then
                ilSAGroupNo = tgVpf(gVpfFind(ExpStnFd, tmVef(ilLoop).iCode)).iSAGroupNo
                ilVehIndex = ilLoop
                If tmVef(ilLoop).iCode = tlStnInfo.iAirVeh Then
                    ilVefSelected = True
                    Exit For
                Else
                    ilVIndex = mFindVpfIndex(tmVef(ilLoop).iCode)
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                        Do While ilVIndex >= 0
                            If tlStnInfo.iAirVeh = tmLkVehInfo(ilVIndex).iVefCode Then
                                ilVefSelected = True
                                Exit For
                            End If
                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                End If
            End If
        Next ilLoop
        If ilVefSelected Then
            ilVefSelected = False
            'Test if vehicle has a rotation to be transmitted
            For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
                If lbcVeh.Selected(ilLoop) Then
                    slNameCode = tmVehCode(ilLoop).sKey    'lbcVehCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                    ilVpfIndex = gVpfFind(ExpStnFd, Val(slCode))
                    If tmVef(ilVehIndex).sType = "C" Then
                        If tmVef(ilVehIndex).iCode = Val(slCode) Then
                            ilVefSelected = True
                            Exit For
                        End If
                        ilVIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
                        If ilVIndex >= 0 Then
                            ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                            Do While ilVIndex >= 0
                                If tmLkVehInfo(ilVIndex).iVefCode = tmVef(ilVehIndex).iCode Then
                                    ilVefSelected = True
                                    Exit For
                                End If
                                ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                            Loop
                        End If
                   Else
                        ilVIndex = mFindVpfIndex(Val(slCode)) 'tmVef(ilVehIndex).iCode)
                        If ilVIndex >= 0 Then
                            ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                            Do While ilVIndex >= 0
                                If tmVef(ilVehIndex).iCode = tmSALink(ilVIndex).iVefCode Then
                                    ilVefSelected = True
                                    Exit For
                                End If
                                ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                            Loop
                        End If
                    End If
                End If
            Next ilLoop
        End If
        If ilVefSelected Then
            ilVpfIndex = gVpfFind(ExpStnFd, tmVef(ilVehIndex).iCode)
            'If tgVpf(ilVpfIndex).sExpBkCpyCart = "Y" Then
            '    ilShowCart = False  'True
            'Else
            '   ilShowCart = False
            '   If (tmVef(ilVehIndex).sType = "A") Or (tmVef(ilVehIndex).sType = "C") Then
            '       ilVIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
            '       If ilVIndex >= 0 Then
            '           ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
            '           Do While ilVIndex >= 0
            '               ilVpfIndex = gVpfFind(ExpStnFd, tmLkVehInfo(ilVIndex).iVefCode)
            '               If tgVpf(ilVpfIndex).sExpBkCpyCart = "Y" Then
            '                   ilShowCart = True
            '                   Exit Do
            '               End If
            '               ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
            '           Loop
            '       End If
            '   End If
            'End If
            If tgVpf(ilVpfIndex).sStnFdCart = "Y" Then
                ilShowCart = True
            Else
                ilShowCart = False
                If (tmVef(ilVehIndex).sType = "A") Or (tmVef(ilVehIndex).sType = "C") Then
                    ilVIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                        Do While ilVIndex >= 0
                            ilVpfIndex = gVpfFind(ExpStnFd, tmLkVehInfo(ilVIndex).iVefCode)
                            If tgVpf(ilVpfIndex).sStnFdCart = "Y" Then
                                ilShowCart = True
                                Exit Do
                            End If
                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                End If
            End If
            If tlStnInfo.sType = "S" Then
                ''slVehName = Trim$(tlStnInfo.sCallFreq)
                'slVehName = Trim$(tlStnInfo.sCallLetter) & "-" & Trim$(tlStnInfo.sBand)
                slCallLettersBand = Trim$(tlStnInfo.sCallLetter) & "-" & Trim$(tlStnInfo.sBand)
            End If
            lacProcessing.Caption = "Processing: " & slCallLettersBand
            DoEvents
            'slStnCode = Trim$(tmVef(ilVehIndex).sCodeStn)
            'slExportFile = sgExportPath & slStnCode & smFeedNo & ".msg"
            'slExportFile = sgExportPath & Trim$(tlStnInfo.sFileName) & ".trf"
            If rbcInterface(0).Value Then
                If tlStnInfo.sType = "G" Then
                    If rbcGen(3).Value Then
                        If rbcFormat(0).Value Then
                            slExportFile = sgExportPath & Trim$(tlStnInfo.sFileName) & ".PDF"
                            slPDFFileName = Trim$(tlStnInfo.sFileName) & ".PDF"
                        Else
                            slExportFile = sgExportPath & Trim$(tlStnInfo.sFileName) & ".Txt"
                        End If
                    Else
                        slExportFile = sgExportPath & Trim$(tlStnInfo.sFileName) & "." & tlStnInfo.sStnFdCode & "X"
                        slPDFFileName = Trim$(tlStnInfo.sFileName) & "." & tlStnInfo.sStnFdCode & "X"
                    End If
                Else
                    slExportFile = sgExportPath & Trim$(tlStnInfo.sFileName) & "." & tlStnInfo.sStnFdCode & "Z"
                    slPDFFileName = Trim$(tlStnInfo.sFileName) & "." & tlStnInfo.sStnFdCode & "Z"
                End If
            Else
                If tlStnInfo.sType = "G" Then
                    slExportFile = sgExportPath & Trim$(tlStnInfo.sFileName) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Traffic_Instructions(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Txt"
                    slPDFFileName = Trim$(tlStnInfo.sFileName) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Traffic_Instructions(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Pdf"
                Else
                    slExportFile = sgExportPath & Trim$(tlStnInfo.sFileName) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Traffic_Instructions(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Txt"
                    slPDFFileName = Trim$(tlStnInfo.sFileName) & "_Week_" & smWeekNo & "_" & Format(lmInputStartDate, "mm-dd-yyyy") & "_" & gFileNameFilter(slVehName) & "_Traffic_Instructions(Run_" & smRunLetter & "_" & Format(smTranDate, "mm-dd-yyyy") & ")" & ".Pdf"
                End If
            End If
            ilRet = 0
            'On Error GoTo cmcExportErr:
            If ((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
                smPDFDate = Format$(gNow(), "m/d/yy")
                gPackDate smPDFDate, imPDFDate(0), imPDFDate(1)
                smPDFTime = Format$(gNow(), "h:mm:ssAM/PM")
                gPackTime smPDFTime, imPDFTime(0), imPDFTime(1)
                lmPDFSeqNo = 0
            Else
                'hmExport = FreeFile
                ''Create file name based on vehicle name
                'Open slExportFile For Output As hmExport
                ilRet = gFileOpen(slExportFile, "Output", hmExport)
                If ilRet <> 0 Then
                    'Print #hmMsg, "Error Opening Export file" & str$(ilRet) & " processing terminated"
                    gAutomationAlertAndLogHandler "Error Opening Export file" & str$(ilRet) & " processing terminated"
                    Screen.MousePointer = vbDefault
                    ''MsgBox "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                    gAutomationAlertAndLogHandler "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                    cmcCancel.SetFocus
                    Exit Function
                End If
            End If
            smFileNames(UBound(smFileNames)) = slExportFile
            ReDim Preserve smFileNames(0 To UBound(smFileNames) + 1) As String 'Use if abort issued
            'Output new inventory
            ilPageNo = 0
            ilCopyNo = 1
            ilLineNo = 52
            ilTotalTime = 0
            slPrevNewInv = ""
            ilLastInvIndex = -1
            ilAnyTransmitted = False
            ilBeginMsgSent = False
            ilPrtFirstCopy = True
            ilPrtFirstRot = True
            'For ilLoop = 0 To UBound(tmAddCyf) - 1 Step 1
            '    'Eliminate Cart definition at top of export (1/21/01)
            '    If tmAddCyf(ilLoop).tCyf.iVefCode = tmVef(ilVehIndex).iCode Then
            '        If Not ilBeginMsgSent Then    'Add Start Page Note-
            '            ilNoCopyLines = 1
            '            GoSub cmcExportCopyHeader
            '            ilMsgType = 0
            '            slMsgFileName = Trim$(tlStnInfo.sStnFdCode) & "000000.SBg"
            '            GoSub cmcExportSendMsg
            '            slMsgFileName = Trim$(tlStnInfo.sStnFdCode) & smFeedNo & ".SBg"
            '            GoSub cmcExportSendMsg
            '            slMsgFileName = "SF" & smFeedNo & ".SBg"
            '            GoSub cmcExportSendMsg
            '            slMsgFileName = "SF000000.SBg"
            '            GoSub cmcExportSendMsg
            '            ilBeginMsgSent = True
            '        End If
            '        'Bypass any inventory that has been shown previously but is
            '        'only for a different zone
            '        If ilLastInvIndex <> -1 Then
            '            If (tmAddCyf(ilLastInvIndex).tCyf.lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode) And (tmAddCyf(ilLastInvIndex).tCyf.iVefCode = tmAddCyf(ilLoop).tCyf.iVefCode) Then
            '                ilShowInv = False
            '            Else
            '                ilShowInv = True
            '            End If
            '        Else
            '            ilShowInv = True
            '        End If
            '        If ilShowInv Then
            '            ilLastInvIndex = ilLoop
            '            tmCifSrchKey.lCode = tmAddCyf(ilLoop).tCyf.lCifCode
            '            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            '            If ilRet = BTRV_ERR_NONE Then
            '                tmCpfSrchKey.lCode = tmCif.lCpfCode
            '                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            '                If ilRet <> BTRV_ERR_NONE Then
            '                    tmCpf.sName = ""
            '                    tmCpf.sISCI = ""
            '                    tmCpf.sCreative = ""
            '                End If
            '                ilAnyTransmitted = True
            '                slRotStartDate = Format$(tmAddCyf(ilLoop).lRotStartDate, "mm/dd/yyyy")
            '                slRotEndDate = Format$(tmAddCyf(ilLoop).lRotEndDate, "mm/dd/yyyy")
            '                slRecord = "     "
            '                'If (Trim$(tmCpf.sName) <> "") Then
            '                If (Trim$(tmCpf.sName) <> "") And (tgSpf.sUseProdSptScr <> "P") Then
            '                    slRecord = slRecord & UCase(Trim$(tmCpf.sName))
            '                Else
            '                    slRecord = slRecord & UCase(Trim$(tmAddCyf(ilLoop).sChfProduct))
            '                End If
            '                Do While Len(slRecord) < 35
            '                    slRecord = slRecord & " "
            '                Loop
            '                slRecord = slRecord & "(St. " & slRotStartDate & " - " & slRotEndDate & ")"
            '                If StrComp(Trim$(slRecord), Trim$(slPrevNewInv), 1) <> 0 Then
            '                    slPrevNewInv = slRecord
            '                    ilNoCopyLines = 5   'Make sure enough room for advertiser and inventory
            '                    GoSub cmcExportCopyHeader
            '                    ilNoCopyLines = 1
            '                    If Len(slPrevNewInv) <> 0 Then
            '                        'slBlank = ""  'Blank line
            '                        If Not mExportLine(slBlank, ilLineNo) Then
            '                            Exit Function
            '                        End If
            '                        GoSub cmcExportCopyHeader
            '                        'slBlank = ""  'Blank line
            '                        If Not mExportLine(slBlank, ilLineNo) Then
            '                            Exit Function
            '                        End If
            '                    End If
            '                    If Not mExportLine(slRecord, ilLineNo) Then
            '                        Exit Function
            '                    End If
            '                    GoSub cmcExportCopyHeader
            '                    slRecord = "     "
            '                    slRecord = slRecord & "------------------------------"
            '                    If Not mExportLine(slRecord, ilLineNo) Then
            '                        Exit Function
            '                    End If
            '                    GoSub cmcExportCopyHeader
            '                Else
            '                    ilNoCopyLines = 1   'Make sure enough room for inventory
            '                    GoSub cmcExportCopyHeader
            '                End If
            '                ilNoCopyLines = 1   'Make sure enough room for inventory
            '                slRecord = Trim$(Str$(ilCopyNo))
            '                Do While Len(slRecord) < 2
            '                    slRecord = " " & slRecord
            '                Loop
            '                ilCopyNo = ilCopyNo + 1
            '                Do While Len(slRecord) < 5
            '                    slRecord = slRecord & " "
            '                Loop
            '                slRecord = slRecord & Trim$(tmCpf.sISCI)
            '                Do While Len(slRecord) < 27
            '                    slRecord = slRecord & " "
            '                Loop
            '                slRecord = slRecord & UCase(Trim$(tmCpf.sCreative))
            '                Do While Len(slRecord) < 66
            '                    slRecord = slRecord & " "
            '                Loop
            '                ilTotalTime = ilTotalTime + tmCif.iLen
            '                slRecord = slRecord & Trim$(Str$(tmCif.iLen))
            '                Do While Len(slRecord) < 70
            '                    slRecord = slRecord & " "
            '                Loop
            '                slRecord = slRecord & Format$(tmAddCyf(ilLoop).lPrevFdDate, "mm/dd/yyyy")
            '                If Not mExportLine(slRecord, ilLineNo) Then
            '                    Exit Function
            '                End If
            '                'GoSub cmcExportCopyHeader
            '                'slRecord = ""  'Blank line
            '                'If Not mExportLine(slRecord, ilLineNo) Then
            '                '    Exit Sub
            '                'End If
            '                'GoSub cmcExportCopyHeader
            '                'slRecord = ""  'Blank line
            '                'If Not mExportLine(slRecord, ilLineNo) Then
            '                '    Exit Sub
            '                'End If
            '            End If
            '        End If
            '    End If
            'Next ilLoop
            slPrevNewInv = ""
            If Not ilAnyTransmitted Then
                If Not ilBeginMsgSent Then    'Add Start Page Note-
                    'ilNoCopyLines = 1
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportRotHeader
                    If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                        Exit Function
                    End If
                    ilMsgType = 0
                    slMsgFileName = Trim$(tlStnInfo.sStnFdCode) & "000000.SBg"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    mExportSendMsg slMsgLine, slMsgFileName, slMsgFile, ilMsgType, ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord
                    slMsgFileName = Trim$(tlStnInfo.sStnFdCode) & smFeedNo & ".SBg"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    mExportSendMsg slMsgLine, slMsgFileName, slMsgFile, ilMsgType, ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord
                    slMsgFileName = "SF" & smFeedNo & ".SBg"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    mExportSendMsg slMsgLine, slMsgFileName, slMsgFile, ilMsgType, ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord
                    slMsgFileName = "SF000000.SBg"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    mExportSendMsg slMsgLine, slMsgFileName, slMsgFile, ilMsgType, ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord
                    ilBeginMsgSent = True
                End If
                'ilNoCopyLines = 1   'Make sure enough room for advertiser and inventory
                'GoSub cmcExportCopyHeader
                'slRecord = "*** No Commercials Being Fed Today ***"
                'If Not mExportLine(slRecord, ilLineNo) Then
                '    Exit Function
                'End If
                'If Not mExportLine(slBlank, ilLineNo) Then
                '    Exit Function
                'End If
            End If
            'GoSub cmcExportCopyHeader
            ''slBlank = ""  'Blank line
            'If Not mExportLine(slBlank, ilLineNo) Then
            '    Exit Function
            'End If
            'GoSub cmcExportCopyHeader
            'If Not mExportLine(slBlank, ilLineNo) Then
            '    Exit Function
            'End If
            'GoSub cmcExportCopyHeader
            'If Not mExportLine(slBlank, ilLineNo) Then
            '    Exit Function
            'End If
            'GoSub cmcExportCopyHeader
            ''slBlank = ""  'Blank line
            'If Not mExportLine(slBlank, ilLineNo) Then
            '    Exit Function
            'End If
            'GoSub cmcExportCopyHeader
            'slRecord = " "
            'Do While Len(slRecord) < 37
            '    slRecord = slRecord & " "
            'Loop
            'slRecord = slRecord & "TOTAL RUN TIME IN MINUTES:"
            'slSec = Trim$(Str$((100 * (ilTotalTime Mod 60)) / 60))
            'Do While Len(slSec) < 2
            '    slSec = slSec & "0"
            'Loop
            'slStr = Trim$(Str$(ilTotalTime \ 60)) & "." & slSec
            'Do While Len(slStr) < 10
            '    slStr = " " & slStr
            'Loop
            'slRecord = slRecord & slStr
            'If Not mExportLine(slRecord, ilLineNo) Then
            '    Exit Function
            'End If
            slRecord = "#   SHORT TITLE"
            Do While Len(slRecord) < 22
                slRecord = slRecord & " "
            Loop
            slRecord = slRecord & "FLIGHT DATES"
            Do While Len(slRecord) < 52
                slRecord = slRecord & " "
            Loop
            slRecord = slRecord & "LENGTH  ROTATION"
            If Not mExportLine(slRecord, ilLineNo, 2) Then
                Exit Function
            End If
            'Output rotation instructions
            ilNoRotSent = 0
            'ilLineNo = 52   'Force new page
            For ilLoop = 0 To UBound(tmRotInfo) - 1 Step 1
                If tmRotInfo(ilLoop).iStatus = 1 Then
                    If tmRotInfo(ilLoop).iVefCode = tmVef(ilVehIndex).iCode Then
                        tmCrfSrchKey.lCode = tmRotInfo(ilLoop).lCrfCode
                        ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            ilNewHdRot = True
                            ilAnyTransmitted = True
                            If tmChf.lCode <> tmCrf.lChfCode Then
                                tmChfSrchKey.lCode = tmCrf.lChfCode
                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            Else
                                ilRet = BTRV_ERR_NONE
                            End If
                            If ilRet = BTRV_ERR_NONE Then
                                slShortTitle = mGetShortTitle(tmChf, tmAdf, tmCrf.iVefCode)
                                mCheckForDuplOrCombines ilLoop, llRotStartDate, llRotEndDate, slShortTitle
                                ilNoRotSent = ilNoRotSent + 1
                                ilLineNo = 52   'Force new page with each rotation
                                '6/3/16: Replaced GoSub
                                'GoSub cmcExportRotHeader
                                If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                    Exit Function
                                End If
                                'slRecord = " "
                                'Do While Len(slRecord) < 24
                                '    slRecord = slRecord & " "
                                'Loop
                                slRecord = Trim$(str$(ilNoRotSent)) & "."
                                If tmAdf.iCode <> tmChf.iAdfCode Then
                                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                End If
'                                llSifCode = 0
'                                If tmChf.lVefCode < 0 Then
'                                    tmVsfSrchKey.lCode = -tmChf.lVefCode
'                                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                                    Do While ilRet = BTRV_ERR_NONE
'                                        For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
'                                            If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
'                                                If tmVsf.lFSComm(ilVsf) > 0 Then
'                                                    llSifCode = tmVsf.lFSComm(ilVsf)
'                                                End If
'                                                Exit For
'                                            End If
'                                        Next ilVsf
'                                        If llSifCode <> 0 Then
'                                            Exit Do
'                                        End If
'                                        If tmVsf.lLkVsfCode <= 0 Then
'                                            Exit Do
'                                        End If
'                                        tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
'                                        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                                    Loop
'                                End If
                                ''slRecord = slRecord & "  " & gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf)    'Trim$(tmChf.sProduct)
                                'slRecord = slRecord & "  " & UCase$(gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf))    'Trim$(tmChf.sProduct)
                                slRecord = slRecord & "  " & UCase$(slShortTitle)    'Trim$(tmChf.sProduct)
                                Do While Len(slRecord) < 22
                                    slRecord = slRecord & " "
                                Loop
'                                gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
'                                llRotStartDate = gDateValue(slDate)
'                                slDate = Format$(gDateValue(slDate), "mm/dd/yyyy")
                                slDate = Format$(llRotStartDate, "mm/dd/yyyy")
                                slRecord = slRecord & "(" & slDate
'                                gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
'                                llRotEndDate = gDateValue(slDate)
'                                slDate = Format$(gDateValue(slDate), "mm/dd/yyyy")
                                slDate = Format$(llRotEndDate, "mm/dd/yyyy")
                                slRecord = slRecord & "-" & slDate & ")"
                                ilFound = False
                                ilNoTimesMod = 0
                                For ilTest = 0 To UBound(lmReadyCRF) - 1 Step 1
                                    If lmReadyCRF(ilTest) = tmCrf.lCode Then
                                        ilFound = True
                                        ilNoTimesMod = imNoTimesMod(ilTest)
                                    End If
                                Next ilTest
                                'If (tmCrf.sAffFdStatus = "R") Or (ilFound) Then
                                ilIncludeNewMessage = False
                                'Changed 3/14/05:  Changed to handle rotation resent and show the correct title
                                If rbcInterface(0).Value Then
                                    If ((tmCrf.sAffFdStatus = "R") And (tmCrf.iAffFdWk = 0)) Or (ilFound) Then
                                        ilIncludeNewMessage = True
                                        ''slRecord = slRecord & "     New or Modified"
                                        'slRecord = slRecord & "     Modified"
                                        If (tmCrf.iNoTimesMod = 0) Or ((ilFound) And (ilNoTimesMod = 0)) Then
                                            slRecord = slRecord & "     New"
                                        Else
                                            slRecord = slRecord & "     Modified"
                                        End If
                                    End If
                                Else
                                    If ((tmCrf.sKCFdStatus = "R") And (tmCrf.iKCFdWk = 0)) Or (ilFound) Then
                                        ilIncludeNewMessage = True
                                        If (tmCrf.iKCNoTimesMod = 0) Or ((ilFound) And (ilNoTimesMod = 0)) Then
                                            slRecord = slRecord & "     New"
                                        Else
                                            slRecord = slRecord & "     Modified"
                                        End If
                                    End If
                                End If
                                If Not mExportLine(slRecord, ilLineNo, 3) Then
                                    Exit Function
                                End If
                                '6/3/16: Replaced GoSub
                                'GoSub cmcExportRotHeader
                                If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                    Exit Function
                                End If
                                'slRecord = ""  'Blank line
                                If Not mExportLine(slBlank, ilLineNo, 5) Then
                                    Exit Function
                                End If
                                '6/3/16: Replaced GoSub
                                'GoSub cmcExportRotHeader
                                If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                    Exit Function
                                End If
                                'slRecord = "Please air using the flight data and rotation indicated below:"
                                'If Not mExportLine(slRecord, ilLineNo) Then
                                '    Exit Function
                                'End If
                                'GoSub cmcExportRotHeader
                                ''slRecord = ""
                                'If Not mExportLine(slBlank, ilLineNo) Then
                                '    Exit Function
                                'End If
                                'GoSub cmcExportRotHeader
                                'If tmRotInfo(ilLoop).iRevised Then
                                '    slRecord = "        REVISED INSTRUCTIONS!!"
                                '    If Not mExportLine(slRecord, ilLineNo) Then
                                '        Exit Function
                                '    End If
                                '    GoSub cmcExportRotHeader
                                'End If
                                'GoSub SpecialInformation
                                ''slRecord = ""
                                'If Not mExportLine(slBlank, ilLineNo) Then
                                '    Exit Function
                                'End If
                                'GoSub cmcExportRotHeader
                                'If Not mExportLine(slBlank, ilLineNo) Then
                                '    Exit Function
                                'End If
                                'GoSub cmcExportRotHeader
                                'If Not mExportLine(slBlank, ilLineNo) Then
                                '    Exit Function
                                'End If
                                'GoSub cmcExportRotHeader
                                'slRecord = "COMM. NUMBER         COMMERCIAL NAME               LENGTH ROTATION" ' PRODUCT"
                                'If Not mExportLine(slRecord, ilLineNo) Then
                                '    Exit Function
                                'End If
                                'GoSub cmcExportRotHeader
                                'If Not mExportLine(slBlank, ilLineNo) Then
                                '    Exit Function
                                'End If
                                'GoSub cmcExportRotHeader
                                'If Not mExportLine(slBlank, ilLineNo) Then
                                '    Exit Function
                                'End If
                                'GoSub cmcExportRotHeader
                                ReDim tmCnfRot(0 To 0) As CNF
'                                ilCombineIndex = -1
                                ilCombineIndex = LBound(tgCombineCrf)
                                Do
                                    tmCnfSrchKey.lCrfCode = tmCrf.lCode
                                    tmCnfSrchKey.iInstrNo = 0
                                    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                                        ilFound = False
                                        For ilTest = 0 To UBound(tmCnfRot) - 1 Step 1
                                            If tmCnf.lCifCode = tmCnfRot(ilTest).lCifCode Then
                                                ilFound = True
                                                tmCnfRot(ilTest).iInstrNo = tmCnfRot(ilTest).iInstrNo + 1
                                                Exit For
                                            End If
                                        Next ilTest
                                        If Not ilFound Then
                                            tmCnfRot(UBound(tmCnfRot)) = tmCnf
                                            tmCnfRot(UBound(tmCnfRot)).iInstrNo = 1
                                            ReDim Preserve tmCnfRot(0 To UBound(tmCnfRot) + 1) As CNF
                                        End If
                                        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
'                                    ilCombineOk = False
'                                    Do
'                                        If ilCombineIndex = -1 Then
'                                            ilCrfIndex = tmRotInfo(ilLoop).iSortCrfIndex
'                                            If ilCrfIndex >= 0 Then
'                                                ilCombineIndex = tgSortCrf(ilCrfIndex).iCombineIndex
'                                            End If
'                                        Else
'                                            ilCombineIndex = tgCombineCrf(ilCombineIndex).iCombineIndex
'                                        End If
'                                        If ilCombineIndex >= 0 Then
'                                            'Check if this rotation if sent to this vehicle
'                                            ilVpfIndex = tgCombineCrf(ilCombineIndex).iVpfIndex
'                                            'For ilCheck = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
'                                            '    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilCheck) <> 0 Then
'                                            '        If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilCheck) = tmVef(ilVehIndex).iCode Then
'                                            '            tmCrf = tgCombineCrf(ilCombineIndex).tCrf
'                                            '            ilCombineOk = True
'                                            '            Exit For
'                                            '        End If
'                                            '    End If
'                                            'Next ilCheck
'                                            'ilVIndex = tmVpfInfo(ilVpfIndex).iFirstLkVehInfo
'                                            'Do While ilVIndex > 0
'                                            '    ilVpfIndex = gVpfFind(ExpStnFd, tmLkVehInfo(ilVIndex).iVefCode)
'                                            '    If (tgVpf(ilVpfIndex).iSAGroupNo = ilSAGroupNo) And (ilSAGroupNo <> 0) Then
'                                            '        ilVefSelected = True
'                                            '        Exit Do
'                                            '    End If
'                                            '    ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
'                                            'Loop
'                                            'If conventional- must use SAGroupNo- Code Later
'                                            If ilVpfIndex >= 0 Then
'                                                ilVIndex = tmVpfInfo(ilVpfIndex).iFirstSALink
'                                                Do While ilVIndex >= 0
'                                                    If tmVef(ilVehIndex).iCode = tmSALink(ilVIndex).iVefCode Then
'                                                        'tmCrf = tgCombineCrf(ilCombineIndex).tCrf
'                                                        'ilCombineOk = True
'                                                        'Exit Do
'                                                        If tmRotInfo(ilLoop).lCrfCode <> tgCombineCrf(ilCombineIndex).tCrf.lCode Then
'                                                            tmCrf = tgCombineCrf(ilCombineIndex).tCrf
'                                                            ilCombineOk = True
'                                                            Exit Do
'                                                        End If
'                                                    End If
'                                                    ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
'                                                Loop
'                                            End If
'                                        Else
'                                            ilCombineOk = True
'                                        End If
'                                    Loop While Not ilCombineOk
'                                Loop While ilCombineIndex >= 0
                                    If ilCombineIndex >= UBound(tgCombineCrf) Then
                                        Exit Do
                                    End If
                                    tmCrf = tgCombineCrf(ilCombineIndex).tCrf
                                    ilCombineIndex = ilCombineIndex + 1
                                Loop
                                ilTotalNoInstr = 0
                                For ilTest = 0 To UBound(tmCnfRot) - 1 Step 1
                                    ilTotalNoInstr = ilTotalNoInstr + tmCnfRot(ilTest).iInstrNo
                                Next ilTest
                                For ilTest = 0 To UBound(tmCnfRot) - 1 Step 1
                                    tmCnf = tmCnfRot(ilTest)
                                    tmCifSrchKey.lCode = tmCnf.lCifCode
                                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        tmCpfSrchKey.lCode = tmCif.lcpfCode
                                        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            tmCpf.sName = ""
                                            tmCpf.sISCI = ""
                                            tmCpf.sCreative = ""
                                        End If
                                        '6/3/16: Replaced GoSub
                                        'GoSub cmcExportRotHeader
                                        If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                            Exit Function
                                        End If
                                        slRecord = Trim$(tmCpf.sISCI)
                                        Do While Len(slRecord) < 22
                                            slRecord = slRecord & " "
                                        Loop
                                        slRecord = slRecord & UCase(Trim$(tmCpf.sCreative))
                                        Do While Len(slRecord) < 54
                                            slRecord = slRecord & " "
                                        Loop
                                        slRecord = slRecord & Trim$(str$(tmCif.iLen))
                                        Do While Len(slRecord) < 61
                                            slRecord = slRecord & " "
                                        Loop
                                        'Rotation %
                                        slRecord = slRecord & gDivStr(gMulStr(Trim$(str$(tmCnf.iInstrNo)), "100"), str$(ilTotalNoInstr)) & "%"
                                        'Product Name
                                        Do While Len(slRecord) < 70
                                            slRecord = slRecord & " "
                                        Loop
                                        'If (Trim$(tmCpf.sName) <> "") And (StrComp(Trim$(tmChf.sProduct), Trim$(tmCpf.sName), 1) <> 0) Then
                                        '    slRecord = slRecord & Trim$(tmCpf.sName)
                                        'End If
                                        For ilCheckInv = 0 To UBound(tmAddCyf) - 1 Step 1
                                            If (tmAddCyf(ilCheckInv).tCyf.lCifCode = tmCnf.lCifCode) And (tmAddCyf(ilCheckInv).tCyf.iVefCode = tmVef(ilVehIndex).iCode) And (tmAddCyf(ilCheckInv).tCyf.sTimeZone = tmCrf.sZone) And (tmAddCyf(ilCheckInv).iFdDateNew) Then
                                                slRecord = slRecord & "New"
                                                Exit For
                                            End If
                                        Next ilCheckInv
                                        Do While Len(slRecord) < 75
                                            slRecord = slRecord & " "
                                        Loop
                                        If ilShowCart Then
                                            If tmMcf.iCode <> tmCif.iMcfCode Then
                                                tmMcfSrchKey.iCode = tmCif.iMcfCode
                                                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    tmMcf.sName = ""
                                                End If
                                            End If
                                            slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                                            If (Len(Trim$(tmCif.sCut)) <> 0) Then
                                                slCart = slCart & "-" & tmCif.sCut
                                            End If
                                            slRecord = slRecord & slCart
                                        End If
                                        If Not mExportLine(slRecord, ilLineNo, 4) Then
                                            Exit Function
                                        End If
                                        '6/3/16: Replaced GoSub
                                        'GoSub cmcExportRotHeader
                                        If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                            Exit Function
                                        End If
                                        ''Output cart
                                        ''ilShowCart = False
                                        ''If (InStr(UCase(slVehName), "SATELLITE MUSIC") > 0) Or (InStr(UCase(slVehName), "24-HOUR FORMATS") > 0) Then
                                        ''    ilShowCart = True
                                        ''End If
                                        ''If (InStr(UCase(slVehName), "SMN MIX") > 0) Or (InStr(UCase(slVehName), "MIX 24-HOUR FORMATS") > 0) Then
                                        ''    ilShowCart = True
                                        ''End If
                                        'If ilShowCart Then
                                        '    If tmMcf.iCode <> tmCif.iMcfCode Then
                                        '        tmMcfSrchKey.iCode = tmCif.iMcfCode
                                        '        ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        '        If ilRet <> BTRV_ERR_NONE Then
                                        '            tmMcf.sName = ""
                                        '        End If
                                        '    End If
                                        '    slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                                        '    If (Len(Trim$(tmCif.sCut)) <> 0) Then
                                        '        slCart = slCart & "-" & tmCif.sCut
                                        '    End If
                                        '    slRecord = " "
                                        '    Do While Len(slRecord) < 21
                                        '        slRecord = slRecord & " "
                                        '    Loop
                                        '    slRecord = slRecord & slCart
                                        '    If Not mExportLine(slRecord, ilLineNo) Then
                                        '        Exit Function
                                        '    End If
                                        '    GoSub cmcExportRotHeader
                                        'End If
                                        'Output comment if defined
                                        If tmCif.lCsfCode > 0 Then
                                            tmCsfSrchKey.lCode = tmCif.lCsfCode
                                            tmCsf.sComment = ""
                                            imCsfRecLen = Len(tmCsf) '5011
                                            ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                'Output 70 characters per line
                                                'If tmCsf.iStrLen > 0 Then
                                                slStr = gStripChr0(tmCsf.sComment)
                                                If slStr <> "" Then
                                                    slComment = slStr 'Trim$(Left$(tmCsf.sComment, tmCsf.iStrLen))
                                                    slRecord = " "
                                                    Do While Len(slRecord) < 21
                                                        slRecord = slRecord & " "
                                                    Loop
                                                    Do While Len(slComment) > 0
                                                        ilPos = InStr(slComment, " ")
                                                        If ilPos > 0 Then
                                                            If Len(slRecord) + ilPos - 1 > 70 Then
                                                                If Not mExportLine(slRecord, ilLineNo, 5) Then
                                                                    Exit Function
                                                                End If
                                                                '6/3/16: Replaced GoSub
                                                                'GoSub cmcExportRotHeader
                                                                If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                                                    Exit Function
                                                                End If
                                                                slRecord = " "
                                                                Do While Len(slRecord) < 21
                                                                    slRecord = slRecord & " "
                                                                Loop
                                                            End If
                                                            slRecord = slRecord & Left$(slComment, ilPos)
                                                            slComment = right$(slComment, Len(slComment) - ilPos)
                                                        Else
                                                            If Len(slRecord) + Len(slComment) > 70 Then
                                                                If Not mExportLine(slRecord, ilLineNo, 5) Then
                                                                    Exit Function
                                                                End If
                                                                '6/3/16: Replaced GoSub
                                                                'GoSub cmcExportRotHeader
                                                                If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                                                    Exit Function
                                                                End If
                                                                slRecord = " "
                                                                Do While Len(slRecord) < 21
                                                                    slRecord = slRecord & " "
                                                                Loop
                                                            End If
                                                            slRecord = slRecord & slComment
                                                            If Not mExportLine(slRecord, ilLineNo, 5) Then
                                                                Exit Function
                                                            End If
                                                            '6/3/16: Replaced GoSub
                                                            'GoSub cmcExportRotHeader
                                                            If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                                                Exit Function
                                                            End If
                                                            slComment = ""
                                                            Exit Do
                                                        End If
                                                    Loop
                                                End If
                                            End If
                                        End If
                                        'slRecord = " "
                                        'Do While Len(slRecord) < 21
                                        '    slRecord = slRecord & " "
                                        'Loop
                                        'ilNewInv = False
                                        'For ilCheckInv = 0 To UBound(tmAddCyf) - 1 Step 1
                                        '    If (tmAddCyf(ilCheckInv).tCyf.lCifCode = tmCnf.lCifCode) And (tmAddCyf(ilCheckInv).tCyf.iVefCode = tmVef(ilVehIndex).iCode) And (tmAddCyf(ilCheckInv).tCyf.sTimeZone = tmCrf.sZone) Then
                                        '        ilNewInv = True
                                        '        Exit For
                                        '    End If
                                        'Next ilCheckInv
                                        'If ilNewInv Then
                                        '    slRecord = slRecord & "Date fed (sent): " & smTranDate
                                        '    If Not mExportLine(slRecord, ilLineNo) Then
                                        '        Exit Function
                                        '    End If
                                        '    GoSub cmcExportRotHeader
                                        'Else
                                        '    'Find last sent date
                                        '    tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                                        '    tmCyfSrchKey.iVefCode = tmVef(ilVehIndex).iCode
                                        '    tmCyfSrchKey.sSource = "S"
                                        '    If Trim$(tmCrf.sZone) <> "R" Then
                                        '        tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                         '       tmCyfSrchKey.lRafCode = 0
                                        '    Else
                                        '        tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                        '        tmCyfSrchKey.lRafCode = tmCrf.lRafCode
                                        '    End If
                                        '    ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        '    If ilRet <> BTRV_ERR_NONE Then
                                        '        slRecord = slRecord & "Date fed (sent): " & smTranDate
                                        '        If Not mExportLine(slRecord, ilLineNo) Then
                                        '            Exit Function
                                        '        End If
                                        '        If ilTest <> UBound(tmCnfRot) - 1 Then
                                        '            GoSub cmcExportRotHeader
                                        '        End If
                                        '    Else
                                        '        gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slDate
                                        '        slDate = Format$(gDateValue(slDate), "mm/dd/yyyy")
                                        '        slRecord = slRecord & "Date fed (sent): " & slDate
                                        '        If Not mExportLine(slRecord, ilLineNo) Then
                                        '            Exit Function
                                        '        End If
                                        '        GoSub cmcExportRotHeader
                                        '        slRecord = " "
                                        '        Do While Len(slRecord) < 19
                                        '            slRecord = slRecord & " "
                                        '        Loop
                                        '        slRecord = slRecord & "(FED (SENT) PREVIOUSLY, NOT TODAY!)"
                                        '        If Not mExportLine(slRecord, ilLineNo) Then
                                        '            Exit Function
                                        '        End If
                                        '        If ilTest <> UBound(tmCnfRot) - 1 Then
                                        '            GoSub cmcExportRotHeader
                                        '        End If
                                        '    End If
                                        'End If
                                        'slDate = ""
                                        'For ilCheckInv = 0 To UBound(tmAddCyf) - 1 Step 1
                                        '    If (tmAddCyf(ilCheckInv).tCyf.lCifCode = tmCnf.lCifCode) And (tmAddCyf(ilCheckInv).tCyf.iVefCode = tmVef(ilVehIndex).iCode) And (tmAddCyf(ilCheckInv).tCyf.sTimeZone = tmCrf.sZone) Then
                                        '        If Trim$(tmCrf.sZone) <> "R" Then
                                        '            If (tmAddCyf(ilLoop).tCyf.sTimeZone = tmCrf.sZone) And (tmAddCyf(ilLoop).tCyf.lRafCode = 0) Then
                                        '                slDate = Format$(tmAddCyf(ilLoop).lPrevFdDate, "m/d/yy")
                                        '                ilNewInv = tmAddCyf(ilLoop).iFdDateNew
                                        '                Exit For
                                        '            End If
                                        '        Else
                                        '            If (tmAddCyf(ilLoop).tCyf.sTimeZone = tmCrf.sZone) And (tmAddCyf(ilLoop).tCyf.lRafCode = tmCrf.lRafCode) Then
                                        '                slDate = Format$(tmAddCyf(ilLoop).lPrevFdDate, "m/d/yy")
                                        '                ilNewInv = tmAddCyf(ilLoop).iFdDateNew
                                        '                Exit For
                                        '            End If
                                        '        End If
                                        '        Exit For
                                        '    End If
                                        'Next ilCheckInv
                                        'If slDate = "" Then
                                        '    mGetSentDate tmVef(ilVehIndex).iCode, ilNewInv, slDate
                                        'End If
                                        'slRecord = slRecord & "Date fed (sent): " & slDate
                                        'If Not mExportLine(slRecord, ilLineNo) Then
                                        '    Exit Function
                                        'End If
                                        'GoSub cmcExportRotHeader
                                        'If Not ilNewInv Then
                                        '    slRecord = " "
                                        '    Do While Len(slRecord) < 19
                                        '        slRecord = slRecord & " "
                                        '    Loop
                                        '    slRecord = slRecord & "(FED (SENT) PREVIOUSLY, NOT TODAY!)"
                                        '    If Not mExportLine(slRecord, ilLineNo) Then
                                        '        Exit Function
                                        '    End If
                                        'End If
                                        'If ilTest <> UBound(tmCnfRot) - 1 Then
                                        '    GoSub cmcExportRotHeader
                                        'End If
                                        'If ilTest <> UBound(tmCnfRot) - 1 Then
                                        '    'slRecord = ""
                                        '    If Not mExportLine(slBlank, ilLineNo) Then
                                        '        Exit Function
                                        '    End If
                                        '    GoSub cmcExportRotHeader
                                        '    'slRecord = ""
                                        '    If Not mExportLine(slBlank, ilLineNo) Then
                                        '        Exit Function
                                        '    End If
                                        'End If
                                    End If
                                Next ilTest
                                '6/3/16: Replaced GoSub
                                'GoSub SpecialInstructions
                                If Not mSpecialInstructions(llRotStartDate, llRotEndDate, ilIncludeNewMessage, ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next ilLoop
            'slRecord = ""
            If Not mExportLine(slBlank, ilLineNo, 5) Then
                Exit Function
            End If
            '6/3/16: Replaced GoSub
            'GoSub cmcExportRotHeader
            If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                Exit Function
            End If
            'slRecord = ""
            If Not mExportLine(slBlank, ilLineNo, 5) Then
                Exit Function
            End If
            '6/3/16: Replaced GoSub
            'GoSub cmcExportRotHeader
            If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                Exit Function
            End If
            'slRecord = "TOTAL NUMBER OF PAGES FOR " & slVehName & ":"
            'slStr = Trim$(Str$(ilPageNo))
            slRecord = "TOTAL NUMBER OF ROTATIONS FOR " & slCallLettersBand & ":"
            slStr = Trim$(str$(ilNoRotSent))
            Do While Len(slStr) < 5
                slStr = " " & slStr
            Loop
            slRecord = slRecord & slStr
            If Not mExportLine(slRecord, ilLineNo, 5) Then
                Exit Function
            End If
            If ilLineNo < 50 Then
                '6/3/16: Replaced GoSub
                'GoSub cmcExportRotHeader
                If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                    Exit Function
                End If
                'Add Messages
                If Not mExportLine(slBlank, ilLineNo, 5) Then
                    Exit Function
                End If
                '6/3/16: Replaced GoSub
                'GoSub cmcExportRotHeader
                If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                    Exit Function
                End If
            End If
            ilMsgType = 1
            slMsgFileName = Trim$(tlStnInfo.sStnFdCode) & "000000.SEn"
            '6/3/16: Replaced GoSub
            'GoSub cmcExportSendMsg
            mExportSendMsg slMsgLine, slMsgFileName, slMsgFile, ilMsgType, ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord
            slMsgFileName = Trim$(tlStnInfo.sStnFdCode) & smFeedNo & ".SEn"
            '6/3/16: Replaced GoSub
            'GoSub cmcExportSendMsg
            mExportSendMsg slMsgLine, slMsgFileName, slMsgFile, ilMsgType, ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord
            slMsgFileName = "SF" & smFeedNo & ".SEn"
            '6/3/16: Replaced GoSub
            'GoSub cmcExportSendMsg
            mExportSendMsg slMsgLine, slMsgFileName, slMsgFile, ilMsgType, ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord
            slMsgFileName = "SF000000.SEn"
            '6/3/16: Replaced GoSub
            'GoSub cmcExportSendMsg
            mExportSendMsg slMsgLine, slMsgFileName, slMsgFile, ilMsgType, ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord
            If ((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
                'Output report so that file is generated
                igRptCallType = STATIONFEEDJOB
                igRptType = 1
                slOutput = "2"
                If (Not igStdAloneMode) And (imShowHelpMsg) Then
                    If igTestSystem Then
                        slStr = "ExpStnFd^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & smPDFDate & "\" & smPDFTime
                    Else
                        slStr = "ExpStnFd^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & smPDFDate & "\" & smPDFTime
                    End If
                Else
                    If igTestSystem Then
                        slStr = "ExpStnFd^Test^NOHELP\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & smPDFDate & "\" & smPDFTime
                    Else
                        slStr = "ExpStnFd^Prod^NOHELP\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & smPDFDate & "\" & smPDFTime
                    End If
                End If
                'ilShell = Shell(sgExePath & "RptSelTx.Exe " & slStr, 1)
                'While GetModuleUsage(ilShell) > 0
                '    ilRet = DoEvents()
                'Wend
                sgCommandStr = slStr
                RptSelTx.Show vbModal
                ''Rename file
                'ilRet = 0
                'On Error GoTo cmcExportErr:
                'slStr = FileDateTime(slExportFile)
                'If ilRet = 0 Then
                '    Kill slExportFile
                'End If
                'On Error GoTo 0
                'ilRet = 0
                'FileCopy "c:\csi\csirpt.pdf", slExportFile
                'Remove records
                mClearTxr
            Else
                Close hmExport
            End If
        End If
    Next ilVeh
    If (rbcInterface(0).Value) And (rbcGen(3).Value) Then
        mExpRot = True
        Exit Function
    End If
    'Update files
    ilRet = btrBeginTrans(hmCrf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        'Print #hmMsg, "Begin Transaction Failed" & str$(ilRet) & " processing terminated"
        gAutomationAlertAndLogHandler "Begin Transaction Failed: " & str$(ilRet) & ", processing terminated"
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("File in Use [Re-press Export], BeginTran" & str(ilRet), vbOkOnly + vbExclamation, "Export")
        Exit Function
    End If
    'lacProcessing.Caption = "Updating Copy Inventory"
    'DoEvents
    'For ilLoop = 0 To UBound(tmAddCyf) - 1 Step 1
    '    'If transmitted- remove old record, then insert instead of updating
    '    Do
    '        tmCyfSrchKey.lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode
    '        tmCyfSrchKey.iVefCode = tmAddCyf(ilLoop).tCyf.iVefCode
    '        tmCyfSrchKey.sSource = "S"
    '        If Trim$(tmAddCyf(ilLoop).tCyf.sTimeZone) <> "R" Then
    '            tmCyfSrchKey.sTimeZone = tmAddCyf(ilLoop).tCyf.sTimeZone
    '            tmCyfSrchKey.lRafCode = 0
    '        Else
    '            tmCyfSrchKey.sTimeZone = tmAddCyf(ilLoop).tCyf.sTimeZone
    '            tmCyfSrchKey.lRafCode = tmAddCyf(ilLoop).tCyf.lRafCode
    '        End If
    '        ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    '        If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
    '            ilRet = btrDelete(hmCyf)
    '        Else
    '            ilRet = BTRV_ERR_NONE
    '        End If
    '    Loop While ilRet = BTRV_ERR_CONFLICT
    '    If ilRet <> BTRV_ERR_NONE Then
    '        Print #hmMsg, "Delete CYF Failed" & Str$(ilRet) & " processing terminated"
    '        mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
    '        Screen.MousePointer = vbDefault
    '        ilRet = MsgBox("File in Use [Re-press Export], Delete Cyf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
    '        Exit Function
    '    End If
    '    ilRet = btrInsert(hmCyf, tmAddCyf(ilLoop).tCyf, imCyfRecLen, INDEXKEY0)
    '    If ilRet <> BTRV_ERR_NONE Then
    '        Print #hmMsg, "Insert CYF Failed" & Str$(ilRet) & " processing terminated"
    '        mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
    '        Screen.MousePointer = vbDefault
    '        ilRet = MsgBox("File in Use [Re-press Export], Insert Cyf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
    '        Exit Function
    '    End If
    '    'Test if airing and group vehicles defined- if so insert for other
    '    'vehicles
    '    For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
    '        If tmVef(ilLoop1).iCode = tmAddCyf(ilLoop).tCyf.iVefCode Then
    '            If (tmVef(ilLoop1).sType = "A") Or (tmVef(ilLoop1).sType = "C") Then
    '                'Update cyf for all vehicles
    '                ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
    '                'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
    '                If ilVIndex >= 0 Then
    '                    ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
    '                    Do While ilVIndex >= 0
    '                        Do
    '                            tmCyfSrchKey.lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode
    '                            tmCyfSrchKey.iVefCode = tmLkVehInfo(ilVIndex).iVefCode'tmVpfInfo(ilVIndex).iVefLink(ilLoop2)
    '                            tmCyfSrchKey.sSource = "S"
    '                            'tmCyfSrchKey.sTimeZone = tmAddCyf(ilLoop).tCyf.sTimeZone
    '                            If Trim$(tmAddCyf(ilLoop).tCyf.sTimeZone) <> "R" Then
    '                                tmCyfSrchKey.sTimeZone = tmAddCyf(ilLoop).tCyf.sTimeZone
    '                                tmCyfSrchKey.lRafCode = 0
    '                            Else
    '                                tmCyfSrchKey.sTimeZone = tmAddCyf(ilLoop).tCyf.sTimeZone
    '                                tmCyfSrchKey.lRafCode = tmAddCyf(ilLoop).tCyf.lRafCode
    '                            End If
    '                            ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    '                            If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
    '                                ilRet = btrDelete(hmCyf)
    '                            Else
    '                                ilRet = BTRV_ERR_NONE
    '                            End If
    '                        Loop While ilRet = BTRV_ERR_CONFLICT
    '                        If ilRet <> BTRV_ERR_NONE Then
    '                            Print #hmMsg, "Delete Linked CYF Failed" & Str$(ilRet) & " processing terminated"
    '                            mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
    '                            Screen.MousePointer = vbDefault
    '                            ilRet = MsgBox("File in Use [Re-press Export], Delete Cyf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
    '                            Exit Function
    '                        End If
    '                        tlCyf = tmAddCyf(ilLoop).tCyf
    '                        tlCyf.iVefCode = tmLkVehInfo(ilVIndex).iVefCode'tmVpfInfo(ilVIndex).iVefLink(ilLoop2)
    '                        ilRet = btrInsert(hmCyf, tlCyf, imCyfRecLen, INDEXKEY0)
    '                        If ilRet <> BTRV_ERR_NONE Then
    '                            Print #hmMsg, "Insert Linked CYF Failed" & Str$(ilRet) & " processing terminated"
    '                            mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
    '                            Screen.MousePointer = vbDefault
    '                            ilRet = MsgBox("File in Use [Re-press Export], Insert Cyf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
    '                            Exit Function
    '                        End If
    '                    'Next ilLoop2
    '                        ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
    '                    Loop
    '                End If
    '            End If
    '        End If
    '    Next ilLoop1
    'Next ilLoop
    lacProcessing.Caption = "Updating Copy Rotations"
    DoEvents
    ReDim tgDuplCrf(0 To 0) As DUPLCRF
    ReDim tgCombineCrf(0 To 0) As COMBINECRF
    For ilLoop = 0 To UBound(tmRotInfo) - 1 Step 1
        Do
            tmCrfSrchKey.lCode = tmRotInfo(ilLoop).lCrfCode
            ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                'Print #hmMsg, "Get CRF to Update Feed Dates Failed" & str$(ilRet) & " processing terminated"
                gAutomationAlertAndLogHandler "Get CRF to Update Feed Dates Failed" & str$(ilRet) & " processing terminated"
                mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("File in Use [Re-press Export], GetEqual Crf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
                Exit Function
            End If
            'If tmCrf.sAffFdStatus = "R" Then
            'Changed 2/14/05:  Changed so that resent will show the correct title
            If rbcInterface(0).Value Then
                If (tmCrf.sAffFdStatus = "R") And (tmCrf.iAffFdWk = 0) Then
                    ilFound = False
                    For ilTest = 0 To UBound(lmReadyCRF) - 1 Step 1
                        If lmReadyCRF(ilTest) = tmCrf.lCode Then
                            ilFound = True
                        End If
                    Next ilTest
                    If (Not ilFound) Then
                        lmReadyCRF(UBound(lmReadyCRF)) = tmCrf.lCode
                        ReDim Preserve lmReadyCRF(0 To UBound(lmReadyCRF) + 1) As Long
                        imNoTimesMod(UBound(imNoTimesMod)) = tmCrf.iNoTimesMod
                        ReDim Preserve imNoTimesMod(0 To UBound(imNoTimesMod) + 1) As Integer
                    End If
                End If
                tmCrf.sAffFdStatus = "S" '"S"
                tmCrf.sAffXMitChar = smRunLetter
                'tmCrf.iFeedDate(0) = ilTranDate0
                'tmCrf.iFeedDate(1) = ilTranDate1
                'gPackDateLong lmInputStartDate, tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1)
                mSetAffFdDate
            Else
                If (tmCrf.sKCFdStatus = "R") And (tmCrf.iKCFdWk = 0) Then
                    ilFound = False
                    For ilTest = 0 To UBound(lmReadyCRF) - 1 Step 1
                        If lmReadyCRF(ilTest) = tmCrf.lCode Then
                            ilFound = True
                        End If
                    Next ilTest
                    If (Not ilFound) Then
                        lmReadyCRF(UBound(lmReadyCRF)) = tmCrf.lCode
                        ReDim Preserve lmReadyCRF(0 To UBound(lmReadyCRF) + 1) As Long
                        imNoTimesMod(UBound(imNoTimesMod)) = tmCrf.iKCNoTimesMod
                        ReDim Preserve imNoTimesMod(0 To UBound(imNoTimesMod) + 1) As Integer
                    End If
                End If
                tmCrf.sKCFdStatus = "S" '"S"
                tmCrf.sKCXMitChar = smRunLetter
                'tmCrf.iFeedDate(0) = ilTranDate0
                'tmCrf.iFeedDate(1) = ilTranDate1
                'gPackDateLong lmInputStartDate, tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1)
                mSetKCFdDate
            End If
            ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            'Print #hmMsg, "Update CRF Feed Dates Failed" & str$(ilRet) & " processing terminated"
            gAutomationAlertAndLogHandler "Update CRF Feed Dates Failed" & str$(ilRet) & " processing terminated"
            mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("File in Use [Re-press Export], Update Crf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
            Exit Function
        End If
'        ilCrfIndex = tmRotInfo(ilLoop).iSortCrfIndex
'        If ilCrfIndex >= 0 Then
'            ilCombineIndex = tgSortCrf(ilCrfIndex).iCombineIndex
'            Do While ilCombineIndex >= 0
'                Do
'                    ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgCombineCrf(ilCombineIndex).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        Print #hmMsg, "Get CRF Combo to update Feed Dates Failed" & Str$(ilRet) & " processing terminated"
'                        mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
'                        Screen.MousePointer = vbDefault
'                        ilRet = MsgBox("File in Use [Re-press Export], GetDirect Combine Crf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
'                        Exit Function
'                    End If
'                    'tmRec = tmCrf
'                    'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
'                    'tmCrf = tmRec
'                    'If ilRet <> BTRV_ERR_NONE Then
'                    '    Print #hmMsg, "Get by Key CRF Combo to update Feed Dates Failed" & Str$(ilRet) & " processing terminated"
'                    '    mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
'                    '    Screen.MousePointer = vbDefault
'                    '    ilRet = MsgBox("File in Use [Re-press Export], GetByKey Combine Crf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
'                    '    Exit Function
'                    'End If
'                    tmCrf.sAffFdStatus = "S" '"S"
'                    'tmCrf.iFeedDate(0) = ilTranDate0
'                    'tmCrf.iFeedDate(1) = ilTranDate1
'                    'gPackDateLong lmInputStartDate, tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1)
'                    mSetAffFdDate
'                    ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
'                Loop While ilRet = BTRV_ERR_CONFLICT
'                If ilRet <> BTRV_ERR_NONE Then
'                    Print #hmMsg, "Update CRF Combo to update Feed Dates Failed" & Str$(ilRet) & " processing terminated"
'                    mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
'                    Screen.MousePointer = vbDefault
'                    ilRet = MsgBox("File in Use [Re-press Export], Update Combine Crf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
'                    Exit Function
'                End If
'                ilCombineIndex = tgCombineCrf(ilCombineIndex).iCombineIndex
'            Loop
'            ilDuplIndex = tgSortCrf(ilCrfIndex).iDuplIndex
'            Do While ilDuplIndex >= 0
'                Do
'                    ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgDuplCrf(ilDuplIndex).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        Print #hmMsg, "Get CRF Duplicate to update Feed Dates Failed" & Str$(ilRet) & " processing terminated"
'                        mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
'                        Screen.MousePointer = vbDefault
'                        ilRet = MsgBox("File in Use [Re-press Export], GetDirect Dupl Crf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
'                        Exit Function
'                    End If
'                    'tmRec = tmCrf
'                    'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
'                    'tmCrf = tmRec
'                    'If ilRet <> BTRV_ERR_NONE Then
'                    '    Print #hmMsg, "Get by Key CRF Duplicate to update Feed Dates Failed" & Str$(ilRet) & " processing terminated"
'                    '    mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
'                    '    Screen.MousePointer = vbDefault
'                    '    ilRet = MsgBox("File in Use [Re-press Export], GetByKey Dupl Crf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
'                    '    Exit Function
'                    'End If
'                    tmCrf.sAffFdStatus = "S" '"S"
'                    'tmCrf.iFeedDate(0) = ilTranDate0
'                    'tmCrf.iFeedDate(1) = ilTranDate1
'                    'gPackDateLong lmInputStartDate, tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1)
'                    mSetAffFdDate
'                    ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
'                Loop While ilRet = BTRV_ERR_CONFLICT
'                If ilRet <> BTRV_ERR_NONE Then
'                    Print #hmMsg, "Update CRF Duplicate to update Feed Dates Failed" & Str$(ilRet) & " processing terminated"
'                    mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
'                    Screen.MousePointer = vbDefault
'                    ilRet = MsgBox("File in Use [Re-press Export], Update Dupl Crf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
'                    Exit Function
'                End If
'                ilDuplIndex = tgDuplCrf(ilDuplIndex).iDuplIndex
'            Loop
'        End If
    Next ilLoop
    ilRet = btrEndTrans(hmCrf)
    DoEvents
    'For ilLoop = 0 To UBound(tgDuplCrf) - 1 Step 1
    '    Do
    '        ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgDuplCrf(ilLoop).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
    '        If ilRet = BTRV_ERR_NONE Then
    '            tmCrf.sAffFdStatus = "S" '"S"
    '            'tmCrf.iFeedDate(0) = ilTranDate0
    '            'tmCrf.iFeedDate(1) = ilTranDate1
    '            gPackDateLong lmInputStartDate, tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1)
    '            ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
    '        End If
    '    Loop While ilRet = BTRV_ERR_CONFLICT
    'Next ilLoop
    mExpRot = True
    Exit Function
'cmcExportErr:
'    ilRet = Err.Number
'    Resume Next
'cmcExportCopyHeader:
'    ''If ilLineNo + ilNoCopyLines > 52 Then
'    'If ilPrtFirstCopy Then
'    '    ilPrtFirstCopy = False
'    '    If ilPageNo = 0 Then
'    '        'slRecord = ""
'    '        If Not mExportLine(slBlank, ilLineNo) Then
'    '            Exit Function
'    '        End If
'    '    Else
'    '        slCopyHeader = Chr(12)  'Form Feed
'    '        If Not mExportLine(slCopyHeader, ilLineNo) Then
'    '            Exit Function
'    '        End If
'    '    End If
'    '    ilPageNo = ilPageNo + 1
'    '    ilLineNo = 0
'    '    slCopyHeader = " "
'    '    Do While Len(slCopyHeader) < 35
'    '        slCopyHeader = slCopyHeader & " "
'    '    Loop
'    '    slCopyHeader = slCopyHeader & Trim$(tgSpf.sGClient)
'    '    If Not mExportLine(slCopyHeader, ilLineNo) Then
'    '        Exit Function
'    '    End If
'    '    slCopyHeader = " "
'    '    Do While Len(slCopyHeader) < 35
'    '        slCopyHeader = slCopyHeader & " "
'    '    Loop
'    '    slCopyHeader = slCopyHeader & slVehName
'    '    If Not mExportLine(slCopyHeader, ilLineNo) Then
'    '        Exit Function
'    '    End If
'    '    slCopyHeader = " "
'    '    Do While Len(slCopyHeader) < 35
'    '        slCopyHeader = slCopyHeader & " "
'    '    Loop
'    '    'slCopyHeader = slCopyHeader & "Commercial Feed # " & Trim$(tlStnInfo.sSiteID) & "-" & smFeedNo
'    '    If tlStnInfo.sType = "G" Then
'    '        slCopyHeader = slCopyHeader & "Commercial Feed # " & Trim$(tlStnInfo.sFileName) & " (" & tlStnInfo.sStnFdCode & ")"
'    '    Else
'    '        slCopyHeader = slCopyHeader & "Commercial Feed # " & Trim$(tlStnInfo.sFileName) & " (" & tlStnInfo.sStnFdCode & ")"
'    '    End If
'    '    If Not mExportLine(slCopyHeader, ilLineNo) Then
'    '        Exit Function
'    '    End If
'    '    slCopyHeader = " "
'    '    Do While Len(slCopyHeader) < 35
'    '        slCopyHeader = slCopyHeader & " "
'    '    Loop
'    '    slCopyHeader = slCopyHeader & smTranDate & "  "
'    '    'slCopyHeader = slCopyHeader & "Page:"
'    '    'slStr = Trim$(Str$(ilPageNo))
'    '    'Do While Len(slStr) < 5
'    '    '    slStr = " " & slStr
'    '    'Loop
'    '    'slCopyHeader = slCopyHeader & slStr
'    '    If Not mExportLine(slCopyHeader, ilLineNo) Then
'    '        Exit Function
'    '    End If
'    '    slCopyHeader = ""
'    '    If Not mExportLine(slBlank, ilLineNo) Then
'    '        Exit Function
'    '    End If
'    '    slCopyHeader = "      Short Title                   Copy Active Dates"
'    '    If Not mExportLine(slCopyHeader, ilLineNo) Then
'    '        Exit Function
'    '    End If
'    '    slCopyHeader = "      ------------------------------"
'    '    If Not mExportLine(slCopyHeader, ilLineNo) Then
'    '        Exit Function
'    '    End If
'    '    slCopyHeader = "#    ISCI                  Creative Title                        Len  Sent Date"
'    '    If Not mExportLine(slCopyHeader, ilLineNo) Then
'    '        Exit Function
'    '    End If
'    '    If Not mExportLine(slBlank, ilLineNo) Then
'    '        Exit Function
'    '    End If
'    '    If (slPrevNewInv <> "") And (ilNoCopyLines = 1) Then
'    '        If Not mExportLine(slPrevNewInv, ilLineNo) Then
'    '            Exit Function
'    '        End If
'    '        slCopyHeader = "     "
'    '        slCopyHeader = slCopyHeader & "------------------------------"
'    '        If Not mExportLine(slCopyHeader, ilLineNo) Then
'    '            Exit Function
'    '        End If
'    '    End If
'    'End If
'    Return
'cmcExportRotHeader:
'    'If ilLineNo >= 52 Then
'    If ilPrtFirstRot Then
'        ilPrtFirstRot = False
'        If ilPageNo > 0 Then
'            If ((rbcInterface(0).Value) And (rbcFormat(1).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(3).Value)) Then
'                slRecord = Chr(12)  'Form Feed
'                If Not mExportLine(slRecord, ilLineNo, 5) Then
'                    Exit Function
'                End If
'            End If
'        End If
'        ilPageNo = ilPageNo + 1
'        ilLineNo = 0
'        'slRecord = "-"
'        'Do While Len(slRecord) < 78
'        '    slRecord = slRecord & "-"
'        'Loop
'        'If Not mExportLine(slRecord, ilLineNo) Then
'        '    Exit Function
'        'End If
'        'If Not mExportLine(slRecord, ilLineNo) Then
'        '    Exit Function
'        'End If
'        ''slRecord = " "
'        ''Do While Len(slRecord) < 68
'        ''    slRecord = slRecord & " "
'        ''Loop
'        ''slRecord = slRecord & "Page:"
'        ''slStr = Trim$(Str$(ilPageNo))
'        ''Do While Len(slStr) < 5
'        ''    slStr = " " & slStr
'        ''Loop
'        ''slRecord = slRecord & slStr
'        ''If Not mExportLine(slRecord, ilLineNo) Then
'        ''    Exit Function
'        ''End If
'        'slRecord = ""
'        'If Not mExportLine(slRecord, ilLineNo) Then
'        '    Exit Function
'        'End If
'        'If Not mExportLine(slRecord, ilLineNo) Then
'        '    Exit Function
'        'End If
'        'slRecord = Trim$(tgSpf.sGClient) & " " & slVehName & " Network Feed Instructions " & smTranDate
'        'If Not mExportLine(slRecord, ilLineNo) Then
'        '    Exit Function
'        'End If
'        If ((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
'            slRecord = smTranDate
'            Do While Len(slRecord) < 11
'                slRecord = slRecord & " "
'            Loop
'            If tlStnInfo.sType = "G" Then
'                slRecord = slRecord & UCase$(slVehName & " Network Commercial Instructions")
'            Else
'                slRecord = slRecord & Trim$(tlStnInfo.sCallLetter) & "-" & tlStnInfo.sBand & ", " & UCase$(slVehName & " Network Commercial Instructions")
'            End If
'            If Not mExportLine(slRecord, ilLineNo, 1) Then
'                Exit Function
'            End If
'        Else
'            slRecord = UCase$(Trim$(tgSpf.sGClient))
'            slRecord = slRecord & "          " & smTranDate
'            If Not mExportLine(slRecord, ilLineNo, 5) Then
'                Exit Function
'            End If
'            If Not mExportLine(slBlank, ilLineNo, 5) Then
'                Exit Function
'            End If
'            If Not mExportLine(slBlank, ilLineNo, 5) Then
'                Exit Function
'            End If
'            If tlStnInfo.sType = "G" Then
'                slRecord = UCase$(slVehName & " Network Commercial Instructions")
'            Else
'                slRecord = Trim$(tlStnInfo.sCallLetter) & "-" & tlStnInfo.sBand & ", " & UCase$(slVehName & " Network Commercial Instructions")
'            End If
'            If Not mExportLine(slRecord, ilLineNo, 5) Then
'                Exit Function
'            End If
'        End If
'        If Not mExportLine(slBlank, ilLineNo, 5) Then
'            Exit Function
'        End If
'        If Not mExportLine(slBlank, ilLineNo, 5) Then
'            Exit Function
'        End If
'    Else
'        If ilNewHdRot = True Then
'            ilNewHdRot = False
'            If Not mExportLine(slBlank, ilLineNo, 5) Then
'                Exit Function
'            End If
'            'slRecord = "-"
'            'Do While Len(slRecord) < 60
'            '    slRecord = slRecord & "-"
'            'Loop
'            'If Not mExportLine(slRecord, ilLineNo) Then
'            '    Exit Function
'            'End If
'            'If Not mExportLine(slBlank, ilLineNo) Then
'            '    Exit Function
'            'End If
'        End If
'    End If
'    Return
'cmcExportSendMsg:
'    ilRet = 0
'    On Error GoTo cmcExportErr:
'    hlMsg = FreeFile
'    slMsgFile = sgExportPath & slMsgFileName
'    Open slMsgFile For Input Access Read As hlMsg
'    If ilRet = 0 Then
'        Do
'            On Error GoTo cmcExportErr:
'            Line Input #hlMsg, slMsgLine
'            On Error GoTo 0
'            If (ilRet <> 0) Then    'Ctrl Z
'                Exit Do
'            End If
'            If Len(slMsgLine) > 0 Then
'                If (Asc(slMsgLine) = 26) Then    'Ctrl Z
'                    Exit Do
'                End If
'                ilPos = InStr(UCase$(slMsgLine), "XX/XX/XXXX")
'                If ilPos > 0 Then
'                    Mid$(slMsgLine, ilPos) = smTranDate
'                End If
'            End If
'            If ilMsgType = 0 Then
'                If Not mExportLine(slMsgLine, ilLineNo, -1) Then
'                    Exit Function
'                End If
'                '6/3/16: Replaced GoSub
'                'GoSub cmcExportCopyHeader
'                mExportCopyHeader
'            ElseIf ilMsgType = 1 Then
'                If Not mExportLine(slMsgLine, ilLineNo, 5) Then
'                    Exit Function
'                End If
'                '6/3/16: Replaced GoSub
'                'GoSub cmcExportRotHeader
'                If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
'                    Exit Function
'                End If
'            End If
'        Loop
'        Close hlMsg
'    End If
'    Return
'SpecialInstructions:
'    If Not mExportLine(slBlank, ilLineNo, 5) Then
'        Exit Function
'    End If
'    slRecord = "Instructions:"
'    Do While Len(slRecord) < 15
'        slRecord = slRecord & " "
'    Loop
'    If (tmCrf.iStartTime(0) <> 0) Or (tmCrf.iStartTime(1) <> 0) Or (tmCrf.iEndTime(0) <> 0) Or (tmCrf.iEndTime(1) <> 0) Then
'        'slRecord = "Air This Copy Between "
'        slRecord = slRecord & "Air this copy between "
'        gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slTime
'        slTime = UCase(slTime)
'        If slTime = "12AM" Then
'            slTime = "12M"
'        ElseIf slTime = "12PM" Then
'            slTime = "12N"
'        End If
'        'If slTime = "12M" Then
'        '    slTime = "12AM"
'        'ElseIf slTime = "12N" Then
'        '    slTime = "12PM"
'        'End If
'        slRecord = slRecord & slTime & " and "
'        gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slTime
'        slTime = UCase(slTime)
'        If slTime = "12AM" Then
'            slTime = "12M"
'        ElseIf slTime = "12PM" Then
'            slTime = "12N"
'        End If
'        'If slTime = "12M" Then
'        '    slTime = "12AM"
'        'ElseIf slTime = "12N" Then
'        '    slTime = "12PM"
'        'End If
'        slRecord = slRecord & slTime
'        If Not mExportLine(slRecord, ilLineNo, 5) Then
'            Exit Function
'        End If
'        '6/3/16: Replaced GoSub
'        'GoSub cmcExportRotHeader
'        If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
'            Exit Function
'        End If
'        slRecord = " "
'        Do While Len(slRecord) < 15
'            slRecord = slRecord & " "
'        Loop
'    End If
'    ilDayReq = False
'    For ilDay = 0 To 6 Step 1
'        ilDayOn(ilDay) = False
'    Next ilDay
'    If llRotEndDate - llRotStartDate >= 6 Then
'        For ilDay = 0 To 6 Step 1
'            If tmCrf.sDay(ilDay) <> "Y" Then
'                ilDayReq = True
'                ilDayOn(ilDay) = False
'            Else
'                ilDayOn(ilDay) = True
'            End If
'        Next ilDay
'    Else
'        For llDate = llRotStartDate To llRotEndDate Step 1
'            ilDay = gWeekDayLong(llDate)
'            If tmCrf.sDay(ilDay) <> "Y" Then
'                'ilDayReq = True
'                ilDayOn(ilDay) = False
'            Else
'                ilDayOn(ilDay) = True
'            End If
'        Next llDate
'        ilDayReq = True
'    End If
'    If ilDayReq Then
'        slStr = ""
'        If (ilDayOn(0) = True) And (ilDayOn(1) = True) And (ilDayOn(2) = True) And (ilDayOn(3) = True) And (ilDayOn(4) = True) And (ilDayOn(5) = False) And (ilDayOn(6) = False) Then
'            slStr = "Mon thru Fri"
'        ElseIf (ilDayOn(0) = False) And (ilDayOn(1) = False) And (ilDayOn(2) = False) And (ilDayOn(3) = False) And (ilDayOn(4) = False) And (ilDayOn(5) = True) And (ilDayOn(6) = True) Then
'            slStr = "Sat and Sun"
'        Else
'            For ilDay = 0 To 6 Step 1
'                If ilDayOn(ilDay) = True Then
'                    Select Case ilDay
'                        Case 0
'                            slStr = slStr & " Mon"
'                        Case 1
'                            slStr = slStr & " Tue"
'                        Case 2
'                            slStr = slStr & " Wed"
'                        Case 3
'                            slStr = slStr & " Thu"
'                        Case 4
'                            slStr = slStr & " Fri"
'                        Case 5
'                            slStr = slStr & " Sat"
'                        Case 6
'                            slStr = slStr & " Sun"
'                    End Select
'                End If
'            Next ilDay
'        End If
'        slRecord = slRecord & "Air this copy on" & slStr
'        If Not mExportLine(slRecord, ilLineNo, 5) Then
'            Exit Function
'        End If
'        '6/3/16: Replaced GoSub
'        'GoSub cmcExportRotHeader
'        If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
'            Exit Function
'        End If
'        slRecord = " "
'        Do While Len(slRecord) < 15
'            slRecord = slRecord & " "
'        Loop
'    End If
'    'If (Trim$(tmCrf.sZone) <> "") Then
'    If (Trim$(tmCrf.sZone) <> "") And (Trim$(tmCrf.sZone) <> "R") Then
'        Select Case Trim$(tmCrf.sZone)
'            Case "EST"
'                slStr = "EASTERN TIME ZONE"
'            Case "CST"
'                slStr = "CENTRAL TIME ZONE"
'            Case "MST"
'                slStr = "MOUNTAIN TIME ZONE"
'            Case "PST"
'                slStr = "PACIFIC TIME ZONE"
'            Case "R"
'                tmRafSrchKey.lCode = tlStnInfo.lRafCode
'                ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'                If ilRet = BTRV_ERR_NONE Then
'                    slStr = Trim$(tmRaf.sName)
'                Else
'                    slStr = ""
'                End If
'        End Select
'        slRecord = slRecord & "Air this copy only if you are in the " & slStr
'        If Not mExportLine(slRecord, ilLineNo, 5) Then
'            Exit Function
'        End If
'        '6/3/16: Replaced GoSub
'        'GoSub cmcExportRotHeader
'        If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
'            Exit Function
'        End If
'        slRecord = " "
'        Do While Len(slRecord) < 15
'            slRecord = slRecord & " "
'        Loop
'    End If
'    If tmCrf.lCsfCode <> 0 Then
'        tmCsfSrchKey.lCode = tmCrf.lCsfCode
'        tmCsf.sComment = ""
'        imCsfRecLen = Len(tmCsf) '5011
'        ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            'Output 70 characters per line
'            'If tmCsf.iStrLen > 0 Then
'            slStr = gStripChr0(tmCsf.sComment)
'            If slStr <> "" Then
'                slComment = slStr 'Trim$(Left$(tmCsf.sComment, tmCsf.iStrLen))
'                If Not ilIncludeNewMessage Then
'                    If InStr(1, slComment, "Revised For New", vbTextCompare) > 0 Then
'                        slComment = ""
'                    End If
'                End If
'                Do While Len(slComment) > 0
'                    'Repeat all CR/LF with Space/LF
'                    For ilPos = 1 To Len(slComment) Step 1
'                        If Asc(Mid$(slComment, ilPos, 1)) = Asc(sgCR) Then
'                            Mid$(slComment, ilPos, 1) = " "
'                        End If
'                    Next ilPos
'                    ilPos = InStr(slComment, " ")
'                    If ilPos > 0 Then
'                        If Len(slRecord) + ilPos - 1 > 70 Then
'                            If Not mExportLine(slRecord, ilLineNo, 5) Then
'                                Exit Function
'                            End If
'                            '6/3/16: Replaced GoSub
'                            'GoSub cmcExportRotHeader
'                            If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
'                                Exit Function
'                            End If
'                            slRecord = " "
'                            Do While Len(slRecord) < 15
'                                slRecord = slRecord & " "
'                            Loop
'                        End If
'                        slRecord = slRecord & Left$(slComment, ilPos)
'                        slComment = right$(slComment, Len(slComment) - ilPos)
'                        If (Asc(slComment) = Asc(sgLF)) Then
'                            If Not mExportLine(slRecord, ilLineNo, 5) Then
'                                Exit Function
'                            End If
'                            '6/3/16: Replaced GoSub
'                            'GoSub cmcExportRotHeader
'                            If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
'                                Exit Function
'                            End If
'                            slRecord = " "
'                            Do While Len(slRecord) < 15
'                                slRecord = slRecord & " "
'                            Loop
'                            slComment = right$(slComment, Len(slComment) - 1)
'                        End If
'                    Else
'                        If Len(slRecord) + Len(slComment) > 70 Then
'                            If Not mExportLine(slRecord, ilLineNo, 5) Then
'                                Exit Function
'                            End If
'                            '6/3/16: Replaced GoSub
'                            'GoSub cmcExportRotHeader
'                            If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
'                                Exit Function
'                            End If
'                            slRecord = " "
'                            Do While Len(slRecord) < 15
'                                slRecord = slRecord & " "
'                            Loop
'                        End If
'                        slRecord = slRecord & slComment
'                        If Not mExportLine(slRecord, ilLineNo, 5) Then
'                            Exit Function
'                        End If
'                        '6/3/16: Replaced GoSub
'                        'GoSub cmcExportRotHeader
'                        If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
'                            Exit Function
'                        End If
'                        slRecord = " "
'                        Do While Len(slRecord) < 15
'                            slRecord = slRecord & " "
'                        Loop
'                        slComment = ""
'                        Exit Do
'                    End If
'                Loop
'            End If
'        End If
'    End If
'    If InStr(1, slRecord, "Instructions:", 1) <= 0 Then
'        If Not mExportLine(slBlank, ilLineNo, 5) Then
'            Exit Function
'        End If
'    End If
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mExpSpots                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Export Recs                    *
'*                                                     *
'*******************************************************
Private Function mExpSpots(ilLCFType As Integer, sLCP As String, ilCallCode As Integer, slSDate As String, slEDate As String, slStartTime As String, slEndTime As String, ilEvtType() As Integer) As Integer
'
'   iRet = mExpSpots()
'   Where:
'       iRet (O)- True if record exported,
'                 False if error
'
    Dim ilLoop As Integer
    Dim ilRet As Integer    'Return status
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilSsfDate0 As Integer
    Dim ilSsfDate1 As Integer
    Dim ilEvt As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim ilSpot As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilTerminated As Integer
    Dim ilStartTime0 As Integer
    Dim ilStartTime1 As Integer
    'Dim ilSeqNo As Integer
    Dim ilVpfIndex As Integer
    Dim ilVehCode As Integer
    Dim ilVeh As Integer
    Dim ilDlfDate0 As Integer
    Dim ilDlfDate1 As Integer
    Dim ilDlfFound As Integer
    Dim ilVlfDate0 As Integer
    Dim ilVlfDate1 As Integer
    Dim ilSIndex As Integer
    Dim slSsfDate As String
    Dim ilBreakNo As Integer    'Reset to zero for each program
    Dim ilPositionNo As Integer 'Reset to zero for each avail
    'Spot summary
    Dim tlSsfSrchKey As SSFKEY0 'SSF key record image
    Dim ilSsfRecLen As Integer  'SSF record length
    Dim llEvtTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilWithinTime As Integer
    Dim ilEvtFdIndex As Integer
    Dim ilAirHour As Integer
    Dim ilLocalHour As Integer
    Dim slAdjDate As String
    Dim ilType As Integer   '1=Prog; 2=Comment; 3=Avail; 4=Spot
    Dim ilEvtRet As Integer
    Dim slZone As String
    Dim ilSeqNo As Integer      'Used to order spots within an avail
    Dim ilSSFType As Integer
    ReDim tlLLC(0 To 0) As LLC  'Image
    ReDim ilVehicle(0 To 1) As Integer
    Dim ilVff As Integer
    Dim blBypassZeroUnits As Boolean

    ilVehicle(0) = ilCallCode
    llStartTime = CLng(gTimeToCurrency(slStartTime, False))
    llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
    ilSeqNo = 0

    'If slType <> "O" Then
    '    ilSSFType = 1
    'Else
        ilSSFType = 0
    'End If

    tmEVef.iCode = 0
    llSDate = gDateValue(slSDate)
    llEDate = gDateValue(slEDate)
    For llDate = llSDate To llEDate Step 1
        If imTerminate Then
            mExpSpots = False
            Exit Function
        End If
        ilWithinTime = False
        slDate = Format$(llDate, "m/d/yy")
        ilDay = gWeekDayStr(slDate)
        gPackDate slDate, ilLogDate0, ilLogDate1
        For ilVeh = LBound(ilVehicle) To UBound(ilVehicle) - 1 Step 1
            ilVehCode = ilVehicle(ilVeh)
            If ilVehCode <> tmEVef.iCode Then
                'tmVefSrchKey.iCode = ilVehCode
                'ilRet = btrGetEqual(hmVef, tmEVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                ilRet = gBinarySearchVef(ilVehCode)
                If ilRet <> -1 Then
                    tmEVef = tgMVef(ilRet)
                Else
                    mExpSpots = False
                    Exit Function
                End If
                ilVpfIndex = -1
                'For ilLoop = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
                '    If tmEVef.iCode = tgVpf(ilLoop).iVefKCode Then
                    ilLoop = gBinarySearchVpf(tmEVef.iCode)
                    If ilLoop <> -1 Then
                        ilVpfIndex = ilLoop
                '        Exit For
                    End If
                'Next ilLoop
                If ilVpfIndex = -1 Then
                    mExpSpots = False
                    Exit Function
                End If
            End If
            ilDlfFound = False
            'If (((tmEVef.sType = "A") Or ((tmEVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(1) <> 0)))) Then
            If (((tmEVef.sType = "A") Or ((tmEVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(0) <> 0)))) Then
                'Obtain delivery records for date
                If (ilDay >= 0) And (ilDay <= 4) Then
                    slDay = "0"
                ElseIf ilDay = 5 Then
                    slDay = "6"
                Else
                    slDay = "7"
                End If
                'Obtain the start date of DLF
                tmDlfSrchKey.iVefCode = ilVehCode
                tmDlfSrchKey.sAirDay = slDay
                tmDlfSrchKey.iStartDate(0) = ilLogDate0
                tmDlfSrchKey.iStartDate(1) = ilLogDate1
                tmDlfSrchKey.iAirTime(0) = 0
                tmDlfSrchKey.iAirTime(1) = 6144 '24*256
                ilRet = btrGetLessOrEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) Then
                    ilDlfDate0 = tmDlf.iStartDate(0)
                    ilDlfDate1 = tmDlf.iStartDate(1)
                    ilDlfFound = True
                Else
                    ilDlfDate0 = 0
                    ilDlfDate1 = 0
                End If
                'Obtain the start date of VLF
                If tmEVef.sType = "A" Then
                    ilVlfDate0 = 0
                    ilVlfDate1 = 0
                    tmVlfSrchKey1.iAirCode = ilVehCode
                    tmVlfSrchKey1.iAirDay = Val(slDay)
                    tmVlfSrchKey1.iEffDate(0) = ilLogDate0
                    tmVlfSrchKey1.iEffDate(1) = ilLogDate1
                    tmVlfSrchKey1.iAirTime(0) = 0
                    tmVlfSrchKey1.iAirTime(1) = 6144    '24*256
                    tmVlfSrchKey1.iAirPosNo = 32000
                    tmVlfSrchKey1.iAirSeq = 32000
                    ilRet = btrGetLessOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) 'And (tmVlf.iAirDay = Val(slDay))
                        If (tmVlf.iAirDay = Val(slDay)) Then
                            ilTerminated = False
                            If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                    ilTerminated = True
                                End If
                            End If
                            If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                ilVlfDate0 = tmVlf.iEffDate(0)
                                ilVlfDate1 = tmVlf.iEffDate(1)
                                Exit Do
                            End If
                        End If
                        ilRet = btrGetPrevious(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
            End If
            
            blBypassZeroUnits = False
            If (tmEVef.sType = "A") Then
                ilVff = gBinarySearchVff(tmEVef.iCode)
                If ilVff <> -1 Then
                    If tgVff(ilVff).sHonorZeroUnits = "Y" Then
                        blBypassZeroUnits = True
                    End If
                End If
            End If
            
            DoEvents
            'gObtainVlf hlVlf, ilVehCode, llDate, tlVlf0(), tlVlf5(), tlVlf6()
            ilDay = gWeekDayStr(slDate)
            gPackDate slDate, ilLogDate0, ilLogDate1
            ilSsfRecLen = Len(tgSsf(0)) 'Max size of variable length record
            ilSsfDate0 = ilLogDate0
            ilSsfDate1 = ilLogDate1
            tlSsfSrchKey.iType = ilSSFType
            tlSsfSrchKey.iVefCode = ilVehCode
            tlSsfSrchKey.iDate(0) = ilSsfDate0
            tlSsfSrchKey.iDate(1) = ilSsfDate1
            tlSsfSrchKey.iStartTime(0) = 0
            tlSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetEqual(hmSsf, tgSsf(0), ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If (ilRet <> BTRV_ERR_NONE) Or (tgSsf(0).iType <> ilSSFType) Or (tgSsf(0).iVefCode <> ilVehCode) Or (tgSsf(0).iDate(0) <> ilSsfDate0) Or (tgSsf(0).iDate(1) <> ilSsfDate1) Then
                'If airing- then use first Ssf prior to date defined
                If tmEVef.sType = "A" Then
                    ilSsfDate0 = 0
                    ilSsfDate1 = 0
                    ilSsfRecLen = Len(tgSsf(0)) 'Max size of variable length record
                    tlSsfSrchKey.iType = ilSSFType
                    tlSsfSrchKey.iVefCode = ilVehCode
                    tlSsfSrchKey.iDate(0) = ilLogDate0
                    tlSsfSrchKey.iDate(1) = ilLogDate1
                    tlSsfSrchKey.iStartTime(0) = 0
                    tlSsfSrchKey.iStartTime(1) = 6144   '24*256
                    ilRet = gSSFGetLessOrEqual(hmSsf, tgSsf(0), ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    Do While (ilRet = BTRV_ERR_NONE) And (tgSsf(0).iType = ilSSFType) And (tgSsf(0).iVefCode = ilVehCode)
                        gUnpackDate tgSsf(0).iDate(0), tgSsf(0).iDate(1), slSsfDate
                        If (ilDay = gWeekDayStr(slSsfDate)) And (tgSsf(0).iStartTime(0) = 0) And (tgSsf(0).iStartTime(1) = 0) Then
                            ilSsfDate0 = tgSsf(0).iDate(0)
                            ilSsfDate1 = tgSsf(0).iDate(1)
                            Exit Do
                        End If
                        ilSsfRecLen = Len(tgSsf(0)) 'Max size of variable length record
                        ilRet = gSSFGetPrevious(hmSsf, tgSsf(0), ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
            End If
            DoEvents
            If (ilRet = BTRV_ERR_NONE) And (tgSsf(0).iType = ilSSFType) And (tgSsf(0).iVefCode = ilVehCode) Then
                gUnpackDate ilSsfDate0, ilSsfDate1, slSsfDate
                'If (ilEvtType(0) = True) Or (ilEvtType(10) = True) Or (ilEvtType(11) = True) Or (ilEvtType(12) = True) Or (ilEvtType(13) = True) Or (ilEvtType(14) = True) Then
                '    ReDim tlLLC(0 To 0) As LLC  'Image
                '    ilEvtRet = gBuildEventDay(slType, slCP, ilVehCode, slSsfDate, slStartTime, slEndTime, ilEvtType(), tlLLC())
                'Else
                '    If ilEvtType(1) = True Then
                '       ReDim tlLLC(0 To 0) As LLC  'Image
                '       ilEvtRet = gBuildEventDay(slType, slCP, ilVehCode, slSsfDate, slStartTime, slEndTime, ilEvtType(), tlLLC())
                '    End If
                'End If
                ReDim tlLLC(0 To 0) As LLC  'Image
                ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, slSsfDate, slStartTime, slEndTime, ilEvtType(), tlLLC())
                ilBreakNo = 0
                ilPositionNo = 0
                Do While (ilRet = BTRV_ERR_NONE) And (tgSsf(0).iType = ilSSFType) And (tgSsf(0).iVefCode = ilVehCode) And (tgSsf(0).iDate(0) = ilSsfDate0) And (tgSsf(0).iDate(1) = ilSsfDate1)
                    'Loop thru Ssf and move records to tmOdf
                    ilEvt = 1
                    Do While ilEvt <= tgSsf(0).iCount
                       LSet tmProg = tgSsf(0).tPas(ADJSSFPASBZ + ilEvt)
                        If (tmProg.iRecType = 1) Or ((tmProg.iRecType >= 2) And (tmProg.iRecType <= 9)) Then
                            gUnpackTimeLong tmProg.iStartTime(0), tmProg.iStartTime(1), False, llEvtTime
                            If llEvtTime > llEndTime Then
                                ilWithinTime = False
                                Exit Do
                            End If
                            If llEvtTime >= llStartTime Then
                                ilWithinTime = True
                            End If
                        End If
                        If blBypassZeroUnits And (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then
                           LSet tmAvail = tgSsf(0).tPas(ADJSSFPASBZ + ilEvt)
                            If (tmAvail.iAvInfo And &H1F <= 0) Or (tmAvail.iLen <= 0) Then
                                tmProg.iRecType = 0
                            End If
                        End If
                        If tmProg.iRecType = 1 Then
                            ilBreakNo = 0
                            ilPositionNo = 0
                        ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then
                            ilBreakNo = ilBreakNo + 1
                            ilPositionNo = 0
                        End If
                        ilEvtFdIndex = -1
                        If ilWithinTime Then
                            If tmProg.iRecType = 1 Then    'Program
                                'For ilLoop = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                                '    'Match start time and length
                                '    If tlLLC(ilLoop).iEtfCode = 1 Then
                                '        gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                '        If (ilStartTime0 = tmProg.iStartTime(0)) And (ilStartTime1 = tmProg.iStartTime(1)) Then
                                '            gAddTimeLength tlLLC(ilLoop).sStartTime, tlLLC(ilLoop).sLength, "A", "1", slTime
                                '            gPackTime slTime, ilEndTime0, ilEndTime1
                                '            If (ilEndTime0 = tmProg.iEndTime(0)) And (ilEndTime1 = tmProg.iEndTime(1)) Then
                                '                ilEvtFdIndex = ilLoop
                                '                tlLLC(ilLoop).iEtfCode = -1 'Remove event
                                '                Exit For
                                '            End If
                                '        End If
                                '    End If
                                'Next ilLoop
                            ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then 'Avail
                               LSet tmAvail = tgSsf(0).tPas(ADJSSFPASBZ + ilEvt)
                                For ilLoop = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                                    'Match start time and length
                                    If (tlLLC(ilLoop).iEtfCode >= 2) And (tlLLC(ilLoop).iEtfCode <= 9) Then
                                        gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                        If (ilStartTime0 = tmAvail.iTime(0)) And (ilStartTime1 = tmAvail.iTime(1)) Then
                                            ilEvtFdIndex = ilLoop
                                            'Loop on spots, then add conflicting spots
                                            If (tmEVef.sType = "A") Then
                                                tmVlfSrchKey1.iAirCode = ilVehCode
                                                tmVlfSrchKey1.iAirDay = Val(slDay)
                                                tmVlfSrchKey1.iEffDate(0) = ilVlfDate0
                                                tmVlfSrchKey1.iEffDate(1) = ilVlfDate1
                                                tmVlfSrchKey1.iAirTime(0) = tmAvail.iTime(0)
                                                tmVlfSrchKey1.iAirTime(1) = tmAvail.iTime(1)
                                                tmVlfSrchKey1.iAirPosNo = 0
                                                tmVlfSrchKey1.iAirSeq = 1
                                                ilRet = btrGetGreaterOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                                Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) And (tmVlf.iAirDay = Val(slDay)) And (tmVlf.iEffDate(0) = ilVlfDate0) And (tmVlf.iEffDate(1) = ilVlfDate1) And (tmVlf.iAirTime(0) = tmAvail.iTime(0)) And (tmVlf.iAirTime(1) = tmAvail.iTime(1))
                                                    ilTerminated = False
                                                    If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                                        If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                                            ilTerminated = True
                                                        End If
                                                    End If
                                                    If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                                        If (tmCTSsf.iType <> ilSSFType) Or (tmCTSsf.iVefCode <> tmVlf.iSellCode) Or (tmCTSsf.iDate(0) <> ilLogDate0) Or (tmCTSsf.iDate(1) <> ilLogDate1) Then
                                                            ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                            tlSsfSrchKey.iType = ilSSFType
                                                            tlSsfSrchKey.iVefCode = tmVlf.iSellCode
                                                            tlSsfSrchKey.iDate(0) = ilLogDate0
                                                            tlSsfSrchKey.iDate(1) = ilLogDate1
                                                            tlSsfSrchKey.iStartTime(0) = 0
                                                            tlSsfSrchKey.iStartTime(1) = 0
                                                            ilRet = gSSFGetEqual(hmCTSsf, tmCTSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                        End If
                                                        Do While (ilRet = BTRV_ERR_NONE) And (tmCTSsf.iType = ilSSFType) And (tmCTSsf.iVefCode = tmVlf.iSellCode) And (tmCTSsf.iDate(0) = ilLogDate0) And (tmCTSsf.iDate(1) = ilLogDate1)
                                                            For ilSIndex = 1 To tmCTSsf.iCount Step 1
                                                                tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSIndex)
                                                                If ((tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9)) Then
                                                                    If (tmAvailTest.iTime(0) = tmVlf.iSellTime(0)) And (tmAvailTest.iTime(1) = tmVlf.iSellTime(1)) Then
                                                                        For ilSpot = 1 To tmAvailTest.iNoSpotsThis Step 1
                                                                            ilType = 4
                                                                           LSet tmSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilSpot + ilSIndex)
                                                                            tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                            If ilRet = BTRV_ERR_NONE Then
                                                                                tmChfSrchKey.lCode = tmSdf.lChfCode
                                                                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                                                If ilRet = BTRV_ERR_NONE Then
                                                                                    ilPositionNo = ilPositionNo + 1
                                                                                    slZone = "EST"  'Use EST as standard, if not found, use OTH
                                                                                    If ilDlfFound Then
                                                                                        'llCifCode = mObtainCifCode(tmSdf, slZone, hmTzf, ilOther)
                                                                                        'ilCrfVefCode = gGetCrfVefCode(hmClf, tmSdf, ilCrfVefCode, ilPkgVefCode)
                                                                                        'llCrfCsfCode = mObtainCrfCsfCode(tmSdf, slZone, hmCrf, hmTzf, ilCrfVefCode, ilPkgVefCode, tmCrf)
                                                                                        ''Remove comment
                                                                                        'llCrfCsfCode = 0
                                                                                        tmDlfSrchKey.iVefCode = ilVehCode
                                                                                        tmDlfSrchKey.sAirDay = slDay
                                                                                        tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                                        tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                                        tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                                        tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                                        ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                        Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                                                            ilTerminated = False
                                                                                            If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                                                                                ilTerminated = True
                                                                                            Else
                                                                                                If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                                                    If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                                                        ilTerminated = True
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                            If Not ilTerminated Then
                                                                                                If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                                                                                    If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                                                                                                        tmDlf.iMnfFeed = 0
                                                                                                        '6/3/16: Replaced GoSub
                                                                                                        'GoSub lProcSpot
                                                                                                        mProcSpot ilVpfIndex, ilLogDate0, ilLogDate1, ilAirHour, ilLocalHour, slAdjDate, ilDlfFound, ilSeqNo, ilRet
                                                                                                        DoEvents
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                            ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                                        Loop
                                                                                    Else
                                                                                        ''Assign copy for all zones
                                                                                        'ilNoZones = 0
                                                                                        'For ilZone = 1 To 5 Step 1
                                                                                        '    Select Case ilZone
                                                                                        '        Case 1
                                                                                        '            slZone = "EST"
                                                                                        '        Case 2
                                                                                        '            slZone = "MST"
                                                                                        '        Case 3
                                                                                        '            slZone = "CST"
                                                                                        '        Case 4
                                                                                        '            slZone = "PST"
                                                                                        '        Case 5
                                                                                        '            slZone = "Oth"
                                                                                        '    End Select
                                                                                        '    'llCifCode = mObtainCifCode(tmSdf, slZone, hmTzf, ilOther)
                                                                                        '    'ilCrfVefCode = gGetCrfVefCode(hmClf, tmSdf, ilCrfVefCode, ilPkgVefCode)
                                                                                        '    'llCrfCsfCode = mObtainCrfCsfCode(tmSdf, slZone, hmCrf, hmTzf, ilCrfVefCode, ilPkgVefCode, tmCrf)
                                                                                            tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                                            tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                                            tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                                            tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                                        '    If (ilOther) And (ilNoZones = 0) Then
                                                                                                tmDlf.sZone = ""
                                                                                        '    Else
                                                                                        '        tmDlf.sZone = slZone
                                                                                        '        If ilNoZones = 0 Then
                                                                                        '            ilType = 4
                                                                                        '            ilNoZones = 1
                                                                                        '        Else
                                                                                        '            ilType = 5
                                                                                        '            ilNoZones = ilNoZones + 1
                                                                                        '        End If
                                                                                        '    End If
                                                                                            '6/3/16: Replaced GoSub
                                                                                            'GoSub lProcSpot
                                                                                            mProcSpot ilVpfIndex, ilLogDate0, ilLogDate1, ilAirHour, ilLocalHour, slAdjDate, ilDlfFound, ilSeqNo, ilRet
                                                                                            DoEvents
                                                                                        '    If (ilOther) Or (ilNoZones = 4) Then
                                                                                        '        Exit For
                                                                                        '    End If
                                                                                        'Next ilZone
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        Next ilSpot
                                                                        Exit Do
                                                                    End If
                                                                End If
                                                            Next ilSIndex
                                                            ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                            ilRet = gSSFGetNext(hmCTSsf, tmCTSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        Loop
                                                    End If
                                                    ilRet = btrGetNext(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                Loop
                                            Else
                                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                                    ilType = 4
                                                    ilEvt = ilEvt + 1
                                                   LSet tmSpot = tgSsf(0).tPas(ADJSSFPASBZ + ilEvt)
                                                    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    If ilRet = BTRV_ERR_NONE Then
                                                        tmChfSrchKey.lCode = tmSdf.lChfCode
                                                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            ilPositionNo = ilPositionNo + 1
                                                            slZone = "EST"  'Use EST as standard, if not found, use OTH
                                                            If ilDlfFound Then
                                                                'llCifCode = mObtainCifCode(tmSdf, slZone, hmTzf, ilOther)
                                                                'ilCrfVefCode = gGetCrfVefCode(hmClf, tmSdf, ilCrfVefCode, ilPkgVefCode)
                                                                'llCrfCsfCode = mObtainCrfCsfCode(tmSdf, slZone, hmCrf, hmTzf, ilCrfVefCode, ilPkgVefCode, tmCrf)
                                                                ''Remove Comment
                                                                'llCrfCsfCode = 0
                                                                'Obtain delivery entry to see is avail is sent
                                                                tmDlfSrchKey.iVefCode = ilVehCode
                                                                tmDlfSrchKey.sAirDay = slDay
                                                                tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                                    ilTerminated = False
                                                                    If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                                                        ilTerminated = True
                                                                    Else
                                                                        If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                            If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                                ilTerminated = True
                                                                            End If
                                                                        End If
                                                                    End If
                                                                    If Not ilTerminated Then
                                                                        If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                                                            If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                                                                                tmDlf.iMnfFeed = 0
                                                                                '6/3/16: Replaced GoSub
                                                                                'GoSub lProcSpot
                                                                                mProcSpot ilVpfIndex, ilLogDate0, ilLogDate1, ilAirHour, ilLocalHour, slAdjDate, ilDlfFound, ilSeqNo, ilRet
                                                                                DoEvents
                                                                            End If
                                                                        End If
                                                                    End If
                                                                    ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                Loop
                                                            Else
                                                                ''Assign copy for all zones
                                                                'ilNoZones = 0
                                                                'For ilZone = 1 To 5 Step 1
                                                                '    Select Case ilZone
                                                                '        Case 1
                                                                '            slZone = "EST"
                                                                '        Case 2
                                                                '            slZone = "MST"
                                                                '        Case 3
                                                                '            slZone = "CST"
                                                                '        Case 4
                                                                '            slZone = "PST"
                                                                '        Case 5
                                                                '            slZone = "Oth"
                                                                '    End Select
                                                                '    'llCifCode = mObtainCifCode(tmSdf, slZone, hmTzf, ilOther)
                                                                '    'ilCrfVefCode = gGetCrfVefCode(hmClf, tmSdf, ilCrfVefCode, ilPkgVefCode)
                                                                '    'llCrfCsfCode = mObtainCrfCsfCode(tmSdf, slZone, hmCrf, hmTzf, ilCrfVefCode, ilPkgVefCode, tmCrf)
                                                                    tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                    tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                    tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                    tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                '    If (ilOther) And (ilNoZones = 0) Then
                                                                        tmDlf.sZone = ""
                                                                '    Else
                                                                '        tmDlf.sZone = slZone
                                                                '        If ilNoZones = 0 Then
                                                                '            ilType = 4
                                                                '            ilNoZones = 1
                                                                '        Else
                                                                '            ilType = 5
                                                                '            ilNoZones = ilNoZones + 1
                                                                '        End If
                                                                '    End If
                                                                    '6/3/16: Replaced GoSub
                                                                    'GoSub lProcSpot
                                                                    mProcSpot ilVpfIndex, ilLogDate0, ilLogDate1, ilAirHour, ilLocalHour, slAdjDate, ilDlfFound, ilSeqNo, ilRet
                                                                    DoEvents
                                                                '    If (ilOther) Or (ilNoZones = 4) Then
                                                                '        Exit For
                                                                '    End If
                                                                'Next ilZone
                                                            End If
                                                        End If
                                                    End If
                                                Next ilSpot
                                            End If
                                            tlLLC(ilLoop).iEtfCode = -1 'Remove event
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                            End If
                        End If
                        ''Any other events to be sent out
                        'If ilEvtFdIndex >= 0 Then
                        '    'Output all event until Program or Avail found
                        '    For ilLoop = ilEvtFdIndex + 1 To UBound(tlLLC) - 1 Step 1
                        '        'Match start time and length
                        '        If (tlLLC(ilLoop).iEtfCode > 9) Then
                        '            'Handle like program
                        '            ilType = 2
                        '            llEvtTime = CLng(gTimeToCurrency(tlLLC(ilLoop).sStartTime, False))
                        '            If llEvtTime > llEndTime Then
                        '                ilWithinTime = False
                        '                Exit Do
                        '            End If
                        '            If llEvtTime >= llStartTime Then
                        '                ilWithinTime = True
                        '            End If
                        '            tlLLC(ilLoop).iEtfCode = -1 'Remove event
                        '        Else
                        '            Exit For
                        '        End If
                        '    Next ilLoop
                        'End If
                        ilEvt = ilEvt + 1
                    Loop
                    ilSsfRecLen = Len(tgSsf(0)) 'Max size of variable length record
                    ilRet = gSSFGetNext(hmSsf, tgSsf(0), ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
        Next ilVeh
    Next llDate
    mExpSpots = True
    Exit Function

    Return

    Return
'lProcSpot:
'    tmCpr.iGenDate(0) = imGenDate(0)
'    tmCpr.iGenDate(1) = imGenDate(1)
'    'tmCpr.iGenTime(0) = imGenTime(0)
'    'tmCpr.iGenTime(1) = imGenTime(1)
'    gUnpackTimeLong imGenTime(0), imGenTime(1), False, tmCpr.lGenTime
'    'Air Vehicle
'    tmCpr.iVefCode = tmEVef.iCode
'    'EDAS Time Window
'    tmCpr.lHd1CefCode = 400
'    If tgVpf(ilVpfIndex).lEDASWindow > 0 Then
'        tmCpr.lHd1CefCode = tgVpf(ilVpfIndex).lEDASWindow
'    End If
'    'Air Date
'    gUnpackDate ilLogDate0, ilLogDate1, slAdjDate
'    '6/3/16:Replaced GoSub
'    'GoSub lProcAdjDate  'Adjust dates prior to adjusting seq numbers
'    mProcAdjDate ilAirHour, ilLocalHour, slAdjDate
'    gPackDate slAdjDate, tmCpr.iSpotDate(0), tmCpr.iSpotDate(1)
'    'Air Time
'    tmCpr.iSpotTime(0) = tmDlf.iLocalTime(0)
'    tmCpr.iSpotTime(1) = tmDlf.iLocalTime(1)
'    'Spot Length
'    tmCpr.iLen = tmSdf.iLen
'    'Sdf Code
'    tmCpr.lCntrNo = tmSdf.lCode
'    'If (ilDlfFound) Or (tmSdf.sPtType <> "3") Then
'    '    tmCpr.sZone = tmDlf.sZone
'    '    ilRet = mObtainCopy(tmDlf.sZone)
'    '    If ilRet Then
'    '        If Trim$(tmCif.sCut) = "" Then
'    '            tmCpr.sCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName)
'    '        Else
'    '            tmCpr.sCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut)
'    '        End If
'    '    Else
'    '        tmCpr.sCartNo = ""
'    '    End If
'    '    ilSeqNo = ilSeqNo + 1
'    '    tmCpr.iLineNo = ilSeqNo
'    '    ilRet = btrUpdate(hmCpr, tmCpr, imCprRecLen)
'    'Else
'    '    For ilZone = 1 To 4 Step 1
'    '        Select Case ilZone
'    '            Case 1
'    '                slZone = "EST"
'    '            Case 2
'    '                slZone = "MST"
'    '            Case 3
'    '                slZone = "CST"
'    '            Case 4
'    '                slZone = "PST"
'    '        End Select
'    '        tmCpr.sZone = slZone
'    '        ilRet = mObtainCopy(slZone)
'    '        If ilRet Then
'    '            If Trim$(tmCif.sCut) = "" Then
'    '                tmCpr.sCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName)
'    '            Else
'    '                tmCpr.sCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut)
'    '            End If
'    '        Else
'    '            tmCpr.sCartNo = ""
'    '        End If
'    '        ilSeqNo = ilSeqNo + 1
'    '        tmCpr.iLineNo = ilSeqNo
'    '        ilRet = btrUpdate(hmCpr, tmCpr, imCprRecLen)
'    '    Next ilZone
'    'End If
'    If ilDlfFound Then
'        tmCpr.sZone = tmDlf.sZone
'    Else
'        tmCpr.sZone = ""
'    End If
'    tmCpr.sCartNo = ""
'    ilSeqNo = ilSeqNo + 1
'    tmCpr.iLineNo = ilSeqNo
'    ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
'    'slRecord = ""
'    ''Selling Vehicle
'    'tmClfSrchKey.lChfCode = tmChf.lCode
'    'tmClfSrchKey.iLine = tmSdf.iLineNo
'    'tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
'    'tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
'    'ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'    'If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) And (tmClf.iLine = tmSdf.iLineNo) Then
'    '    tmVefSrchKey.iCode = tmClf.iVefCode
'    '    ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'    '    slRecord = slRecord & Trim$(tmSVef.sName) & ","
'    'Else
'    '    slRecord = slRecord & "Sell Veh Missing" & ","
'    'End If
'    ''Airing Vehicle
'    'slRecord = slRecord & Trim$(tmEVef.sName) & ","
'    ''Regional Name
'    '
'    ''Air Date
'    'gUnpackDate ilLogDate0, ilLogDate1, slAdjDate
'    'GoSub lProcAdjDate  'Adjust dates prior to adjusting seq numbers
'    'slRecord = slRecord & slAdjDate & ","
'    ''Air Time
'    'gUnpackTime tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), "A", "1", slTime
'    'slTime = Format$(gConvertTime(slTime), "hhmmss")
'    'slRecord = slRecord & slTime & ","
'    'ilRet = mObtainCopy(tmDlf.sZone)
'    ''Cart #
'    'If ilRet Then
'    '    If Trim$(tmCif.sCut) = "" Then
'    '        slRecord = slRecord & Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & ","
'    '    Else
'    '        slRecord = slRecord & Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut) & ","
'    '    End If
'    'Else
'    '    slRecord = slRecord & " " & ","
'    'End If
'    '
'    ''ISCI Code
'    'slRecord = slRecord & Trim$(tmCpf.sISCI) & ","
'    '
'    ''Advertiser
'    'If tmAdf.iCode <> tmChf.iAdfCode Then
'    '    tmAdfSrchKey.iCode = tmChf.iAdfCode
'    '    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    '    If ilRet <> BTRV_ERR_NONE Then
'    '        tmAdf.sName = "Advertiser Missing"
'    '    End If
'    'End If
'    'slRecord = slRecord & Trim$(tmAdf.sName) & ","
'    '
'    ''Product
'    'slRecord = slRecord & Trim$(tmChf.sProduct) & ","
'    ''Short Title
'    'slRecord = slRecord & Trim$(gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)) & ","
'    '
'    'Print #hmTo, slRecord
'    Return
'lProcAdjDate:
'    'Test if Air time is AM and Local Time is PM. If so, adjust date
'    ilAirHour = tmAvail.iTime(1) \ 256  'Obtain month
'    ilLocalHour = tmDlf.iLocalTime(1) \ 256  'Obtain month
'    If (ilAirHour < 6) And (ilLocalHour > 17) Then
'        'If monday convert to next sunday- this is wrong but the same spot
'        'runs each sunday (the spot should have show on the previous week sunday)
'        'If not monday, then subtract one day
'        If gWeekDayStr(slAdjDate) = 0 Then
'            slAdjDate = gObtainNextSunday(slAdjDate)
'        Else
'            slAdjDate = gDecOneDay(slAdjDate)
'        End If
'    End If
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFindVpfIndex                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Find Vpf index for specified    *
'*                     vehicle                         *
'*                                                     *
'*******************************************************
Private Function mFindVpfIndex(ilVefCode As Integer) As Integer
    Dim ilLoop As Integer
    For ilLoop = LBound(tmVpfInfo) To UBound(tmVpfInfo) - 1 Step 1
        If ilVefCode = tmVpfInfo(ilLoop).tVpf.iVefKCode Then
            mFindVpfIndex = ilLoop
            Exit Function
        End If
    Next ilLoop
    mFindVpfIndex = -1
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetStnInfo                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Station Information        *
'*                                                     *
'*******************************************************
Private Function mGetStnInfo(ilTest As Integer) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilVeh As Integer
    Dim ilStn As Integer
    Dim ilFound As Integer
    Dim ilUpper As Integer
    Dim ilFileID As Integer
    Dim ilLoop As Integer
    Dim ilVIndex As Integer
    Dim ilSameGroup As Integer
    Dim ilVpfIndex As Integer
    Dim ilNextLk As Integer
    Dim ilMatch As Integer
    Dim ilTran As Integer
    Dim ilEDAS As Integer
    Dim ilPledge As Integer
    Dim ilIndex As Integer

    ilRet = 0
    'On Error GoTo mGetStnInfoErr:
    'hmStationFile = FreeFile
    'Open smStationFile For Input Access Read As hmStationFile
    ilRet = gFileOpen(smStationFile, "Input Access Read", hmStationFile)
    If ilRet <> 0 Then
        Close hmStationFile
        If ilTest Then
            ''MsgBox "Open " & smStationFile & ", Error #" & str$(ilRet), vbExclamation, "Open Error"
            If igExportType < 1 Then gAutomationAlertAndLogHandler "Open '" & smStationFile & "', Error #" & str$(ilRet) & " - " & Error(ilRet) 'log
            gAutomationAlertAndLogHandler "Open '" & smStationFile & "', Error #" & str$(ilRet) & " - " & Error(ilRet), vbExclamation, "Open Error" 'msgbox
        End If
        edcFrom.SetFocus
        mGetStnInfo = False
        Exit Function
    End If
    mGetStnInfo = True
    
    gAutomationAlertAndLogHandler "** Export Station **"
    
    ilFileID = 0
    ReDim tmStnInfo(0 To 0) As STNINFO
    err.Clear
    Do
        ilRet = 0
        'On Error GoTo mGetStnInfoErr:
        If EOF(hmStationFile) Then
            Exit Do
        End If
        Line Input #hmStationFile, slLine
        On Error GoTo 0
        ilRet = err.Number
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(Trim$(slLine)) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                gParseCDFields slLine, False, smFieldValues()
                For ilLoop = UBound(smFieldValues) - 1 To LBound(smFieldValues) Step 1
                    smFieldValues(ilLoop + 1) = smFieldValues(ilLoop)
                Next ilLoop
                ilUpper = UBound(tmStnInfo)
                ''If StrComp(Trim$(smFieldValues(3)), "Dummy Region", 1) <> 0 Then
                'If ((rbcInterface(0).Value) And ((StrComp(Trim$(smFieldValues(3)), "Dummy Region", 1) <> 0) And (Not rbcGen(2).Value)) Or (rbcGen(2).Value)) Or (rbcInterface(1).Value) Then
                'Use General with KenCast just like StarGuide: 8/17/05
                If StrComp(Trim$(smFieldValues(3)), "Dummy Region", 1) <> 0 Then
                    'Get Vehicle
                    ilFound = False
                    For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        'If StrComp(Trim$(tgMVef(ilVeh).sName), Trim$(smFieldValues(1)), 1) = 0 Then
                        '    ilFound = True
                        '    tmStnInfo(ilUpper).sType = "S"
                        '    tmStnInfo(ilUpper).sCallFreq = Trim$(smFieldValues(4))
                        '    tmStnInfo(ilUpper).iAirVeh = tgMVef(ilVeh).iCode
                        '    tmStnInfo(ilUpper).lRegionCode = Val(smFieldValues(3))
                        '    tmStnInfo(ilUpper).sSiteID = Trim$(smFieldValues(6))
                        '    tmStnInfo(ilUpper).sEDAS = Trim$(smFieldValues(8))
                        '    tmStnInfo(ilUpper).sTransportal = Trim$(smFieldValues(9))
                        '    tmStnInfo(ilUpper).sFileName = ""
                        '    tmStnInfo(ilUpper).sFdZone = Trim$(smFieldValues(10))
                        '    tmStnInfo(ilUpper).iLkStnInfo = -1
                        '    ReDim Preserve tmStnInfo(0 To UBound(tmStnInfo) + 1) As STNINFO
                        '    Exit For
                        'End If
                        If StrComp(Trim$(tgMVef(ilVeh).sName), Trim$(smFieldValues(2)), 1) = 0 Then
                            ilFound = True
                            ilMatch = False
                            For ilLoop = LBound(tmStnInfo) To UBound(tmStnInfo) - 1 Step 1
                                If StrComp(Trim$(tmStnInfo(ilLoop).sCallLetter), Trim$(smFieldValues(5)), 1) = 0 Then
                                    If StrComp(Trim$(tmStnInfo(ilLoop).sBand), Trim$(smFieldValues(6)), 1) = 0 Then
                                        If tmStnInfo(ilLoop).iAirVeh = tgMVef(ilVeh).iCode Then
                                            If tmStnInfo(ilLoop).lRegionCode = Val(smFieldValues(4)) Then
                                                If StrComp(Trim$(tmStnInfo(ilLoop).sSiteID), Trim$(smFieldValues(7)), 1) = 0 Then
                                                    If StrComp(Trim$(tmStnInfo(ilLoop).sFdZone), Trim$(smFieldValues(13)), 1) = 0 Then
                                                        ilMatch = True
                                                        If Trim$(smFieldValues(10)) <> "" Then
                                                            For ilEDAS = LBound(tmStnInfo(ilLoop).sEDAS) To UBound(tmStnInfo(ilLoop).sEDAS) Step 1
                                                                If Trim$(tmStnInfo(ilLoop).sEDAS(ilEDAS)) = "" Then
                                                                    tmStnInfo(ilLoop).sEDAS(ilEDAS) = Trim$(smFieldValues(10))
                                                                    Exit For
                                                                End If
                                                            Next ilEDAS
                                                        End If
                                                        If Trim$(smFieldValues(11)) <> "" Then
                                                            For ilTran = LBound(tmStnInfo(ilLoop).sTransportal) To UBound(tmStnInfo(ilLoop).sTransportal) Step 1
                                                                If Trim$(tmStnInfo(ilLoop).sTransportal(ilTran)) = "" Then
                                                                    tmStnInfo(ilLoop).sTransportal(ilTran) = Trim$(smFieldValues(11))
                                                                    Exit For
                                                                End If
                                                            Next ilTran
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next ilLoop
                            If Not ilMatch Then
                                For ilLoop = LBound(tmStnInfo(ilUpper).sTransportal) To UBound(tmStnInfo(ilUpper).sTransportal) Step 1
                                    tmStnInfo(ilUpper).sTransportal(ilLoop) = ""
                                Next ilLoop
                                For ilLoop = LBound(tmStnInfo(ilUpper).sEDAS) To UBound(tmStnInfo(ilUpper).sEDAS) Step 1
                                    tmStnInfo(ilUpper).sEDAS(ilLoop) = ""
                                Next ilLoop
                                tmStnInfo(ilUpper).sType = "S"
                                tmStnInfo(ilUpper).sCallLetter = Trim$(smFieldValues(5))
                                tmStnInfo(ilUpper).sBand = Trim$(smFieldValues(6))
                                tmStnInfo(ilUpper).iAirVeh = tgMVef(ilVeh).iCode
                                tmStnInfo(ilUpper).lRegionCode = Val(smFieldValues(4))
                                tmStnInfo(ilUpper).sSiteID = Trim$(smFieldValues(7))
                                tmStnInfo(ilUpper).sEDAS(0) = Trim$(smFieldValues(10))
                                tmStnInfo(ilUpper).sTransportal(0) = Trim$(smFieldValues(11))
                                tmStnInfo(ilUpper).sKCNo = Trim$(smFieldValues(12))
                                tmStnInfo(ilUpper).sFileName = ""
                                tmStnInfo(ilUpper).sStnFdCode = tgVpf(gVpfFind(ExpStnFd, tgMVef(ilVeh).iCode)).sStnFdCode
                                tmStnInfo(ilUpper).sFdZone = Trim$(smFieldValues(13))
                                'ABC request 1/17/06:  Show number of air plays in Commercial Log
                                tmStnInfo(ilUpper).iAirPlays = Val(smFieldValues(16))
                                tmStnInfo(ilUpper).sCmmlLogReq = Trim$(smFieldValues(17))
                                ilIndex = 18
                                For ilPledge = LBound(tmStnInfo(ilUpper).sCmmlLogPledge) To UBound(tmStnInfo(ilUpper).sCmmlLogPledge) Step 1
                                    tmStnInfo(ilUpper).sCmmlLogPledge(ilPledge) = Trim$(smFieldValues(ilIndex))
                                    ilIndex = ilIndex + 1
                                Next ilPledge
                                tmStnInfo(ilUpper).sKCEnvCopy = Trim$(smFieldValues(28))
                                tmStnInfo(ilUpper).sCmmlLogDPType = Trim$(smFieldValues(29))
                                tmStnInfo(ilUpper).sCmmlLogCart = Trim$(smFieldValues(30))
                                tmStnInfo(ilUpper).iLkStnInfo = -1
                                tmStnInfo(ilUpper).iLkCartInfo1 = -1
                                tmStnInfo(ilUpper).iLkCartInfo2 = -1
                                tmStnInfo(ilUpper).lRafCode = 0
                                ReDim Preserve tmStnInfo(0 To UBound(tmStnInfo) + 1) As STNINFO
                            End If
                            Exit For
                        End If
                    Next ilVeh
                    If Not ilMatch Then
                        If Not ilFound Then
                            If ilTest Then
                                'Print #hmMsg, "Vehicle: " & smFieldValues(2) & " not defined"
                                If (rbcInterface(0).Value) And (rbcGen(2).Value) Then
                                    For ilVeh = 0 To lbcRegVeh.ListCount - 1 Step 1
                                        If lbcRegVeh.Selected(ilVeh) Then
                                            If StrComp(Trim$(lbcRegVeh.List(ilVeh)), Trim$(smFieldValues(2)), 1) = 0 Then
                                                'Print #hmMsg, "Vehicle: " & smFieldValues(2) & " not defined"
                                                gAutomationAlertAndLogHandler "Vehicle: " & smFieldValues(2) & " not defined"
                                                mGetStnInfo = False
                                                Exit For
                                            End If
                                        End If
                                    Next ilVeh
                                Else
                                    'Print #hmMsg, "Vehicle: " & smFieldValues(2) & " not defined"
                                    gAutomationAlertAndLogHandler "Vehicle: " & smFieldValues(2) & " not defined"
                                    mGetStnInfo = False
                                End If
                            Else
                                mGetStnInfo = False
                            End If
                            'mGetStnInfo = False
                        Else
                            If ((rbcInterface(0).Value) And (rbcGen(0).Value Or rbcGen(1).Value)) Or (rbcInterface(1).Value) Then
                                If Trim$(smFieldValues(4)) <> "" Then
                                    tmRafSrchKey2.lCode = Val(smFieldValues(4))
                                    ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                    If ilRet = BTRV_ERR_NONE Then
                                        tmStnInfo(ilUpper).lRafCode = tmRaf.lCode
                                        ilFound = False
                                        For ilStn = 0 To ilUpper - 1 Step 1
                                            If StrComp(tmStnInfo(ilStn).sSiteID, tmStnInfo(ilUpper).sSiteID, 1) = 0 Then

                                                'Is vehicle within same group
                                                ilSameGroup = False
                                                If tmStnInfo(ilStn).iAirVeh = tmStnInfo(ilUpper).iAirVeh Then
                                                    ilSameGroup = True
                                                Else
                                                    For ilVeh = 0 To UBound(tmVef) - 1 Step 1
                                                        If (tmVef(ilVeh).sType = "A") Or (tmVef(ilVeh).sType = "C") Then
                                                            If tmVef(ilVeh).iCode = tmStnInfo(ilStn).iAirVeh Then
                                                                ilVpfIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                                                                If ilVpfIndex >= 0 Then
                                                                    ilNextLk = tmVpfInfo(ilVpfIndex).iFirstLkVehInfo
                                                                    Do While ilNextLk >= 0
                                                                        If tmLkVehInfo(ilNextLk).iVefCode = tmStnInfo(ilUpper).iAirVeh Then
                                                                            ilSameGroup = True
                                                                            Exit Do
                                                                        End If
                                                                        ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                                                                    Loop
                                                                End If
                                                            Else
                                                                ilVpfIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                                                                If ilVpfIndex >= 0 Then
                                                                    ilNextLk = tmVpfInfo(ilVpfIndex).iFirstLkVehInfo
                                                                    Do While ilNextLk >= 0
                                                                        If tmLkVehInfo(ilNextLk).iVefCode = tmStnInfo(ilStn).iAirVeh Then
                                                                            If tmVef(ilVeh).iCode = tmStnInfo(ilUpper).iAirVeh Then
                                                                                ilSameGroup = True
                                                                            Else
                                                                                ilVpfIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                                                                                If ilVpfIndex >= 0 Then
                                                                                    ilNextLk = tmVpfInfo(ilVpfIndex).iFirstLkVehInfo
                                                                                    Do While ilNextLk >= 0
                                                                                        If tmLkVehInfo(ilNextLk).iVefCode = tmStnInfo(ilUpper).iAirVeh Then
                                                                                            ilSameGroup = True
                                                                                            Exit Do
                                                                                        End If
                                                                                        ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                                                                                    Loop
                                                                                End If
                                                                            End If

                                                                            Exit Do
                                                                        End If
                                                                        ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                                                                    Loop
                                                                End If
                                                            End If
                                                        End If
                                                    Next ilVeh
                                                End If


                                                'DL 4/28/05: If not in same group, then don't link
                                                'If tmStnInfo(ilStn).iLkStnInfo = -1 Then
                                                '    ilFound = True
                                                '    tmStnInfo(ilStn).iLkStnInfo = ilUpper
                                                '    tmStnInfo(ilUpper).sType = "L"
                                                '    Exit For
                                                'ElseIf ilSameGroup Then
                                                If ilSameGroup Then
                                                    ilFound = True
                                                    If tmStnInfo(ilStn).iLkStnInfo = -1 Then
                                                        tmStnInfo(ilStn).iLkStnInfo = ilUpper
                                                        tmStnInfo(ilUpper).sType = "L"
                                                    Else
                                                        ilNextLk = tmStnInfo(ilStn).iLkStnInfo
                                                        Do While ilNextLk >= 0
                                                            If tmStnInfo(ilNextLk).iLkStnInfo = -1 Then
                                                                tmStnInfo(ilNextLk).iLkStnInfo = ilUpper
                                                                tmStnInfo(ilUpper).sType = "L"
                                                                Exit Do
                                                            End If
                                                            ilNextLk = tmStnInfo(ilNextLk).iLkStnInfo
                                                        Loop
                                                    End If
                                                    Exit For
                                                End If
                                            End If
                                        Next ilStn
                                        If Not ilFound Then
                                            ''tmStnInfo(ilStn).sFileName = Trim$(tmStnInfo(ilStn).sSiteID) & smFeedNo
                                            'ilFileID = ilFileID + 1
                                            'slStr = Trim$(Str$(ilFileID))
                                            'Do While Len(slStr) < 4
                                            '    slStr = "0" & slStr
                                            'Loop
                                            'tmStnInfo(ilStn).sFileName = slStr & Right$(smFeedNo, 4)
                                            'tmStnInfo(ilStn).sFileName = Trim$(tmStnInfo(ilStn).sCallLetter) & Trim$(Left$(tmStnInfo(ilStn).sBand, 1)) & smWeekNo & smRunLetter
                                            If rbcInterface(0).Value Then
                                                If (rbcGen(0).Value) Then
                                                    tmStnInfo(ilStn).sFileName = Trim$(tmStnInfo(ilStn).sCallLetter) & Trim$(Left$(tmStnInfo(ilStn).sBand, 1)) & smWeekNo & smRunLetter
                                                Else
                                                    tmStnInfo(ilStn).sFileName = Trim$(tmStnInfo(ilStn).sCallLetter) & Trim$(Left$(tmStnInfo(ilStn).sBand, 1)) & smFeedNo & smRunLetter
                                                End If
                                            Else
                                                tmStnInfo(ilStn).sFileName = Trim$(tmStnInfo(ilStn).sCallLetter) & "-" & Trim$(tmStnInfo(ilStn).sBand)
                                            End If
                                        End If
                                    Else
                                        If ilTest Then
                                            'Print #hmMsg, "Region: " & smFieldValues(3) & " " & smFieldValues(4) & " not defined"
                                            gAutomationAlertAndLogHandler "Region: " & smFieldValues(3) & " " & smFieldValues(4) & " not defined"
                                        End If
                                        mGetStnInfo = False
                                    End If
                                End If
                            ElseIf (rbcInterface(0).Value) And (rbcGen(2).Value) Then
                                ilFound = False
                                For ilStn = 0 To ilUpper - 1 Step 1
                                    If StrComp(tmStnInfo(ilStn).sSiteID, tmStnInfo(ilUpper).sSiteID, 1) = 0 Then
                                        If tmStnInfo(ilStn).iLkStnInfo = -1 Then
                                            ilFound = True
                                            tmStnInfo(ilStn).iLkStnInfo = ilUpper
                                            tmStnInfo(ilUpper).sType = "L"
                                            Exit For
                                        End If
                                    End If
                                Next ilStn
                                If Not ilFound Then
                                    ilStn = ilUpper
                                    tmStnInfo(ilStn).sFileName = Trim$(tmStnInfo(ilStn).sCallLetter) & Trim$(Left$(tmStnInfo(ilStn).sBand, 1)) & smFeedNo & smRunLetter
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    Close hmStationFile
    'Use generic file names for KenCast like StarGuide to reduce generation time: 8/17/05
    'If rbcInterface(1).Value Then
    '    'Force all stations to have a file name instead of generic
    '    For ilStn = 0 To UBound(tmStnInfo) - 1 Step 1
    '        If Trim$(tmStnInfo(ilStn).sFileName) = "" Then
    '            tmStnInfo(ilStn).sFileName = Trim$(tmStnInfo(ilStn).sCallLetter) & "-" & Trim$(tmStnInfo(ilStn).sBand)
    '        End If
    '    Next ilStn
    'End If
    If lbcRegVeh.ListCount > 0 Then
        Exit Function
    End If
    ReDim imRegVefCode(0 To 0) As Integer
    For ilLoop = 0 To UBound(tmStnInfo) - 1 Step 1
        'If (tmStnInfo(ilLoop).sType = "S") And (tmStnInfo(ilLoop).lRafCode > 0) Then 'Ignore linked records- handled within
        If ((rbcInterface(0).Value) And ((tmStnInfo(ilLoop).sType = "S") And (tmStnInfo(ilLoop).lRafCode > 0) And (Not rbcGen(2).Value)) Or ((tmStnInfo(ilLoop).sType = "S") And (rbcGen(2).Value))) Or (rbcInterface(1).Value) Then 'Ignore linked records- handled within
            ilFound = False
            For ilVeh = 0 To UBound(imRegVefCode) - 1 Step 1
                If tmStnInfo(ilLoop).iAirVeh = imRegVefCode(ilVeh) Then
                    ilFound = True
                    Exit For
                End If
            Next ilVeh
            If Not ilFound Then
                imRegVefCode(UBound(imRegVefCode)) = tmStnInfo(ilLoop).iAirVeh
                ReDim Preserve imRegVefCode(0 To UBound(imRegVefCode) + 1) As Integer
            End If
            ilVIndex = tmStnInfo(ilLoop).iLkStnInfo
            Do While ilVIndex >= 0
                ilFound = False
                For ilVeh = 0 To UBound(imRegVefCode) - 1 Step 1
                    If tmStnInfo(ilVIndex).iAirVeh = imRegVefCode(ilVeh) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilVeh
                If Not ilFound Then
                    imRegVefCode(UBound(imRegVefCode)) = tmStnInfo(ilVIndex).iAirVeh
                    ReDim Preserve imRegVefCode(0 To UBound(imRegVefCode) + 1) As Integer
                End If
                ilVIndex = tmStnInfo(ilVIndex).iLkStnInfo
            Loop
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(imRegVefCode) - 1 Step 1
        tmVefSrchKey.iCode = imRegVefCode(ilLoop)
        ilRet = btrGetEqual(hmVef, tmEVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) Then
            lbcRegVeh.AddItem Trim$(tmEVef.sName)
        End If
    Next ilLoop
    Exit Function
'mGetStnInfoErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDateTime As String
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    igJobShowing(STATIONFEEDJOB) = True
    imIgnoreCkcAll = False
    imLastIndex = -1
    imShiftKey = 0
    imIgnoreVbcChg = False
    imTypeIndex = 0
    imTerminate = False
    ReDim tgSortCrf(0 To 0) As SORTCRF
    'mParseCmmdLine
    ExpStnFd.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone ExpStnFd
    'ExpStnFd.Show
    DoEvents
    'If (Trim$(tgUrf(0).sPDFDrvChar) <> "") And (tgUrf(0).iPDFDnArrowCnt >= 0) And (Trim$(tgUrf(0).sPrtDrvChar) <> "") And (tgUrf(0).iPrtDnArrowCnt >= 0) Then
        rbcFormat(0).Enabled = True
        rbcFormat(0).Value = True
        rbcFormat(2).Enabled = True
        rbcFormat(2).Value = True
    'Else
    '    rbcFormat(0).Enabled = False
    '    rbcFormat(1).Value = True
    'End If
    imWaitCount = csiHandleValue(0, 12)
    imTimeDelay = csiHandleValue(0, 13)
    imLockValue = csiHandleValue(0, 14)
    imTranLog = csiHandleValue(0, 15)
    csiSetValue 90, 2, imLockValue, imTranLog
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    imDateBox = -1
    imExptPrevWeek = True
    imIgnoreRightMove = False
    imButton = 0
    imListFieldRot(1) = 15
    imListFieldRot(2) = 15 * igAlignCharWidth
    imListFieldRot(3) = 23 * igAlignCharWidth
    imListFieldRot(4) = 33 * igAlignCharWidth
    imListFieldRot(5) = 40 * igAlignCharWidth
    imListFieldRot(6) = 52 * igAlignCharWidth
    imListFieldRot(7) = 67 * igAlignCharWidth
    imListFieldRot(8) = 78 * igAlignCharWidth
    imListFieldRot(9) = 82 * igAlignCharWidth
    imListFieldRot(10) = 92 * igAlignCharWidth
    imListFieldRot(11) = 97 * igAlignCharWidth
    imListFieldRot(12) = 100 * igAlignCharWidth
    imListFieldVeh(1) = 15
    imListFieldVeh(2) = 75 * igAlignCharWidth
    imListFieldVeh(3) = 90 * igAlignCharWidth

    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Chf.btr)", ExpStnFd
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)     'Get and save CHF record length
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Clf.btr)", ExpStnFd
    On Error GoTo 0
    imClfRecLen = Len(tmClf)     'Get and save CHF record length
    hmCrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Crf.Btr)", ExpStnFd
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)     'Get and save CRF record length
    hmSif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Sif.Btr)", ExpStnFd
    On Error GoTo 0
    imSifRecLen = Len(tmSif)
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Vsf.Btr)", ExpStnFd
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmCnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Cnf.Btr)", ExpStnFd
    On Error GoTo 0
    imCnfRecLen = Len(tmCnf)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Mcf.Btr)", ExpStnFd
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Cif.Btr)", ExpStnFd
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmCyf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmCyf, "", sgDBPath & "Cyf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Cyf.Btr)", ExpStnFd
    On Error GoTo 0
    imCyfRecLen = Len(tmCyf)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Cpf.Btr)", ExpStnFd
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmBof = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmBof, "", sgDBPath & "Bof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Bof.Btr)", ExpStnFd
    On Error GoTo 0
    imBofRecLen = Len(tmBof)
    hmCpr = CBtrvTable(TEMPHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Cpr.Btr)", ExpStnFd
    On Error GoTo 0
    imCprRecLen = Len(tmCpr)
    hmCsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCsf, "", sgDBPath & "Csf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Csf.Btr)", ExpStnFd
    On Error GoTo 0
    imCsfRecLen = Len(tmCsf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Adf.Btr)", ExpStnFd
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Anf.Btr)", ExpStnFd
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Vef.Btr)", ExpStnFd
    On Error GoTo 0
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Vlf.Btr)", ExpStnFd
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)
    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Raf.Btr)", ExpStnFd
    imRafRecLen = Len(tmRaf)
    On Error GoTo 0
    hmRsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Rsf.Btr)", ExpStnFd
    imRsfRecLen = Len(tmRsf)
    On Error GoTo 0
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Sdf.Btr)", ExpStnFd
    imSdfRecLen = Len(tmSdf)
    On Error GoTo 0
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Ssf.Btr)", ExpStnFd
    imSsfRecLen = Len(tmCTSsf)
    On Error GoTo 0
    hmCTSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCTSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Ssf.Btr)", ExpStnFd
    imSsfRecLen = Len(tmCTSsf)
    On Error GoTo 0
    hmTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Tzf.Btr)", ExpStnFd
    imTzfRecLen = Len(tmTzf)
    On Error GoTo 0
    hmDlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmDlf, "", sgDBPath & "Dlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Dlf.Btr)", ExpStnFd
    imDlfRecLen = Len(tmDlf)
    On Error GoTo 0
    'hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    'ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'On Error GoTo mInitErr
    'gBtrvErrorMsg ilRet, "mInit (btrOpen-Vlf.Btr)", ExpStnFd
    'imVlfRecLen = Len(tmVlf)
    'On Error GoTo 0
    ReDim imRegVefCode(0 To 0) As Integer
    ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
    mVehPop True
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    tmAdf.iCode = 0
    tmAnf.iCode = 0
    ilRet = gObtainVef()
    lbcVeh.Clear
    lbcRegVeh.Clear
    'ilRet = gPopUserVehicleBox(ExpStnFd, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH, lbcVeh, lbcVehCode)
    ilRet = gPopUserVehicleBox(ExpStnFd, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH, lbcVeh, tmVehCode(), smVehCodeTag)
    'mRotPop
    smTodaysDate = Format$(gNow(), "mm/dd/yy")
    edcStartDate.Text = gObtainNextMonday(smTodaysDate)
    slStr = edcStartDate.Text
    edcEndDate.Text = gObtainNextSunday(slStr)
    tmcRot.Enabled = False
    edcTranDate.Text = smTodaysDate
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    lacDate.Visible = False
    cmcExport.Enabled = False
    'cmcView.Enabled = False
    imNoRotations = 0
    lacProcessing.Caption = ""
    'tmcDDE.Enabled = True
    plcRotInfo.Move 1410, 4455
    'gCenterStdAlone ExpStnFd
    For ilLoop = LBound(imEvtType) To UBound(imEvtType) Step 1
        imEvtType(ilLoop) = False
    Next ilLoop
    'Get avails only
    For ilLoop = 2 To 9 Step 1
        imEvtType(ilLoop) = True
    Next ilLoop
    imEvtType(0) = False 'Don't include library names
    Screen.MousePointer = vbDefault
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    Exit Sub
mInitErr:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move plcDates.Left + edcTranDate.Left, plcDates.Top + edcTranDate.Top + edcTranDate.Height
    pbcLbcRot.Move lbcRot.Left + 15, lbcRot.Top + 15, pbcLbcRot.Width, lbcRot.Height - 30
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeCartStn                    *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Make cross ref table of carts  *
'*                      and stations                   *
'*                                                     *
'*******************************************************
Private Sub mMakeCartStn(tlStnInfo As STNINFO)
    Dim ilLoop As Integer
    Dim slKey As String
    Dim ilRet As Integer
    Dim slProduct As String
    Dim slCart As String
    Dim slISCI As String
    Dim ilUpper1 As Integer
    Dim ilUpper2 As Integer
    Dim ilLkIndex1 As Integer
    Dim ilLkIndex2 As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer

    lacProcessing.Caption = "Merge Copy/Station Cross Reference"
    For ilLoop = LBound(tmAddCyf) To UBound(tmAddCyf) - 1 Step 1
        slKey = Trim$(tmAddCyf(ilLoop).sXFKey)
        ilRet = gParseItem(slKey, 1, "|", slProduct)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 2, "|", slCart)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 3, "|", slISCI)  'Obtain Index and code number
        'If tlStnInfo.iLkCartInfo = -1 Then
        '    ilUpper = UBound(tgCartStnXRef)
        '    tlStnInfo.iLkCartInfo = ilUpper
        '    tgCartStnXRef(ilUpper).lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode
        '    tgCartStnXRef(ilUpper).sShortTitle = gFileNameFilter(slProduct)
        '    tgCartStnXRef(ilUpper).sISCI = gFileNameFilter(slISCI)
        '    tgCartStnXRef(ilUpper).iLkCartInfo = -1
        '    tgCartStnXRef(ilUpper).iAdfCode = tmAddCyf(ilLoop).iAdfCode
        '    tgCartStnXRef(ilUpper).iLen = tmAddCyf(ilLoop).iLen
        '    ReDim Preserve tgCartStnXRef(0 To ilUpper + 1) As CARTSTNXREF
        If tlStnInfo.iLkCartInfo1 = -1 Then
            ilUpper1 = imCartStnXRef1
            ilUpper2 = UBound(tgCartStnXRef, 2)
            tlStnInfo.iLkCartInfo1 = ilUpper1
            tlStnInfo.iLkCartInfo2 = ilUpper2
            tgCartStnXRef(ilUpper1, ilUpper2).lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode
            tgCartStnXRef(ilUpper1, ilUpper2).sShortTitle = gFileNameFilter(slProduct)
            tgCartStnXRef(ilUpper1, ilUpper2).sISCI = gFileNameFilter(slISCI)
            tgCartStnXRef(ilUpper1, ilUpper2).iLkCartInfo1 = -1
            tgCartStnXRef(ilUpper1, ilUpper2).iLkCartInfo2 = -1
            tgCartStnXRef(ilUpper1, ilUpper2).iAdfCode = tmAddCyf(ilLoop).iAdfCode
            tgCartStnXRef(ilUpper1, ilUpper2).iLen = tmAddCyf(ilLoop).iLen
            tgCartStnXRef(ilUpper1, ilUpper2).iFdDateNew = tmAddCyf(ilLoop).iFdDateNew
            imCartStnXRef1 = imCartStnXRef1 + 1
            If imCartStnXRef1 > 32000 Then
                imCartStnXRef1 = 0
                ReDim Preserve tgCartStnXRef(0 To 32000, 0 To ilUpper2 + 1) As CARTSTNXREF
            End If
        Else
            'ilFound = False
            'ilLkIndex = tlStnInfo.iLkCartInfo
            'Do While ilLkIndex <> -1
            '    If tgCartStnXRef(ilLkIndex).lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode Then
            '        ilFound = True
            '        Exit Do
            '    End If
            '    If tgCartStnXRef(ilLkIndex).iLkCartInfo = -1 Then
            '        Exit Do
            '    End If
            '    ilLkIndex = tgCartStnXRef(ilLkIndex).iLkCartInfo
            'Loop
            'If Not ilFound Then
            '    ilUpper = UBound(tgCartStnXRef)
            '    tgCartStnXRef(ilLkIndex).iLkCartInfo = ilUpper
            '    tgCartStnXRef(ilUpper).lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode
            '    tgCartStnXRef(ilUpper).sShortTitle = gFileNameFilter(slProduct)
            '    tgCartStnXRef(ilUpper).sISCI = gFileNameFilter(slISCI)
            '    tgCartStnXRef(ilUpper).iLkCartInfo = -1
            '    tgCartStnXRef(ilUpper).iAdfCode = tmAddCyf(ilLoop).iAdfCode
            '    tgCartStnXRef(ilUpper).iLen = tmAddCyf(ilLoop).iLen
            '    ReDim Preserve tgCartStnXRef(0 To ilUpper + 1) As CARTSTNXREF
            'End If
            ilFound = False
            ilLkIndex1 = tlStnInfo.iLkCartInfo1
            ilLkIndex2 = tlStnInfo.iLkCartInfo2
            Do While ilLkIndex1 <> -1
                If tgCartStnXRef(ilLkIndex1, ilLkIndex2).lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode Then
                    ilFound = True
                    Exit Do
                End If
                If tgCartStnXRef(ilLkIndex1, ilLkIndex2).iLkCartInfo1 = -1 Then
                    Exit Do
                End If
                ilIndex = ilLkIndex1
                ilLkIndex1 = tgCartStnXRef(ilIndex, ilLkIndex2).iLkCartInfo1
                ilLkIndex2 = tgCartStnXRef(ilIndex, ilLkIndex2).iLkCartInfo2
            Loop
            If Not ilFound Then
                ilUpper1 = imCartStnXRef1
                ilUpper2 = UBound(tgCartStnXRef, 2)
                tgCartStnXRef(ilLkIndex1, ilLkIndex2).iLkCartInfo1 = ilUpper1
                tgCartStnXRef(ilLkIndex1, ilLkIndex2).iLkCartInfo2 = ilUpper2
                tgCartStnXRef(ilUpper1, ilUpper2).lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode
                tgCartStnXRef(ilUpper1, ilUpper2).sShortTitle = gFileNameFilter(slProduct)
                tgCartStnXRef(ilUpper1, ilUpper2).sISCI = gFileNameFilter(slISCI)
                tgCartStnXRef(ilUpper1, ilUpper2).iLkCartInfo1 = -1
                tgCartStnXRef(ilUpper1, ilUpper2).iLkCartInfo2 = -1
                tgCartStnXRef(ilUpper1, ilUpper2).iAdfCode = tmAddCyf(ilLoop).iAdfCode
                tgCartStnXRef(ilUpper1, ilUpper2).iLen = tmAddCyf(ilLoop).iLen
                tgCartStnXRef(ilUpper1, ilUpper2).iFdDateNew = tmAddCyf(ilLoop).iFdDateNew
                'ReDim Preserve tgCartStnXRef(0 To ilUpper + 1) As CARTSTNXREF
                imCartStnXRef1 = imCartStnXRef1 + 1
                If imCartStnXRef1 > 32000 Then
                    imCartStnXRef1 = 0
                    ReDim Preserve tgCartStnXRef(0 To 32000, 0 To ilUpper2 + 1) As CARTSTNXREF
                End If
            End If
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMergeAddCyf                    *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Merge tmAddCyf into tgAddCyf   *
'*                                                     *
'*******************************************************
Private Sub mMergeAddCyf()
    Dim ilFound As Integer
    Dim ilCyf As Integer
    Dim ilLoop As Integer
    lacProcessing.Caption = "Merge Copy Inventory"
    'Test for duplicates
    For ilLoop = LBound(tmAddCyf) To UBound(tmAddCyf) - 1 Step 1
        ilFound = False
        For ilCyf = LBound(tgAddCyf) To UBound(tgAddCyf) - 1 Step 1
            If (tmAddCyf(ilLoop).tCyf.lCifCode = tgAddCyf(ilCyf).tCyf.lCifCode) And (tmAddCyf(ilLoop).tCyf.iVefCode = tgAddCyf(ilCyf).tCyf.iVefCode) Then   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) Then
                If (tmAddCyf(ilLoop).tCyf.sTimeZone = tgAddCyf(ilCyf).tCyf.sTimeZone) And (tmAddCyf(ilLoop).tCyf.lRafCode = tgAddCyf(ilCyf).tCyf.lRafCode) Then
                    ilFound = True
                    If tmAddCyf(ilLoop).lRotStartDate < tgAddCyf(ilCyf).lRotStartDate Then
                        tgAddCyf(ilCyf).lRotStartDate = tmAddCyf(ilLoop).lRotStartDate
                    End If
                    If tmAddCyf(ilLoop).lRotEndDate > tgAddCyf(ilCyf).lRotEndDate Then
                        tgAddCyf(ilCyf).lRotEndDate = tmAddCyf(ilLoop).lRotEndDate
                    End If
                    Exit For
                End If
            End If
        Next ilCyf
        If Not ilFound Then
            tgAddCyf(UBound(tgAddCyf)) = tmAddCyf(ilLoop)
            ReDim Preserve tgAddCyf(0 To UBound(tgAddCyf) + 1) As SENDCOPYINFO
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMergeXRefCyf                   *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Merge tmAddCyf into tmXRefCyf  *
'*                                                     *
'*******************************************************
Private Sub mMergeXRefCyf()
    Dim ilFound As Integer
    Dim ilCyf As Integer
    Dim ilLoop As Integer
    lacProcessing.Caption = "Merge Cross Reference"
    'Test for duplicates
    For ilLoop = LBound(tmAddCyf) To UBound(tmAddCyf) - 1 Step 1
        ilFound = False
        For ilCyf = LBound(tmXRefCyf) To UBound(tmXRefCyf) - 1 Step 1
            If (tmAddCyf(ilLoop).tCyf.lCifCode = tmXRefCyf(ilCyf).tCyf.lCifCode) And (tmAddCyf(ilLoop).tCyf.iVefCode = tmXRefCyf(ilCyf).tCyf.iVefCode) Then   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) Then
                If (tmAddCyf(ilLoop).tCyf.sTimeZone = tmXRefCyf(ilCyf).tCyf.sTimeZone) And (tmAddCyf(ilLoop).tCyf.lRafCode = tmXRefCyf(ilCyf).tCyf.lRafCode) Then
                    ilFound = True
                    If tmAddCyf(ilLoop).lRotStartDate < tmXRefCyf(ilCyf).lRotStartDate Then
                        tmXRefCyf(ilCyf).lRotStartDate = tmAddCyf(ilLoop).lRotStartDate
                    End If
                    If tmAddCyf(ilLoop).lRotEndDate > tmXRefCyf(ilCyf).lRotEndDate Then
                        tmXRefCyf(ilCyf).lRotEndDate = tmAddCyf(ilLoop).lRotEndDate
                    End If
                    Exit For
                End If
            End If
        Next ilCyf
        If Not ilFound Then
            tmXRefCyf(UBound(tmXRefCyf)) = tmAddCyf(ilLoop)
            ReDim Preserve tmXRefCyf(0 To UBound(tmXRefCyf) + 1) As SENDCOPYINFO
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mNonRotFileNames                *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add selected vehicles to       *
'*                      station information            *
'*                                                     *
'*******************************************************
Private Function mNonRotFileNames() As Integer
    Dim ilVeh As Integer
    Dim slStnCode As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilVehIndex As Integer
    Dim ilVefSelected As Integer
    Dim slNameTime As String
    Dim slVehName As String
    Dim ilPos As Integer
    'Dim ilLen As Integer
    Dim ilSAGroupNo As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVIndex As Integer
    Dim slExportFile As String
    Dim slTimeStamp As String
    Dim ilUpper As Integer
    Dim ilPledge As Integer

    'Determine if file exist- if so don't allow export
    For ilVeh = 0 To lbcVehicle.ListCount - 1 Step 1
        slNameTime = lbcVehicle.List(ilVeh) 'Airing and conventional vehicles (with and without bulk groups)
        ilPos = InStr(slNameTime, "|")
        slVehName = Left$(slNameTime, ilPos - 1)
        For ilLoop = 0 To UBound(tmVef) - 1 Step 1
            'ilLen = Len(Trim$(tmVef(ilLoop).sName))
            'If (Trim$(tmVef(ilLoop).sName) = Left$(slVehName, ilLen)) Then
            If tmVef(ilLoop).iCode = lbcVehicle.ItemData(ilVeh) Then
                ilVehIndex = ilLoop
                ilSAGroupNo = tgVpf(gVpfFind(ExpStnFd, tmVef(ilLoop).iCode)).iSAGroupNo
                Exit For
            End If
        Next ilLoop
        'Test if vehicle has a rotation to be transmitted
        ilVefSelected = False
        For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1   'lbcVeh: Selling and Conventional
            If lbcVeh.Selected(ilLoop) Then
                slNameCode = tmVehCode(ilLoop).sKey    'Selling and conventional vehicles 'lbcVehCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                If tmVef(ilVehIndex).sType = "C" Then
                    If tmVef(ilVehIndex).iCode = Val(slCode) Then
                        ilVefSelected = True
                        Exit For
                    End If
                    ilVIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                        Do While ilVIndex >= 0
                            If tmLkVehInfo(ilVIndex).iVefCode = tmVef(ilVehIndex).iCode Then
                                ilVefSelected = True
                                Exit Do
                            End If
                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                Else
                    ilVIndex = mFindVpfIndex(Val(slCode)) 'tmVef(ilVehIndex).iCode)
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                        Do While ilVIndex >= 0
                            If tmVef(ilVehIndex).iCode = tmSALink(ilVIndex).iVefCode Then
                                ilVefSelected = True
                                Exit Do
                            End If
                            ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                End If
            End If
        Next ilLoop
        If ilVefSelected Then
            'slStnCode = Trim$(tmVef(ilVehIndex).sCodeStn)
            slStnCode = Left$(tgVpf(gVpfFind(ExpStnFd, tmVef(ilVehIndex).iCode)).sStnFdCode, 2)
            If rbcInterface(0).Value Then
                If rbcGen(3).Value Then
                    smRunLetter = "A"
                    Do
                        If rbcFormat(0).Value Then
                            slExportFile = sgExportPath & slStnCode & smAllInstFileDate & smRunLetter & ".PDF"
                        Else
                            slExportFile = sgExportPath & slStnCode & smAllInstFileDate & smRunLetter & ".Txt"
                        End If
                        ilRet = 0
                        'On Error GoTo mNonRotFileNamesErr:
                        'slTimeStamp = FileDateTime(slExportFile)
                        ilRet = gFileExist(slExportFile)
                        If ilRet = 0 Then
                            smRunLetter = Chr(Asc(smRunLetter) + 1)
                        End If
                    Loop While ilRet = 0
                Else
                    'slExportFile = sgExportPath & slStnCode & smFeedNo & ".trf"
                    slExportFile = sgExportPath & "GN" & smWeekNo & smRunLetter & "." & slStnCode & "X"
                    ilRet = 0
                    'On Error GoTo mNonRotFileNamesErr:
                    'slTimeStamp = FileDateTime(slExportFile)
                    ilRet = gFileExist(slExportFile)
                    If ilRet = 0 Then
                        Screen.MousePointer = vbDefault
                        'MsgBox "Station Feed already generated for this date, Export terminated", vbOkOnly + vbCritical + vbApplicationModal, "Export"
                        'cmcCancel.SetFocus
                        mNonRotFileNames = False
                        Exit Function
                    End If
                End If
            Else
                'No test required
            End If
            ilUpper = UBound(tmStnInfo)
            'tmStnInfo(ilUpper).sType = "G"
            'tmStnInfo(ilUpper).sCallFreq = ""
            'tmStnInfo(ilUpper).iAirVeh = tmVef(ilVehIndex).iCode
            'tmStnInfo(ilUpper).lRegionCode = 0
            'tmStnInfo(ilUpper).sSiteID = slStnCode
            'tmStnInfo(ilUpper).sEDAS = ""
            'tmStnInfo(ilUpper).sTransportal = ""
            'tmStnInfo(ilUpper).sFileName = slStnCode & smFeedNo
            'tmStnInfo(ilUpper).iLkStnInfo = -1
            tmStnInfo(ilUpper).sType = "G"
            tmStnInfo(ilUpper).sCallLetter = ""
            tmStnInfo(ilUpper).sBand = ""
            tmStnInfo(ilUpper).iAirVeh = tmVef(ilVehIndex).iCode
            tmStnInfo(ilUpper).lRegionCode = 0
            tmStnInfo(ilUpper).sSiteID = slStnCode
            'tmStnInfo(ilUpper).sEDAS = ""
            'tmStnInfo(ilUpper).sTransportal = ""
            For ilLoop = LBound(tmStnInfo(ilUpper).sTransportal) To UBound(tmStnInfo(ilUpper).sTransportal) Step 1
                tmStnInfo(ilUpper).sTransportal(ilLoop) = ""
            Next ilLoop
            For ilLoop = LBound(tmStnInfo(ilUpper).sEDAS) To UBound(tmStnInfo(ilUpper).sEDAS) Step 1
                tmStnInfo(ilUpper).sEDAS(ilLoop) = ""
            Next ilLoop
            If rbcInterface(0).Value Then
                If rbcGen(3).Value Then
                    tmStnInfo(ilUpper).sFileName = slStnCode & smAllInstFileDate & smRunLetter
                Else
                    tmStnInfo(ilUpper).sFileName = "GN" & smWeekNo & smRunLetter
                End If
            Else
                tmStnInfo(ilUpper).sFileName = "General"
            End If
            tmStnInfo(ilUpper).sKCNo = ""
            tmStnInfo(ilUpper).sStnFdCode = slStnCode
            tmStnInfo(ilUpper).sFdZone = ""
            tmStnInfo(ilUpper).iAirPlays = 0
            tmStnInfo(ilUpper).sCmmlLogReq = ""
            'For ilPledge = 18 To 27 Step 1
            For ilPledge = LBound(tmStnInfo(ilUpper).sCmmlLogPledge) To UBound(tmStnInfo(ilUpper).sCmmlLogPledge) Step 1
                tmStnInfo(ilUpper).sCmmlLogPledge(ilPledge) = ""    ' - 17) = ""
            Next ilPledge
            tmStnInfo(ilUpper).sKCEnvCopy = "A"
            tmStnInfo(ilUpper).sCmmlLogDPType = "S"
            tmStnInfo(ilUpper).sCmmlLogCart = "C"
            tmStnInfo(ilUpper).iLkStnInfo = -1
            tmStnInfo(ilUpper).iLkCartInfo1 = -1
            tmStnInfo(ilUpper).iLkCartInfo2 = -1
            tmStnInfo(ilUpper).lRafCode = 0
            ReDim Preserve tmStnInfo(0 To UBound(tmStnInfo) + 1) As STNINFO
        End If
    Next ilVeh
    mNonRotFileNames = True
    Exit Function
'mNonRotFileNamesErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCopy                     *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Private Function mObtainCopy(slZone As String, slShortTitle As String) As Integer
'
'   mObtainCopy
'       Where:
'           tmSdf(I)- Spot record
'           tmCif(O)- Inventory
'           tmCpf(O)- Product/ISCI record
'
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilCifFound As Integer
    ilCifFound = False
    tmCpf.sISCI = ""
    tmCpf.sName = ""
    tmCpf.sCreative = ""
    tmMcf.sName = "C"
    tmMcf.sPrefix = "C"
    slShortTitle = ""
    If tmSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilCifFound = True
        End If
    ElseIf tmSdf.sPtType = "2" Then  '  Combo Copy
    ElseIf tmSdf.sPtType = "3" Then  '  Time Zone Copy
        ' Read TZF using lCopyCode from SDF
        tmTzfSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ' Look for the first positive lZone value
        For ilIndex = 1 To 6 Step 1
            If tmTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                If StrComp(tmTzf.sZone(ilIndex - 1), slZone, 1) = 0 Then
                    ' Read CIF using lCopyCode from SDF
                    tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        ilCifFound = True
                    End If
                    Exit For
                End If
            End If
        Next ilIndex
        If Not ilCifFound Then
            For ilIndex = 1 To 6 Step 1
                If tmTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                    If StrComp(tmTzf.sZone(ilIndex - 1), "Oth", 1) = 0 Then
                        ' Read CIF using lCopyCode from SDF
                        tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            ilCifFound = True
                        End If
                        Exit For
                    End If
                End If
            Next ilIndex
        End If
    End If
    If ilCifFound Then
        ' Read CPF using lCpfCode from CIF
        If tmCif.lcpfCode > 0 Then
            tmCpfSrchKey.lCode = tmCif.lcpfCode
            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmCpf.sISCI = ""
                tmCpf.sName = ""
                tmCpf.sCreative = ""
            Else
                If (tgSpf.sUseProdSptScr = "P") Then
                    slShortTitle = Trim$(tmCpf.sName)
                End If
            End If
        Else
            tmCpf.sISCI = ""
            tmCpf.sName = ""
            tmCpf.sCreative = ""
        End If
        If tmMcf.iCode <> tmCif.iMcfCode Then
            tmMcfSrchKey.iCode = tmCif.iMcfCode
            ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Else
            ilRet = BTRV_ERR_NONE
        End If
        If ilRet <> BTRV_ERR_NONE Then
            tmMcf.sName = "C"
            tmMcf.sPrefix = "C"
            mObtainCopy = False
            Exit Function
        End If
        mObtainCopy = True
        Exit Function
    End If
    mObtainCopy = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile() As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    'On Error GoTo mOpenMsgFileErr:
    'slToFile = sgExportPath & "ExpStnFd.Txt"
    If rbcInterface(0).Value Then
        If rbcGen(1).Value Then
            'slToFile = sgExportPath & "ExpRgSpt.Txt"
            slToFile = sgDBPath & "Messages\" & "ExpRgSpt.Txt"
        ElseIf rbcGen(2).Value Then
            'slToFile = sgExportPath & "ExpAlSpt.Txt"
            slToFile = sgDBPath & "Messages\" & "ExpAlSpt.Txt"
        ElseIf rbcGen(3).Value Then
            'slToFile = sgExportPath & "ExpInst.Txt"
            slToFile = sgDBPath & "Messages\" & "ExpInst.Txt"
        Else
            'slToFile = sgExportPath & "ExpStnFd.Txt"
            slToFile = sgDBPath & "Messages\" & "ExpStnFd.Txt"
        End If
    Else
        slToFile = sgDBPath & "Messages\" & "ExpStnFd.Txt"
    End If
    sgMessageFile = slToFile
    
    
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = gDateValue(smTodaysDate) Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            'ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                
                'Print #hmMsg, ""
                
    
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            'ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        'ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    If rbcInterface(0).Value Then
        If rbcGen(1).Value Then
            'Print #hmMsg, "** Export Regional Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            gAutomationAlertAndLogHandler "** Export Regional Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        ElseIf rbcGen(2).Value Then
            'Print #hmMsg, "** Export All Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            gAutomationAlertAndLogHandler "** Export All Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        ElseIf rbcGen(3).Value Then
            'Print #hmMsg, "** Export All Instructions: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            gAutomationAlertAndLogHandler "** Export All Instructions: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        Else
            'Print #hmMsg, "** Export Station Feed-StarGuide: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            gAutomationAlertAndLogHandler "** Export Station Feed-StarGuide: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        End If
    Else
        'Print #hmMsg, "** Export Station Feed-KenCast: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        gAutomationAlertAndLogHandler "** Export Station Feed-KenCast: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    End If
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mRegionExist                    *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test if region exist given non-*
'*                      region rotation                *
'*                                                     *
'*******************************************************
Private Function mRegionExist(tlSortCrf As SORTCRF, llRafCode As Long) As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim tlCrf As CRF
    mRegionExist = False
    tmCrfSrchKey1.sRotType = tlSortCrf.tCrf.sRotType
    tmCrfSrchKey1.iEtfCode = tlSortCrf.tCrf.iEtfCode
    tmCrfSrchKey1.iEnfCode = tlSortCrf.tCrf.iEnfCode
    tmCrfSrchKey1.iAdfCode = tlSortCrf.tCrf.iAdfCode
    tmCrfSrchKey1.lChfCode = tlSortCrf.tCrf.lChfCode
    tmCrfSrchKey1.lFsfCode = 0
    tmCrfSrchKey1.iVefCode = tlSortCrf.tCrf.iVefCode
    tmCrfSrchKey1.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hmCrf, tlCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (tlCrf.sRotType = tlSortCrf.tCrf.sRotType) And (tlCrf.iEtfCode = tlSortCrf.tCrf.iEtfCode) And (tlCrf.iEnfCode = tlSortCrf.tCrf.iEnfCode) And (tlCrf.iAdfCode = tlSortCrf.tCrf.iAdfCode) And (tlCrf.lChfCode = tlSortCrf.tCrf.lChfCode) And (tlCrf.iVefCode = tlSortCrf.tCrf.iVefCode) And (tlCrf.iRotNo > tlSortCrf.tCrf.iRotNo)
        'If Trim$(tlCrf.sZone) = "R" Then
        If (Trim$(tlCrf.sZone) = "R") And (tlCrf.lRafCode = llRafCode) And (tlCrf.sState <> "D") Then
            mRegionExist = True
            Exit Function
        End If
        ilRet = btrGetNext(hmCrf, tlCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilIndex = tlSortCrf.iCombineIndex
    Do While ilIndex >= 0
        'lacRotInfo(ilShow).Caption = "Combining with Rotation #:" & Str$(tgCombineCrf(ilIndex).tCrf.iRotNo) & Str$(tgCombineCrf(ilIndex).lCntrNo) & " " & Trim$(tgCombineCrf(ilIndex).sVehName)
        tmCrfSrchKey1.sRotType = tgCombineCrf(ilIndex).tCrf.sRotType
        tmCrfSrchKey1.iEtfCode = tgCombineCrf(ilIndex).tCrf.iEtfCode
        tmCrfSrchKey1.iEnfCode = tgCombineCrf(ilIndex).tCrf.iEnfCode
        tmCrfSrchKey1.iAdfCode = tgCombineCrf(ilIndex).tCrf.iAdfCode
        tmCrfSrchKey1.lChfCode = tgCombineCrf(ilIndex).tCrf.lChfCode
        tmCrfSrchKey1.lFsfCode = 0
        tmCrfSrchKey1.iVefCode = tgCombineCrf(ilIndex).tCrf.iVefCode
        tmCrfSrchKey1.iRotNo = 32000
        ilRet = btrGetGreaterOrEqual(hmCrf, tlCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (tlCrf.sRotType = tgCombineCrf(ilIndex).tCrf.sRotType) And (tlCrf.iEtfCode = tgCombineCrf(ilIndex).tCrf.iEtfCode) And (tlCrf.iEnfCode = tgCombineCrf(ilIndex).tCrf.iEnfCode) And (tlCrf.iAdfCode = tgCombineCrf(ilIndex).tCrf.iAdfCode) And (tlCrf.lChfCode = tgCombineCrf(ilIndex).tCrf.lChfCode) And (tlCrf.iVefCode = tgCombineCrf(ilIndex).tCrf.iVefCode) And (tlCrf.iRotNo > tgCombineCrf(ilIndex).tCrf.iRotNo)
            'If Trim$(tlCrf.sZone) = "R" Then
            If (Trim$(tlCrf.sZone) = "R") And (tlCrf.lRafCode = llRafCode) And (tlCrf.sState <> "D") Then
                mRegionExist = True
                Exit Function
            End If
            ilRet = btrGetNext(hmCrf, tlCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        ilIndex = tgCombineCrf(ilIndex).iCombineIndex
    Loop
    ilIndex = tlSortCrf.iDuplIndex
    Do While ilIndex >= 0
        'lacRotInfo(ilShow).Caption = "Matching Rotation #:" & Str$(tgDuplCrf(ilIndex).tCrf.iRotNo) & Str$(tgDuplCrf(ilIndex).lCntrNo) & " " & Trim$(tgDuplCrf(ilIndex).sVehName)
        tmCrfSrchKey1.sRotType = tgDuplCrf(ilIndex).tCrf.sRotType
        tmCrfSrchKey1.iEtfCode = tgDuplCrf(ilIndex).tCrf.iEtfCode
        tmCrfSrchKey1.iEnfCode = tgDuplCrf(ilIndex).tCrf.iEnfCode
        tmCrfSrchKey1.iAdfCode = tgDuplCrf(ilIndex).tCrf.iAdfCode
        tmCrfSrchKey1.lChfCode = tgDuplCrf(ilIndex).tCrf.lChfCode
        tmCrfSrchKey1.lFsfCode = 0
        tmCrfSrchKey1.iVefCode = tgDuplCrf(ilIndex).tCrf.iVefCode
        tmCrfSrchKey1.iRotNo = 32000
        ilRet = btrGetGreaterOrEqual(hmCrf, tlCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (tlCrf.sRotType = tgDuplCrf(ilIndex).tCrf.sRotType) And (tlCrf.iEtfCode = tgDuplCrf(ilIndex).tCrf.iEtfCode) And (tlCrf.iEnfCode = tgDuplCrf(ilIndex).tCrf.iEnfCode) And (tlCrf.iAdfCode = tgDuplCrf(ilIndex).tCrf.iAdfCode) And (tlCrf.lChfCode = tgDuplCrf(ilIndex).tCrf.lChfCode) And (tlCrf.iVefCode = tgDuplCrf(ilIndex).tCrf.iVefCode) And (tlCrf.iRotNo > tgDuplCrf(ilIndex).tCrf.iRotNo)
            'If Trim$(tlCrf.sZone) = "R" Then
            If (Trim$(tlCrf.sZone) = "R") And (tlCrf.lRafCode = llRafCode) And (tlCrf.sState <> "D") Then
                mRegionExist = True
                Exit Function
            End If
            ilRet = btrGetNext(hmCrf, tlCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        ilIndex = tgDuplCrf(ilIndex).iDuplIndex
    Loop
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mRotPop                         *
'*                                                     *
'*             Created:8/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain rotation specifications *
'*                      Same code is in BulkFeed.Frm   *
'*                                                     *
'*******************************************************
Private Sub mRotPop()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVsf                         llSifCode                                               *
'******************************************************************************************

'
'   iRet = mRotPop
'   Where:
'
    Dim ilRet As Integer    'Return status
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilOffSet As Integer
    Dim llRevCntrNo As Long
    Dim slRevCntrNo As String
    Dim llRevRotNo As Long
    Dim slRevRotNo As String
    Dim llCntrNo As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilDay As Integer
    Dim ilTest As Integer
    Dim ilExtLen As Integer
    Dim ilUpper As Integer
    Dim ilVehIndex As Integer
    Dim ilVeh As Integer
    Dim ilVpfIndex As Integer
    Dim ilVefSelected As Integer
    Dim llRotStartDate As Long
    Dim llRotEndDate As Long
    Dim ilRotOk As Integer
    Dim ilBit As Integer
    Dim ilVIndex As Integer
    Dim llTstStartDate As Long
    Dim llTstEndDate As Long
    Dim slShortTitle As String

    ReDim tgSortCrf(0 To 0) As SORTCRF
    ReDim tmPSAPromoSortCrf(0 To 0) As SORTCRF
    ilUpper = 0

    lbcRot.Clear
    pbcLbcRot_Paint
    ReDim tgDuplCrf(0 To 0) As DUPLCRF
    ReDim tgCombineCrf(0 To 0) As COMBINECRF

    slStr = Trim$(edcStartDate.Text)
    If (Not gValidDate(slStr)) Or (slStr = "") Then
        imIgnoreVbcChg = True
        vbcRot.Min = 0
        vbcRot.Max = 0
        vbcRot.Value = 0
        imIgnoreVbcChg = False
        Beep
        edcStartDate.SetFocus
        Exit Sub
    End If
    If rbcInterface(0).Value Then
        If rbcGen(0).Value Then
            slStr = gObtainPrevMonday(slStr)
        End If
    Else
        slStr = gObtainPrevMonday(slStr)
    End If
    lmInputStartDate = gDateValue(slStr)
    slStr = Trim$(edcEndDate.Text)
    If (Not gValidDate(slStr)) Or (slStr = "") Then
        imIgnoreVbcChg = True
        vbcRot.Min = 0
        vbcRot.Max = 0
        vbcRot.Value = 0
        imIgnoreVbcChg = False
        Beep
        edcEndDate.SetFocus
        Exit Sub
    End If
    lmInputEndDate = gDateValue(slStr)
    ilVefSelected = False
    For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
        If lbcVeh.Selected(ilLoop) Then
            ilVefSelected = True
        End If
    Next ilLoop
    'If (Not ilVefSelected) Or (rbcGen(1).Value) Then
    If rbcInterface(0).Value Then
        If (Not ilVefSelected) Or (rbcGen(1).Value) Or (rbcGen(2).Value) Then
            imIgnoreVbcChg = True
            vbcRot.Min = 0
            vbcRot.Max = 0
            vbcRot.Value = 0
            imIgnoreVbcChg = False
           Exit Sub
        End If
    Else
        If (Not ilVefSelected) Then
            imIgnoreVbcChg = True
            vbcRot.Min = 0
            vbcRot.Max = 0
            vbcRot.Value = 0
            imIgnoreVbcChg = False
           Exit Sub
        End If
    End If
    btrExtClear hmCrf   'Clear any previous extend operation
    ilExtLen = Len(tmCrf)  'Extract operation record size
    If imTypeIndex = 0 Then
        tmCrfSrchKey1.sRotType = "A"
        tmCrfSrchKey1.iEtfCode = 0
        tmCrfSrchKey1.iEnfCode = 0
        tmCrfSrchKey1.iAdfCode = 0
        tmCrfSrchKey1.lChfCode = 0
        tmCrfSrchKey1.lFsfCode = 0
        tmCrfSrchKey1.iVefCode = 0
        tmCrfSrchKey1.iRotNo = 32000
        ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
        If rbcInterface(0).Value Then
            If Not rbcGen(3).Value Then
                ilOffSet = gFieldOffset("Crf", "CrfAffFdStatus")
                ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "P", 1)
                On Error GoTo mRotPopErr
                gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpStnFd
                On Error GoTo 0
            End If
        Else
            ilOffSet = gFieldOffset("Crf", "CrfKCFdStatus")
            ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "P", 1)
            On Error GoTo mRotPopErr
            gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpStnFd
            On Error GoTo 0
        End If
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "M", 1)
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpStnFd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "S", 1)
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpStnFd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "R", 1)
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpStnFd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "X", 1)
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpStnFd
        On Error GoTo 0
        ilOffSet = 0
        ilRet = btrExtAddField(hmCrf, ilOffSet, ilExtLen)  'Extract start/end time, and days
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddField):" & "Crf.Btr", ExpStnFd
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmClf)    'Extract record
        ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mRotPopErr
            gBtrvErrorMsg ilRet, "mRotPop (btrExtGetNextExt):" & "Clf.Btr", ExpStnFd
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            End If
            Do While ilRet = BTRV_ERR_NONE
                llTstStartDate = lmInputStartDate
                llTstEndDate = lmInputEndDate
                ilRotOk = True
                gUnpackDateLong tmCrf.iStartDate(0), tmCrf.iStartDate(1), llRotStartDate
                gUnpackDateLong tmCrf.iEndDate(0), tmCrf.iEndDate(1), llRotEndDate
                If (llRotEndDate < lmInputStartDate) Or (llRotStartDate > lmInputEndDate) Then
                    'ilRotOk = False
                    If imExptPrevWeek Then
                        llTstStartDate = lmInputStartDate - 7
                        If rbcInterface(0).Value Then
                            If tmCrf.sAffFdStatus <> "R" Then
                                ilRotOk = False
                            Else
                                If (llRotEndDate < (lmInputStartDate - 7)) Or (llRotStartDate > (lmInputEndDate - 7)) Then
                                    ilRotOk = False
                                Else
                                    gUnpackDateLong tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1), llRotEndDate
                                    If lmInputStartDate - 7 = llRotEndDate Then
                                        ilRotOk = False
                                    End If
                                End If
                            End If
                        Else
                            If tmCrf.sKCFdStatus <> "R" Then
                                ilRotOk = False
                            Else
                                If (llRotEndDate < (lmInputStartDate - 7)) Or (llRotStartDate > (lmInputEndDate - 7)) Then
                                    ilRotOk = False
                                Else
                                    gUnpackDateLong tmCrf.iKCFdDate(0), tmCrf.iKCFdDate(1), llRotEndDate
                                    If lmInputStartDate - 7 = llRotEndDate Then
                                        ilRotOk = False
                                    End If
                                End If
                            End If
                        End If
                    Else
                        ilRotOk = False
                    End If
                Else
                    'If rbcGen(0).Value Then
                    '    If Trim$(tmCrf.sZone) = "R" Then
                    '        ilRotOk = False
                    '    End If
                    'ElseIf rbcGen(1).Value Then
                    '    If Trim$(tmCrf.sZone) = "R" Then
                    '        ilRotOk = False
                    '    End If
                    'ElseIf rbcGen(2).Value Then
                    '    If Trim$(tmCrf.sZone) <> "R" Then
                    '        ilRotOk = False
                    '    End If
                    'Else
                    '    ilRotOk = False
                    'End If
                End If
                If ilRotOk Then
                    If tmCrf.sState = "D" Then
                        ilRotOk = False
                    End If
                End If
                If ilRotOk Then
                    If (rbcInterface(0).Value) And (rbcGen(3).Value) Then
                        ilRotOk = mSpotExist(tmCrf.iAdfCode, tmCrf.lChfCode, tmCrf.iVefCode, llTstStartDate, llTstEndDate)
                        If Not ilRotOk Then
                            'If airing rotation, then Ok
                            'tmVefSrchKey.iCode = tmCrf.iVefCode
                            'ilRet = btrGetEqual(hmVef, tmAVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                            'If (ilRet = BTRV_ERR_NONE) Then
                            ilRet = gBinarySearchVef(tmCrf.iVefCode)
                            If ilRet <> -1 Then
                                tmAVef = tgMVef(ilRet)
                                If tmAVef.sType = "A" Then
                                    ilRotOk = True
                                End If
                            End If
                        End If
                    Else
                        If rbcInterface(0).Value Then
                            gUnpackDateLong tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1), llRotEndDate
                            If lmInputStartDate = llRotEndDate Then
                                ilRotOk = False
                            Else
                                'ilRotOk = mSpotExist(tmCrf.iAdfCode, tmCrf.lChfCode, tmCrf.iVefCode)
                                'Test bit map to see if week sent previously
                                'Test bit map to see if week sent previously
                                If llRotEndDate > lmInputStartDate Then
                                    ilBit = (llRotEndDate - lmInputStartDate) \ 7 + 1
                                    Select Case ilBit
                                        Case 1
                                            If (tmCrf.iAffFdWk And 64) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 2
                                            If (tmCrf.iAffFdWk And 32) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 3
                                            If (tmCrf.iAffFdWk And 16) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 4
                                            If (tmCrf.iAffFdWk And 8) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 5
                                            If (tmCrf.iAffFdWk And 4) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 6
                                            If (tmCrf.iAffFdWk And 2) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 7
                                            If (tmCrf.iAffFdWk And 1) <> 0 Then
                                                ilRotOk = False
                                            End If
                                    End Select
                                End If
                                If ilRotOk Then
                                    ilRotOk = mSpotExist(tmCrf.iAdfCode, tmCrf.lChfCode, tmCrf.iVefCode, llTstStartDate, llTstEndDate)
                                    If Not ilRotOk Then
                                        'If airing rotation, then Ok
                                        'tmVefSrchKey.iCode = tmCrf.iVefCode
                                        'ilRet = btrGetEqual(hmVef, tmAVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                        'If (ilRet = BTRV_ERR_NONE) Then
                                        ilRet = gBinarySearchVef(tmCrf.iVefCode)
                                        If ilRet <> -1 Then
                                            tmAVef = tgMVef(ilRet)
                                            If tmAVef.sType = "A" Then
                                                ilRotOk = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            gUnpackDateLong tmCrf.iKCFdDate(0), tmCrf.iKCFdDate(1), llRotEndDate
                            If lmInputStartDate = llRotEndDate Then
                                ilRotOk = False
                            Else
                                'ilRotOk = mSpotExist(tmCrf.iAdfCode, tmCrf.lChfCode, tmCrf.iVefCode)
                                'Test bit map to see if week sent previously
                                'Test bit map to see if week sent previously
                                If llRotEndDate > lmInputStartDate Then
                                    ilBit = (llRotEndDate - lmInputStartDate) \ 7 + 1
                                    Select Case ilBit
                                        Case 1
                                            If (tmCrf.iKCFdWk And 64) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 2
                                            If (tmCrf.iKCFdWk And 32) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 3
                                            If (tmCrf.iKCFdWk And 16) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 4
                                            If (tmCrf.iKCFdWk And 8) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 5
                                            If (tmCrf.iKCFdWk And 4) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 6
                                            If (tmCrf.iKCFdWk And 2) <> 0 Then
                                                ilRotOk = False
                                            End If
                                        Case 7
                                            If (tmCrf.iKCFdWk And 1) <> 0 Then
                                                ilRotOk = False
                                            End If
                                    End Select
                                End If
                                If ilRotOk Then
                                    ilRotOk = mSpotExist(tmCrf.iAdfCode, tmCrf.lChfCode, tmCrf.iVefCode, llTstStartDate, llTstEndDate)
                                    If Not ilRotOk Then
                                        'If airing rotation, then Ok
                                        'tmVefSrchKey.iCode = tmCrf.iVefCode
                                        'ilRet = btrGetEqual(hmVef, tmAVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                        'If (ilRet = BTRV_ERR_NONE) Then
                                        ilRet = gBinarySearchVef(tmCrf.iVefCode)
                                        If ilRet <> -1 Then
                                            tmAVef = tgMVef(ilRet)
                                            If tmAVef.sType = "A" Then
                                                ilRotOk = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If ilRotOk Then
                    ilVefSelected = False
                    For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
                        If lbcVeh.Selected(ilLoop) Then
                            slNameCode = tmVehCode(ilLoop).sKey    'lbcVehCode.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                            If tmCrf.iVefCode = Val(slCode) Then
                                ilVefSelected = True
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If Not ilVefSelected Then
                        For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
                            If lbcVeh.Selected(ilLoop) Then
                                slNameCode = tmVehCode(ilLoop).sKey    'lbcVehCode.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                'Check if airing
                                ilVIndex = mFindVpfIndex(Val(slCode)) 'tmVef(ilVehIndex).iCode)
                                If ilVIndex >= 0 Then
                                    ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                                    Do While ilVIndex >= 0
                                        If tmCrf.iVefCode = tmSALink(ilVIndex).iVefCode Then
                                            ilVefSelected = True
                                            Exit Do
                                        End If
                                        ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                                    Loop
                                End If
                            End If
                        Next ilLoop
                    End If
                    If ilVefSelected Then
                        If tmChf.lCode <> tmCrf.lChfCode Then
                            tmChfSrchKey.lCode = tmCrf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo mRotPopErr
                            gBtrvErrorMsg ilRet, "mRotPop (btrGetEqual):" & "Chf.Btr", ExpStnFd
                            On Error GoTo 0
                        End If
                        'Removing testing of PSA at this point,  will place PSA/Promoms into tmPSAPromoSortCrf at end
                        '11/16/03
                        'If (tmChf.sType <> "S") And (tmChf.sType <> "M") Then
                            ilRet = mDuplRotation(ilVpfIndex, llRecPos)
                            ilVehIndex = -1
                            For ilVeh = 0 To UBound(tmVef) - 1 Step 1
                                If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
                                    ilVehIndex = ilVeh
                                    Exit For
                                End If
                            Next ilVeh
                            If ilVehIndex >= 0 Then
                                If ilRet = 2 Then   'Match
                                    'tgDuplCrf(UBound(tgDuplCrf)).lCntrNo = tmChf.lCntrNo
                                    'tgDuplCrf(UBound(tgDuplCrf)).sVehName = tmVef(ilVehIndex).sName
                                    'tgDuplCrf(UBound(tgDuplCrf)).tCrf = tmCrf
                                    'tgDuplCrf(UBound(tgDuplCrf)).lCrfRecPos = llRecPos
                                    'tgDuplCrf(UBound(tgDuplCrf)).iDuplIndex = -1
                                    'tgDuplCrf(UBound(tgDuplCrf)).iVpfIndex = ilVpfIndex
                                    'ReDim Preserve tgDuplCrf(0 To UBound(tgDuplCrf) + 1) As DUPLCRF
                                ElseIf ilRet = 1 Then   'Combine
                                    'tgCombineCrf(UBound(tgCombineCrf)).lCntrNo = tmChf.lCntrNo
                                    'tgCombineCrf(UBound(tgCombineCrf)).sVehName = tmVef(ilVehIndex).sName
                                    'tgCombineCrf(UBound(tgCombineCrf)).tCrf = tmCrf
                                    'tgCombineCrf(UBound(tgCombineCrf)).lCrfRecPos = llRecPos
                                    'tgCombineCrf(UBound(tgCombineCrf)).iCombineIndex = -1
                                    'tgCombineCrf(UBound(tgCombineCrf)).iVpfIndex = ilVpfIndex
                                    'ReDim Preserve tgCombineCrf(0 To UBound(tgCombineCrf) + 1) As COMBINECRF
                                Else
                                    llRevCntrNo = 99999999 - tmChf.lCntrNo
                                    slRevCntrNo = Trim$(str$(llRevCntrNo))
                                    Do While Len(slRevCntrNo) < 8
                                        slRevCntrNo = "0" & slRevCntrNo
                                    Loop
                                    'Scan for vehicle
                                    'For ilVeh = 0 To UBound(tmVef) - 1 Step 1
                                    '    If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
                                    '        ilVehIndex = ilVeh
                                    '        Exit For
                                    '    End If
                                    'Next ilVeh
                                    If tmAdf.iCode <> tmCrf.iAdfCode Then
                                        tmAdfSrchKey.iCode = tmCrf.iAdfCode
                                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        On Error GoTo mRotPopErr
                                        gBtrvErrorMsg ilRet, "mRotPop (btrGetEqual):" & "Adf.Btr", ExpStnFd
                                        On Error GoTo 0
                                    End If
                                    slName = tmVef(ilVehIndex).sName
                                    llRevRotNo = 99999 - tmCrf.iRotNo
                                    slRevRotNo = Trim$(str$(llRevRotNo))
                                    Do While Len(slRevRotNo) < 6
                                        slRevRotNo = "0" & slRevRotNo
                                    Loop
                                    ilUpper = UBound(tgSortCrf)
                                    tgSortCrf(ilUpper).sKey = tmAdf.sName & "|" & slRevCntrNo & "|" & tmVef(ilVehIndex).sName & "|" & slRevRotNo
                                    tgSortCrf(ilUpper).lCntrNo = tmChf.lCntrNo
'                                    llSifCode = 0
'                                    If tmChf.lVefCode < 0 Then
'                                        tmVsfSrchKey.lCode = -tmChf.lVefCode
'                                        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                                        Do While ilRet = BTRV_ERR_NONE
'                                            For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
'                                                If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
'                                                    If tmVsf.lFSComm(ilVsf) > 0 Then
'                                                        llSifCode = tmVsf.lFSComm(ilVsf)
'                                                    End If
'                                                    Exit For
'                                                End If
'                                            Next ilVsf
'                                            If llSifCode <> 0 Then
'                                                Exit Do
'                                            End If
'                                            If tmVsf.lLkVsfCode <= 0 Then
'                                                Exit Do
'                                            End If
'                                            tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
'                                            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                                        Loop
'                                    End If
                                    slShortTitle = mGetShortTitle(tmChf, tmAdf, tmCrf.iVefCode)
                                    tgSortCrf(ilUpper).sCntrProd = slShortTitle 'gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf) 'tmChf.sProduct
                                    tgSortCrf(ilUpper).sType = tmChf.sType
                                    tgSortCrf(ilUpper).lCrfRecPos = llRecPos
                                    tgSortCrf(ilUpper).iSelected = False
                                    tgSortCrf(ilUpper).iCombineIndex = -1
                                    tgSortCrf(ilUpper).iDuplIndex = -1
                                    tgSortCrf(ilUpper).iVpfIndex = ilVpfIndex
                                    tgSortCrf(ilUpper).tCrf = tmCrf
                                    'Save PSA/Promo for processing later in mPSAPromoProcess
                                    If (tgSortCrf(ilUpper).sType = "S") Or (tgSortCrf(ilUpper).sType = "M") Then
                                        tmPSAPromoSortCrf(UBound(tmPSAPromoSortCrf)) = tgSortCrf(ilUpper)
                                        ReDim Preserve tmPSAPromoSortCrf(0 To UBound(tmPSAPromoSortCrf) + 1) As SORTCRF
                                    Else
                                        ReDim Preserve tgSortCrf(0 To ilUpper + 1) As SORTCRF
                                        ilUpper = ilUpper + 1
                                    End If
                                End If
                            End If
                        'End If
                    End If
                End If
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                End If
            Loop
        End If
        'If (rbcGen(0).Value) Or (rbcGen(1).Value) Then
            ilTest = LBound(tgSortCrf)
            Do While ilTest < UBound(tgSortCrf)
                ilRotOk = True
                ''If rbcGen(0).Value Then
                ''    'Test if Regional copy exist- If so don't include
                ''    If mRegionExist(tgSortCrf(ilTest)) Then
                ''        ilRotOk = False
                ''    End If
                ''ElseIf rbcGen(1).Value Then
                ''    'Test if Regional copy exist- If so include
                ''    If Not mRegionExist(tgSortCrf(ilTest)) Then
                ''        ilRotOk = False
                ''    End If
                ''End If
                'If Trim$(tgSortCrf(ilTest).tCrf.sZone) = "" Then
                '    If mRegionExist(tgSortCrf(ilTest)) Then
                '        ilRotOk = False
                '    End If
                'End If
                If Not ilRotOk Then
                    For ilLoop = ilTest To UBound(tgSortCrf) - 1 Step 1
                        tgSortCrf(ilLoop) = tgSortCrf(ilLoop + 1)
                    Next ilLoop
                    ReDim Preserve tgSortCrf(0 To UBound(tgSortCrf) - 1) As SORTCRF
                Else
                    ilTest = ilTest + 1
                End If
            Loop
        'End If
        ilUpper = UBound(tgSortCrf)
        If ilUpper > 0 Then
            ArraySortTyp fnAV(tgSortCrf(), 0), ilUpper, 0, LenB(tgSortCrf(0)), 0, LenB(tgSortCrf(0).sKey), 0
        End If
        imLastIndex = -1
        imIgnoreVbcChg = True
        vbcRot.Min = 0
        If ilUpper > vbcRot.LargeChange + 1 Then
            vbcRot.Max = ilUpper - vbcRot.LargeChange - 1
        Else
            vbcRot.Max = 0
        End If
        vbcRot.Value = 0
        imIgnoreVbcChg = False
        btrExtClear hmCrf   'Clear any previous extend operation
        For ilLoop = 0 To ilUpper - 1 Step 1
            slNameCode = tgSortCrf(ilLoop).sKey
            tmCrf = tgSortCrf(ilLoop).tCrf
            ilRet = gParseItem(slNameCode, 1, "|", slName)
            slStr = slName & "|"
            If ilRet <> CP_MSG_NONE Then
                slName = "Missing"
            End If
            llCntrNo = tgSortCrf(ilLoop).lCntrNo
            slStr = slStr & Trim$(str$(llCntrNo)) & "|"
            ilRet = gParseItem(slNameCode, 3, "|", slName)
            If ilRet <> CP_MSG_NONE Then
                slName = "Missing"
            End If
            slStr = slStr & Left$(slName, 10) & "|"
            slStr = slStr & Trim$(tmCrf.sZone) & "|"
            gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slStartDate
            If slStartDate <> "" Then
                slStartDate = Left$(slStartDate, Len(slStartDate) - 3)
            End If
            gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slEndDate
            If slEndDate <> "" Then
                slEndDate = Left$(slEndDate, Len(slEndDate) - 3)
            End If
            slStr = slStr & slStartDate & "-" & slEndDate & "|"
            gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slStartTime
            gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slEndTime
            slStr = slStr & slStartTime & "-" & slEndTime & "|"
            For ilDay = 0 To 6 Step 1
                slStr = slStr & tmCrf.sDay(ilDay) '& "|"
            Next ilDay
            slStr = slStr & "|"
            slStr = slStr & Trim$(str$(tmCrf.iLen)) & "|"
            If (tmCrf.sInOut = "I") Or (tmCrf.sInOut = "O") Then
                If tmAnf.iCode <> tmCrf.ianfCode Then
                    tmAnfSrchKey.iCode = tmCrf.ianfCode
                    ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    On Error GoTo mRotPopErr
                    gBtrvErrorMsg ilRet, "mRotPop (btrGetEqual):" & "Anf.Btr", ExpStnFd
                    On Error GoTo 0
                End If
                If tmCrf.sInOut = "O" Then
                    slName = "O" & Trim$(tmAnf.sName)
                Else
                    slName = Trim$(tmAnf.sName)
                End If
            Else
                slName = "All avails"
            End If
            slStr = slStr & slName & "|"
            Select Case tmCrf.sRotType
                Case "A"
                    slStr = slStr & "CS " & "|"
                Case "O"
                    slStr = slStr & "OBB" & "|"
                Case "C"
                    slStr = slStr & "CBB" & "|"
                Case "E"
                    slStr = slStr & "ABB" & "|"
                Case Else
                    slStr = slStr & " |"
            End Select
            If tmCrf.lCsfCode > 0 Then
                slStr = slStr & "C"
            Else
                slStr = slStr & " "
            End If
            If lbcRot.ListCount < vbcRot.LargeChange + 1 Then
                lbcRot.AddItem slStr
            End If
            tgSortCrf(ilLoop).sKey = slStr
        Next ilLoop
    End If
    pbcLbcRot_Paint
    If ckcAll.Value = vbChecked Then
        ckcAll_Click
    End If
    Exit Sub
mRotPopErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetAffFdDate                   *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set Affiliate Feed date and    *
'*                      week bit map                   *
'*                                                     *
'*******************************************************
Private Sub mSetAffFdDate()
    Dim llAffFdDate As Long
    Dim ilBit As Integer
    Dim llLoop As Long
    gPackDate smTranDate, tmCrf.iAffTranDate(0), tmCrf.iAffTranDate(1)
    gUnpackDateLong tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1), llAffFdDate
    If llAffFdDate <= 0 Then
        tmCrf.iAffFdWk = 64
        gPackDateLong lmInputStartDate, tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1)
    Else
        If lmInputStartDate <= llAffFdDate Then
            ilBit = (llAffFdDate - lmInputStartDate) \ 7 + 1
            Select Case ilBit
                Case 1
                    tmCrf.iAffFdWk = tmCrf.iAffFdWk Or 64
                Case 2
                    tmCrf.iAffFdWk = tmCrf.iAffFdWk Or 32
                Case 3
                    tmCrf.iAffFdWk = tmCrf.iAffFdWk Or 16
                Case 4
                    tmCrf.iAffFdWk = tmCrf.iAffFdWk Or 8
                Case 5
                    tmCrf.iAffFdWk = tmCrf.iAffFdWk Or 4
                Case 6
                    tmCrf.iAffFdWk = tmCrf.iAffFdWk Or 2
                Case 7
                    tmCrf.iAffFdWk = tmCrf.iAffFdWk Or 1
            End Select
        ElseIf lmInputStartDate > llAffFdDate Then
            For llLoop = llAffFdDate To lmInputStartDate - 1 Step 7
                tmCrf.iAffFdWk = tmCrf.iAffFdWk \ 2
                If tmCrf.iAffFdWk = 0 Then
                    Exit For
                End If
            Next llLoop
            tmCrf.iAffFdWk = tmCrf.iAffFdWk Or 64
            gPackDateLong lmInputStartDate, tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1)
        End If
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetKCFdDate                   *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set Affiliate Feed date and    *
'*                      week bit map                   *
'*                                                     *
'*******************************************************
Private Sub mSetKCFdDate()
    Dim llKCFdDate As Long
    Dim ilBit As Integer
    Dim llLoop As Long
    gPackDate smTranDate, tmCrf.iKCTranDate(0), tmCrf.iKCTranDate(1)
    gUnpackDateLong tmCrf.iKCFdDate(0), tmCrf.iKCFdDate(1), llKCFdDate
    If llKCFdDate <= 0 Then
        tmCrf.iKCFdWk = 64
        gPackDateLong lmInputStartDate, tmCrf.iKCFdDate(0), tmCrf.iKCFdDate(1)
    Else
        If lmInputStartDate <= llKCFdDate Then
            ilBit = (llKCFdDate - lmInputStartDate) \ 7 + 1
            Select Case ilBit
                Case 1
                    tmCrf.iKCFdWk = tmCrf.iKCFdWk Or 64
                Case 2
                    tmCrf.iKCFdWk = tmCrf.iKCFdWk Or 32
                Case 3
                    tmCrf.iKCFdWk = tmCrf.iKCFdWk Or 16
                Case 4
                    tmCrf.iKCFdWk = tmCrf.iKCFdWk Or 8
                Case 5
                    tmCrf.iKCFdWk = tmCrf.iKCFdWk Or 4
                Case 6
                    tmCrf.iKCFdWk = tmCrf.iKCFdWk Or 2
                Case 7
                    tmCrf.iKCFdWk = tmCrf.iKCFdWk Or 1
            End Select
        ElseIf lmInputStartDate > llKCFdDate Then
            For llLoop = llKCFdDate To lmInputStartDate - 1 Step 7
                tmCrf.iKCFdWk = tmCrf.iKCFdWk \ 2
                If tmCrf.iKCFdWk = 0 Then
                    Exit For
                End If
            Next llLoop
            tmCrf.iKCFdWk = tmCrf.iKCFdWk Or 64
            gPackDateLong lmInputStartDate, tmCrf.iKCFdDate(0), tmCrf.iKCFdDate(1)
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set button state               *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
    Dim slTranDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilLoop As Integer
    Dim ilSelected As Integer
    ilSelected = False
    'If rbcGen(1).Value Then
    If (Not rbcInterface(0).Value) And (Not rbcInterface(1).Value) Then
        cmcExport.Enabled = False
        cmcGetRot.Enabled = False
        Exit Sub
    End If
    If (rbcInterface(0).Value) And ((rbcGen(1).Value) Or (rbcGen(2).Value)) Then
        For ilLoop = 0 To lbcRegVeh.ListCount - 1 Step 1
            If lbcRegVeh.Selected(ilLoop) Then
                ilSelected = True
                Exit For
            End If
        Next ilLoop
    Else
        For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
            If tgSortCrf(ilLoop).iSelected Then
                ilSelected = True
                Exit For
            End If
        Next ilLoop
        If lbcVeh.SelCount > 0 Then
            cmcGetRot.Enabled = True
        Else
            cmcGetRot.Enabled = False
        End If
    End If
    slTranDate = edcTranDate.Text
    slStartDate = edcStartDate.Text
    slEndDate = edcEndDate.Text
    If (ilSelected) And (gValidDate(slTranDate)) And (gValidDate(slStartDate)) And (gValidDate(slEndDate)) Then
        cmcExport.Enabled = True
    Else
        cmcExport.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mShowRotInfo                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show Dulpicate and Combined    *
'*                      rotations                      *
'*                                                     *
'*******************************************************
Private Sub mShowRotInfo()
    Dim ilShow As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilButtonIndex As Integer
    Dim slDate As String
    Dim ilRet As Integer
    ilButtonIndex = imButtonIndex
    If (imButtonIndex < 0) And (imButtonIndex > UBound(tgSortCrf) - 1) Then
        imButtonIndex = -1
        plcRotInfo.Visible = False
        Exit Sub
    End If
    For ilLoop = 0 To 4 Step 1
        lacRotInfo(ilLoop).Caption = ""
    Next ilLoop
    ilShow = 0
    lacRotInfo(ilShow).Caption = "No Matching or Combinations"
    ilIndex = tgSortCrf(ilButtonIndex).iCombineIndex
    Do While ilIndex >= 0
        lacRotInfo(ilShow).Caption = "Combining with Rotation #:" & str$(tgCombineCrf(ilIndex).tCrf.iRotNo) & str$(tgCombineCrf(ilIndex).lCntrNo) & " " & Trim$(tgCombineCrf(ilIndex).sVehName)
        ilShow = ilShow + 1
        If ilShow > 4 Then
            Exit Sub
        End If
        ilIndex = tgCombineCrf(ilIndex).iCombineIndex
    Loop
    ilIndex = tgSortCrf(ilButtonIndex).iDuplIndex
    Do While ilIndex >= 0
        lacRotInfo(ilShow).Caption = "Matching Rotation #:" & str$(tgDuplCrf(ilIndex).tCrf.iRotNo) & str$(tgDuplCrf(ilIndex).lCntrNo) & " " & Trim$(tgDuplCrf(ilIndex).sVehName)
        ilShow = ilShow + 1
        If ilShow > 4 Then
            Exit Sub
        End If
        ilIndex = tgDuplCrf(ilIndex).iDuplIndex
    Loop
    DoEvents
    If (imButtonIndex < 0) And (imButtonIndex > UBound(tgSortCrf) - 1) Then
        imButtonIndex = -1
        plcRotInfo.Visible = False
        Exit Sub
    End If
    If (tgSpf.sUseProdSptScr <> "P") Then
        If ilShow > 0 Then
            lacRotInfo(ilShow).Caption = "Product: " & Trim$(tgSortCrf(ilButtonIndex).sCntrProd)
        Else
            lacRotInfo(1).Caption = "Product: " & Trim$(tgSortCrf(ilButtonIndex).sCntrProd)
            ilShow = 1
        End If
    Else
        If ilShow > 0 Then
            lacRotInfo(ilShow).Caption = "Short Title: " & Trim$(tgSortCrf(ilButtonIndex).sCntrProd)
        Else
            lacRotInfo(1).Caption = "Short Title: " & Trim$(tgSortCrf(ilButtonIndex).sCntrProd)
            ilShow = 1
        End If
    End If
    ilShow = ilShow + 1
    If ilShow <= 4 Then
        If (Left$(tgSortCrf(ilButtonIndex).tCrf.sZone, 1) = "R") And (tgSortCrf(ilButtonIndex).tCrf.lRafCode > 0) Then
            'If tgSortCrf(ilButtonIndex).tCrf.sAffFdStatus = "R" Then
            '   lacRotInfo(ilShow).Caption = "Station Feed:  Ready to Send"
            '   ilShow = ilShow + 1
            'ElseIf tgSortCrf(ilButtonIndex).tCrf.sAffFdStatus = "P" Then
            '   lacRotInfo(ilShow).Caption = "Station Feed:  Suppress "
            '   ilShow = ilShow + 1
            'ElseIf (tgSortCrf(ilButtonIndex).tCrf.sAffFdStatus = "S") Or (tgSortCrf(ilButtonIndex).tCrf.sAffFdStatus = "X") Then
            '   gUnpackDate tgSortCrf(ilButtonIndex).tCrf.iAffFdDate(0), tgSortCrf(ilButtonIndex).tCrf.iAffFdDate(1), slDate
            '   lacRotInfo(ilShow).Caption = "Station Feed: Sent " & slDate
            '   ilShow = ilShow + 1
            'Else
            '   lacRotInfo(ilShow).Caption = ""
            'End If
            'If ilShow <= 4 Then
            '   tmRafSrchKey.lCode = tgSortCrf(ilButtonIndex).tCrf.lRafCode
            '   ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            '   If ilRet = BTRV_ERR_NONE Then
            '       lacRotInfo(ilShow).Caption = Trim$(tmRaf.sName)
            '   Else
            '       ilShow = ilShow - 1
            '   End If
            'End If
            tmRafSrchKey.lCode = tgSortCrf(ilButtonIndex).tCrf.lRafCode
            ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmRaf.sName = ""
            End If
            If rbcInterface(0).Value Then
                If tgSortCrf(ilButtonIndex).tCrf.sAffFdStatus = "R" Then
                    lacRotInfo(ilShow).Caption = Trim$(tmRaf.sName) & " StarGuide Feed:  Ready to Send"
                ElseIf tgSortCrf(ilButtonIndex).tCrf.sAffFdStatus = "P" Then
                    lacRotInfo(ilShow).Caption = Trim$(tmRaf.sName) & " StarGuide Feed:  Suppress "
                ElseIf (tgSortCrf(ilButtonIndex).tCrf.sAffFdStatus = "S") Or (tgSortCrf(ilButtonIndex).tCrf.sAffFdStatus = "X") Then
                    gUnpackDate tgSortCrf(ilButtonIndex).tCrf.iAffFdDate(0), tgSortCrf(ilButtonIndex).tCrf.iAffFdDate(1), slDate
                    lacRotInfo(ilShow).Caption = Trim$(tmRaf.sName) & " StarGuide Feed: Sent " & slDate
                Else
                    lacRotInfo(ilShow).Caption = Trim$(tmRaf.sName)
                End If
            Else
                If tgSortCrf(ilButtonIndex).tCrf.sKCFdStatus = "R" Then
                    lacRotInfo(ilShow).Caption = Trim$(tmRaf.sName) & " KenCast Feed:  Ready to Send"
                ElseIf tgSortCrf(ilButtonIndex).tCrf.sKCFdStatus = "P" Then
                    lacRotInfo(ilShow).Caption = Trim$(tmRaf.sName) & " KenCast Feed:  Suppress "
                ElseIf (tgSortCrf(ilButtonIndex).tCrf.sKCFdStatus = "S") Or (tgSortCrf(ilButtonIndex).tCrf.sKCFdStatus = "X") Then
                    gUnpackDate tgSortCrf(ilButtonIndex).tCrf.iKCFdDate(0), tgSortCrf(ilButtonIndex).tCrf.iKCFdDate(1), slDate
                    lacRotInfo(ilShow).Caption = Trim$(tmRaf.sName) & " KenCast Feed: Sent " & slDate
                Else
                    lacRotInfo(ilShow).Caption = Trim$(tmRaf.sName)
                End If
            End If
        Else
            If rbcInterface(0).Value Then
                If (tgSortCrf(ilButtonIndex).tCrf.sAffFdStatus = "S") Or (tgSortCrf(ilButtonIndex).tCrf.sAffFdStatus = "X") Then
                    gUnpackDate tgSortCrf(ilButtonIndex).tCrf.iAffFdDate(0), tgSortCrf(ilButtonIndex).tCrf.iAffFdDate(1), slDate
                    lacRotInfo(ilShow).Caption = "StarGuide Feed: Sent " & slDate
                End If
            Else
                If (tgSortCrf(ilButtonIndex).tCrf.sKCFdStatus = "S") Or (tgSortCrf(ilButtonIndex).tCrf.sKCFdStatus = "X") Then
                    gUnpackDate tgSortCrf(ilButtonIndex).tCrf.iKCFdDate(0), tgSortCrf(ilButtonIndex).tCrf.iKCFdDate(1), slDate
                    lacRotInfo(ilShow).Caption = "KenCast Feed: Sent " & slDate
                End If
            End If
        End If
    End If
    plcRotInfo.Visible = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpotExist                      *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if spot exist for    *
'*                      contract within Dates          *
'*                                                     *
'*******************************************************
Private Function mSpotExist(ilAdfCode As Integer, llChfCode As Long, ilVefCode As Integer, llInputStartDate As Long, llInputEndDate As Long) As Integer
    Dim ilRet As Integer
    Dim tlSsfSrchKey As SSFKEY0 'SSF key record image
    Dim ilSsfRecLen As Integer  'SSF record length
    Dim llDate As Long
    Dim ilSsfDate0 As Integer
    Dim ilSsfDate1 As Integer
    Dim ilType As Integer
    Dim ilEvt As Integer
    ilType = 0
    For llDate = llInputStartDate To llInputEndDate Step 1
        ilSsfRecLen = Len(tgSsf(0)) 'Max size of variable length record
        gPackDateLong llDate, ilSsfDate0, ilSsfDate1
        tlSsfSrchKey.iType = ilType
        tlSsfSrchKey.iVefCode = ilVefCode
        tlSsfSrchKey.iDate(0) = ilSsfDate0
        tlSsfSrchKey.iDate(1) = ilSsfDate1
        tlSsfSrchKey.iStartTime(0) = 0
        tlSsfSrchKey.iStartTime(1) = 0
        ilRet = gSSFGetEqual(hmSsf, tgSsf(0), ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        Do While (ilRet = BTRV_ERR_NONE) And (tgSsf(0).iType = ilType) And (tgSsf(0).iVefCode = ilVefCode) And (tgSsf(0).iDate(0) = ilSsfDate0) And (tgSsf(0).iDate(1) = ilSsfDate1)
            ilEvt = 1
            Do While ilEvt <= tgSsf(0).iCount
               LSet tmSpot = tgSsf(0).tPas(ADJSSFPASBZ + ilEvt)
                If ((tmSpot.iRecType And &HF) >= 10) And ((tmSpot.iRecType And &HF) <= 11) Then
                    If tmSpot.iAdfCode = ilAdfCode Then
                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                        If (ilRet = BTRV_ERR_NONE) And (tmSdf.lChfCode = llChfCode) Then
                            mSpotExist = True
                            Exit Function
                        End If
                    End If
                End If
                ilEvt = ilEvt + 1
            Loop
            ilSsfRecLen = Len(tgSsf(0)) 'Max size of variable length record
            ilRet = gSSFGetNext(hmSsf, tgSsf(0), ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next llDate
    For llDate = llInputStartDate To llInputEndDate Step 1
        gPackDateLong llDate, ilSsfDate0, ilSsfDate1
        tmSdfSrchKey1.iVefCode = ilVefCode
        tmSdfSrchKey1.iDate(0) = ilSsfDate0
        tmSdfSrchKey1.iDate(1) = ilSsfDate1
        tmSdfSrchKey1.iTime(0) = 0
        tmSdfSrchKey1.iTime(1) = 0
        tmSdfSrchKey1.sSchStatus = "M"
        'ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.iDate(0) = ilSsfDate0) And (tmSdf.iDate(1) = ilSsfDate1)
            If (tmSdf.sSchStatus = "M") Or (tmSdf.sSchStatus = "C") Or (tmSdf.sSchStatus = "H") Or (tmSdf.sSchStatus = "U") Or (tmSdf.sSchStatus = "R") Then
                If tmSdf.lChfCode = llChfCode Then
                    mSpotExist = True
                    Exit Function
                End If
            End If
            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next llDate
    mSpotExist = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    csiSetValue imWaitCount, imTimeDelay, imLockValue, imTranLog
    '5/31/06:  Show comments on first occurrance only

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpStnFd
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the vehicles to generate *
'*                     bulk feed                       *
'*                                                     *
'*******************************************************
Private Sub mVehPop(ilfirstTime As Integer)
'
'   tmVef will contain the primary vehicle (primary vehicle is
'        the one within the group that is earliest alphabetically)
'        Find no group defined for the vehicle, then tmVef contain
'        the vehicle and tmVpfInfo will contain no links
'   tmVpfInfo will contain the other vehicles grouped with the primary
'   tVpf.iGLink will only contain primary and vehicles without groups only
'               the other vehicles within a group are removed from iGLink
'
'   For Prime remove 24-Hour Format from link table as the spots only
'   map into 12m-6am and rotation should not be generated for 24-hour format
'
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim slName As String
    Dim slNameCode As String
    Dim slChar As String
    Dim slPrefix As String
    Dim slSecond As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilTest As Integer
    Dim ilOk As Integer
    Dim ilLetter As Integer
    Dim ilFound As Integer
    Dim ilUse As Integer
    Dim ilVpfIndex As Integer
    Dim ilLoop1 As Integer
    Dim ilLoop2 As Integer
    Dim ilNextLk As Integer
    Dim slDate As String
    Dim ilPrevIndex As Integer
    Dim slStr As String
    Dim tlVef As VEF
    Dim tlSVef As VEF

    lbcVehicle.Clear
    lbcVehicleCode.Clear
    pbclbcVehicle_Paint
    DoEvents
    If ilfirstTime Then

        imVefRecLen = Len(tlVef)
        'Determine which vehicles are to be combined (same bulk group Number)
        ReDim tmVpfInfo(LBound(tgVpf) To LBound(tgVpf)) As VPFINFO
        ReDim tmLkVehInfo(0 To 0) As LKVEHINFO
        For ilLoop = LBound(tgVpf) To UBound(tgVpf) Step 1
            'tmVefSrchKey.iCode = tgVpf(ilLoop).iVefKCode
            'ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            ilRet = gBinarySearchVef(tgVpf(ilLoop).iVefKCode)
            If ilRet <> -1 Then
                tlVef = tgMVef(ilRet)
                'If ((tlVef.sType = "A") Or (tlVef.sType = "C") Or (tlVef.sType = "S")) And (tgVpf(ilLoop).sGGroupNo <> "-") Then
                If (((tlVef.sType = "A") Or (tlVef.sType = "C")) And (Trim$(tgVpf(ilLoop).sStnFdCode) <> "") And (Asc(tgVpf(ilLoop).sStnFdCode) <> 0)) Or (tlVef.sType = "S") Then
                    ilFound = False
                    'If tgVpf(ilLoop).sGGroupNo <> " " Then  'Group number only defined for airing and conventional vehicles
                        For ilTest = LBound(tmVpfInfo) To UBound(tmVpfInfo) - 1 Step 1
                            'If tmVpfInfo(ilTest).tVpf.sGGroupNo = tgVpf(ilLoop).sGGroupNo Then
                            If (tmVpfInfo(ilTest).tVpf.sStnFdCode = tgVpf(ilLoop).sStnFdCode) And (tlVef.sType <> "S") Then
                                ilFound = True
                                slStr = Trim$(tlVef.sName)
                                'tmVefSrchKey.iCode = tmVpfInfo(ilTest).tVpf.iVefKCode
                                'ilRet = btrGetEqual(hmVef, tlSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                ilRet = gBinarySearchVef(tmVpfInfo(ilTest).tVpf.iVefKCode)
                                If ilRet <> -1 Then
                                    tlSVef = tgMVef(ilRet)
                                Else
                                    tlSVef.sName = "Missing VEFCode" & str$(tmVpfInfo(ilTest).tVpf.iVefKCode)
                                End If
                                slStr = slStr & " and " & Trim$(tlSVef.sName)
                                ''MsgBox "Station Feed Codes (" & tmVpfInfo(ilTest).tVpf.sStnFdCode & ") Match for " & slStr & ", " & Trim$(tlSVef.sName) & " ignored", vbExclamation, "Name Error"
                                gAutomationAlertAndLogHandler "Station Feed Codes (" & tmVpfInfo(ilTest).tVpf.sStnFdCode & ") Match for " & slStr & ", " & Trim$(tlSVef.sName) & " ignored", vbExclamation, "Name Error"
                                ''Save primary alphabetically
                                'tmVefSrchKey.iCode = tmVpfInfo(ilTest).tVpf.iVefKCode
                                'ilRet = btrGetEqual(hmVef, tlSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                'If StrComp(tlSVef.sName, tlVef.sName, 1) <= 0 Then
                                '    'If tmVpfInfo(ilTest).iNoVefLinks <= UBound(tmVpfInfo(ilTest).iVefLink) Then
                                '    '    tmVpfInfo(ilTest).iVefLink(tmVpfInfo(ilTest).iNoVefLinks) = tgVpf(ilLoop).iVefKCode
                                '    '    tmVpfInfo(ilTest).iNoVefLinks = tmVpfInfo(ilTest).iNoVefLinks + 1
                                '    'End If
                                '    tmVpfInfo(ilTest).iNoVefLinks = tmVpfInfo(ilTest).iNoVefLinks + 1
                                '    ilNextLk = tmVpfInfo(ilTest).iFirstLkVehInfo
                                '    ilIndex = UBound(tmLkVehInfo)
                                '    If ilNextLk = -1 Then
                                '        tmVpfInfo(ilTest).iFirstLkVehInfo = ilIndex
                                '    Else
                                '        Do
                                '            If tmLkVehInfo(ilNextLk).iNextLkVehInfo = -1 Then
                                '                tmLkVehInfo(ilNextLk).iNextLkVehInfo = ilIndex
                                '                Exit Do
                                '            End If
                                '            ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                                '        Loop
                                '    End If
                                '    tmLkVehInfo(ilIndex).iVefCode = tlVef.iCode
                                '    tmLkVehInfo(ilIndex).sVefName = tlVef.sName
                                '    tmLkVehInfo(ilIndex).iNextLkVehInfo = -1
                                '    ReDim Preserve tmLkVehInfo(0 To ilIndex + 1) As LKVEHINFO
                                'Else
                                '    'Switch places
                                '    'If tmVpfInfo(ilTest).iNoVefLinks <= UBound(tmVpfInfo(ilTest).iVefLink) Then
                                '    '    tmVpfInfo(ilTest).iVefLink(tmVpfInfo(ilTest).iNoVefLinks) = tmVpfInfo(ilTest).tVpf.iVefKCode
                                '    '    tmVpfInfo(ilTest).iNoVefLinks = tmVpfInfo(ilTest).iNoVefLinks + 1
                                '    'End If
                                '    tmVpfInfo(ilTest).iNoVefLinks = tmVpfInfo(ilTest).iNoVefLinks + 1
                                '    ilNextLk = tmVpfInfo(ilTest).iFirstLkVehInfo
                                '    ilIndex = UBound(tmLkVehInfo)
                                '    If ilNextLk = -1 Then
                                '        tmVpfInfo(ilTest).iFirstLkVehInfo = ilIndex
                                '    Else
                                '        Do
                                '            If tmLkVehInfo(ilNextLk).iNextLkVehInfo = -1 Then
                                '                tmLkVehInfo(ilNextLk).iNextLkVehInfo = ilIndex
                                '                Exit Do
                                '            End If
                                '            ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                                '        Loop
                                '    End If
                                '    tmLkVehInfo(ilIndex).iVefCode = tlSVef.iCode
                                '    tmLkVehInfo(ilIndex).sVefName = tlSVef.sName
                                '    tmLkVehInfo(ilIndex).iNextLkVehInfo = -1
                                '    ReDim Preserve tmLkVehInfo(0 To ilIndex + 1) As LKVEHINFO
                                '    tmVpfInfo(ilTest).tVpf = tgVpf(ilLoop)
                                'End If
                                Exit For
                            End If
                        Next ilTest
                    'End If
                    If Not ilFound Then
                        ilIndex = UBound(tmVpfInfo)
                        tmVpfInfo(ilIndex).tVpf = tgVpf(ilLoop)
                        'If StrComp(Trim$(tlVef.sName), "Prime", 1) = 0 Then
                        '    'Remove 24-Hour format from tgVpf.iGLink
                        '    For ilTest = LBound(tgVpf(ilLoop).iGLink) To UBound(tgVpf(ilLoop).iGLink) Step 1
                        '        If tgVpf(ilLoop).iGLink(ilTest) > 0 Then
                        '            tmVefSrchKey.iCode = tgVpf(ilTest).iVefKCode
                        '            ilRet = btrGetEqual(hmVef, tlSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        '            If StrComp(Trim$(tlSVef.sName), "24-HOUR FORMATS", 1) = 0 Then
                        '                tmVpfInfo(ilIndex).tVpf.iGLink(ilTest) = 0
                        '                For ilMove = ilTest To UBound(tgVpf(ilLoop).iGLink) - 1 Step 1
                        '                    tmVpfInfo(ilIndex).tVpf.iGLink(ilMove) = tmVpfInfo(ilIndex).tVpf.iGLink(ilMove + 1)
                        '                Next ilMove
                        '                Exit For
                        '                tmVpfInfo(ilIndex).tVpf.iGLink(UBound(tgVpf(ilLoop).iGLink)) = 0
                        '            End If
                        '        End If
                        '    Next ilTest
                        'End If
                        'tmVpfInfo(ilIndex).iNoVefLinks = LBound(tmVpfInfo(ilTest).iVefLink)
                        'tmVpfInfo(ilIndex).iVefLink(LBound(tmVpfInfo(ilTest).iVefLink)) = 0
                        'tmVpfInfo(ilIndex).sVefName(LBound(tmVpfInfo(ilTest).iVefLink)) = ""
                        tmVpfInfo(ilIndex).iNoVefLinks = 0
                        tmVpfInfo(ilIndex).iFirstLkVehInfo = -1
                        tmVpfInfo(ilIndex).iFirstSALink = -1
                        'tmVpfInfo(ilIndex).iVefLink(LBound(tmVpfInfo(ilTest).iVefLink)) = 0
                        'tmVpfInfo(ilIndex).sVefName(LBound(tmVpfInfo(ilTest).iVefLink)) = ""
                        ReDim Preserve tmVpfInfo(LBound(tgVpf) To ilIndex + 1) As VPFINFO
                    End If
                End If
            End If
        Next ilLoop
        'Order groups alphabetically
        For ilLoop = LBound(tmVpfInfo) To UBound(tmVpfInfo) - 1 Step 1
            'If tmVpfInfo(ilLoop).iNoVefLinks > LBound(tmVpfInfo(ilLoop).iVefLink) + 1 Then
            '    'Sort names- only the first two for now
            '    tmVefSrchKey.iCode = tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink))
            '    ilRet = btrGetEqual(hmVef, tlSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            '    tmVefSrchKey.iCode = tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink) + 1)
            '    ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            '    If StrComp(tlSVef.sName, tlVef.sName, 1) > 0 Then
            '        'Switch names
            '        tmVefSrchKey.iCode = tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink))
            '        tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink)) = tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink) + 1)
            '        tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink) + 1) = tmVefSrchKey.iCode
            '    End If
            'End If
            If tmVpfInfo(ilLoop).iNoVefLinks > 1 Then
                ilIndex = tmVpfInfo(ilLoop).iFirstLkVehInfo
                ilNextLk = tmLkVehInfo(ilIndex).iNextLkVehInfo
                Do While ilNextLk >= 0
                    Do
                        'tmVefSrchKey.iCode = tmLkVehInfo(ilIndex).iVefCode
                        'ilRet = btrGetEqual(hmVef, tlSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        ilRet = gBinarySearchVef(tmLkVehInfo(ilIndex).iVefCode)
                        If ilRet <> -1 Then
                            tlSVef = tgMVef(ilRet)
                        Else
                            tlSVef.sName = ""
                        End If
                        'tmVefSrchKey.iCode = tmLkVehInfo(ilNextLk).iVefCode
                        'ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        ilRet = gBinarySearchVef(tmLkVehInfo(ilNextLk).iVefCode)
                        If ilRet <> -1 Then
                            tlVef = tgMVef(ilRet)
                        Else
                            tlVef.sName = ""
                        End If
                        If StrComp(tlSVef.sName, tlVef.sName, 1) > 0 Then
                            'Switch names
                            tmVefSrchKey.iCode = tmLkVehInfo(ilIndex).iVefCode
                            slName = tmLkVehInfo(ilIndex).sVefName
                            tmLkVehInfo(ilIndex).iVefCode = tmLkVehInfo(ilNextLk).iVefCode
                            tmLkVehInfo(ilIndex).sVefName = tmLkVehInfo(ilNextLk).sVefName
                            tmLkVehInfo(ilNextLk).iVefCode = tmVefSrchKey.iCode
                            tmLkVehInfo(ilNextLk).sVefName = slName
                        End If
                        ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                    Loop While ilNextLk >= 0
                    ilIndex = tmLkVehInfo(ilIndex).iNextLkVehInfo
                    ilNextLk = tmLkVehInfo(ilIndex).iNextLkVehInfo
                Loop
            End If
        Next ilLoop
        ReDim tmVef(0 To 0) As VEF
        imVefRecLen = Len(tmVef(0))
        ilUpper = 0
        ilRet = btrGetFirst(hmVef, tmVef(ilUpper), imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            ilUse = False
            For ilTest = LBound(tmVpfInfo) To UBound(tmVpfInfo) - 1 Step 1
                If tmVpfInfo(ilTest).tVpf.iVefKCode = tmVef(ilUpper).iCode Then
                    ilUse = True
                    Exit For
                End If
                ilFound = False
                'For ilIndex = LBound(tmVpfInfo(ilTest).iVefLink) To tmVpfInfo(ilTest).iNoVefLinks - 1 Step 1
                '    If tmVpfInfo(ilTest).iVefLink(ilIndex) = tmVef(ilUpper).iCode Then
                '        tmVpfInfo(ilTest).sVefName(ilIndex) = Trim$(tmVef(ilUpper).sName)
                '        ilFound = True
                '        Exit For
                '    End If
                'Next ilIndex
                ilNextLk = tmVpfInfo(ilTest).iFirstLkVehInfo
                Do While ilNextLk >= 0
                    If tmLkVehInfo(ilNextLk).iVefCode = tmVef(ilUpper).iCode Then
                        ilFound = True
                        Exit For
                    End If
                    ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                Loop
                If ilFound Then
                    Exit For
                End If
            Next ilTest
            If ilUse Then
                ilUpper = ilUpper + 1
                ReDim Preserve tmVef(0 To ilUpper) As VEF
            End If
            ilRet = btrGetNext(hmVef, tmVef(ilUpper), imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        'Build Selling to Airing link table
        ReDim tmSALink(0 To 0) As LKVEHINFO
        For ilLoop = 0 To UBound(tmVef) - 1 Step 1
            If (tmVef(ilLoop).sType = "S") Then
                'slDate = Format$(Now, "m/d/yy")
                'gBuildLinkArray hmVlf, tmVef(ilLoop), slDate, imDVefCode()
                slDate = Format$(gNow(), "m/d/yy")
                For ilUse = 0 To 8 Step 1
                    gBuildLinkArray hmVlf, tmVef(ilLoop), slDate, imDVefCode()
                    If UBound(imDVefCode) > LBound(imDVefCode) Then
                        Exit For
                    End If
                    slDate = gIncOneWeek(slDate)
                Next ilUse
                'slDate = gObtainPrevMonday(slDate)
                'llDate = gDateValue(slDate)
                'ReDim tmTVlf(1 To 1) As VLF
                'gObtainVlf "S", hmVlf, tmVef(ilLoop).iCode, llDate, tmTVlf()
                'ReDim imDVefCode(0 To 0) As Integer
                'ilPrevAir = -1
                'For ilLoop1 = LBound(tmTVlf) To UBound(tmTVlf) - 1 Step 1
                '    If ilPrevAir <> tmTVlf(ilLoop1).iAirCode Then
                '        ilFound = False
                '        For ilTest = 0 To UBound(imDVefCode) - 1 Step 1
                '            If tmTVlf(ilLoop1).iAirCode = imDVefCode(ilTest) Then
                '                ilFound = True
                '                Exit For
                '            End If
                '        Next ilTest
                '        If Not ilFound Then
                '            If tgVpf(gVpfFind(ExpStnFd, tmTVlf(ilLoop1).iAirCode)).sGGroupNo <> "-" Then
                '                imDVefCode(UBound(imDVefCode)) = tmTVlf(ilLoop1).iAirCode
                '                ReDim Preserve imDVefCode(0 To UBound(imDVefCode) + 1) As Integer
                '            End If
                '        End If
                '        ilPrevAir = tmTVlf(ilLoop1).iAirCode
                '    End If
                'Next ilLoop1
                'llDate = llDate + 5
                'ReDim tmTVlf(1 To 1) As VLF
                'gObtainVlf "S", hmVlf, tmVef(ilLoop).iCode, llDate, tmTVlf()
                'ilPrevAir = -1
                'For ilLoop1 = LBound(tmTVlf) To UBound(tmTVlf) - 1 Step 1
                '    If ilPrevAir <> tmTVlf(ilLoop1).iAirCode Then
                '        ilFound = False
                '        For ilTest = 0 To UBound(imDVefCode) - 1 Step 1
                '            If tmTVlf(ilLoop1).iAirCode = imDVefCode(ilTest) Then
                '                ilFound = True
                '                Exit For
                '            End If
                '        Next ilTest
                '        If Not ilFound Then
                '            If tgVpf(gVpfFind(ExpStnFd, tmTVlf(ilLoop1).iAirCode)).sGGroupNo <> "-" Then
                '                imDVefCode(UBound(imDVefCode)) = tmTVlf(ilLoop1).iAirCode
                '                ReDim Preserve imDVefCode(0 To UBound(imDVefCode) + 1) As Integer
                '            End If
                '        End If
                '        ilPrevAir = tmTVlf(ilLoop1).iAirCode
                '    End If
                'Next ilLoop1
                'llDate = llDate + 1
                'ReDim tmTVlf(1 To 1) As VLF
                'gObtainVlf "S", hmVlf, tmVef(ilLoop).iCode, llDate, tmTVlf()
                'ilPrevAir = -1
                'For ilLoop1 = LBound(tmTVlf) To UBound(tmTVlf) - 1 Step 1
                '    If ilPrevAir <> tmTVlf(ilLoop1).iAirCode Then
                '        ilFound = False
                '        For ilTest = 0 To UBound(imDVefCode) - 1 Step 1
                '            If tmTVlf(ilLoop1).iAirCode = imDVefCode(ilTest) Then
                '                ilFound = True
                '                Exit For
                '            End If
                '        Next ilTest
                '        If Not ilFound Then
                '            If tgVpf(gVpfFind(ExpStnFd, tmTVlf(ilLoop1).iAirCode)).sGGroupNo <> "-" Then
                '                imDVefCode(UBound(imDVefCode)) = tmTVlf(ilLoop1).iAirCode
                '                ReDim Preserve imDVefCode(0 To UBound(imDVefCode) + 1) As Integer
                '            End If
                '        End If
                '        ilPrevAir = tmTVlf(ilLoop1).iAirCode
                '    End If
                'Next ilLoop1
                ilVpfIndex = mFindVpfIndex(tmVef(ilLoop).iCode)
                ilIndex = UBound(tmSALink)
                ilPrevIndex = -1
                For ilTest = 0 To UBound(imDVefCode) - 1 Step 1
                    'If tgVpf(gVpfFind(ExpStnFd, imDVefCode(ilTest))).sGGroupNo <> "-" Then
                    If (Trim$(tgVpf(gVpfFind(ExpStnFd, imDVefCode(ilTest))).sStnFdCode) <> "") And (Asc(tgVpf(gVpfFind(ExpStnFd, imDVefCode(ilTest))).sStnFdCode) <> 0) Then
                        If ilPrevIndex >= 0 Then
                            tmSALink(ilPrevIndex).iNextLkVehInfo = ilIndex
                        Else
                            tmVpfInfo(ilVpfIndex).iFirstSALink = ilIndex
                        End If
                        tmSALink(ilIndex).iVefCode = imDVefCode(ilTest)
                        tmSALink(ilIndex).sVefName = ""
                        tmSALink(ilIndex).iNextLkVehInfo = -1
                        ReDim Preserve tmSALink(0 To ilIndex + 1) As LKVEHINFO
                        ilPrevIndex = ilIndex
                        ilIndex = ilIndex + 1
                    End If
                Next ilTest
            End If
        Next ilLoop
        ReDim imDVefCode(0 To 0) As Integer
        'Generate names for file prefix (only for conventional and airing)
        For ilLoop = 0 To UBound(tmVef) - 1 Step 1
            If (tmVef(ilLoop).sType = "A") Or (tmVef(ilLoop).sType = "C") Then
                slName = Trim$(tmVef(ilLoop).sName)
                slPrefix = Left$(slName, 1)
                ilOk = False
                'Try first letter plus letter after blank or -
                For ilIndex = 2 To Len(slName) Step 1
                    slChar = Mid$(slName, ilIndex, 1)
                    If (slChar = " ") Or (slChar = "-") Then
                        If slPrefix <> "X" Then
                            ilOk = True
                            slSecond = Mid$(slName, ilIndex + 1, 1)
                            If Len(slSecond) > 0 Then
                                If ((Asc(slSecond) >= Asc("0")) And (Asc(slSecond) <= Asc("9"))) Or ((Asc(slSecond) >= Asc("A")) And (Asc(slSecond) <= Asc("Z"))) Then
                                    slPrefix = slPrefix & slSecond
                                    For ilTest = 0 To ilLoop - 1 Step 1
                                        If Trim$(tmVef(ilTest).sCodeStn) = slPrefix Then
                                            ilOk = False
                                            Exit For
                                        End If
                                    Next ilTest
                                Else
                                    ilOk = False
                                End If
                            Else
                                ilOk = False
                            End If
                        End If
                        If ilOk Then
                            tmVef(ilTest).sCodeStn = slPrefix
                            Exit For
                        Else
                            slPrefix = Left$(slName, 1)
                        End If
                    End If
                Next ilIndex
                If Not ilOk Then
                    'Try first two letters of the name
                    slPrefix = Left$(slName, 2)
                    If slPrefix <> "X" Then
                        ilOk = True
                        For ilTest = 0 To ilLoop - 1 Step 1
                            If Trim$(tmVef(ilTest).sCodeStn) = slPrefix Then
                                ilOk = False
                                Exit For
                            End If
                        Next ilTest
                    End If
                    If ilOk Then
                        tmVef(ilTest).sCodeStn = slPrefix
                    Else
                        For ilLetter = Asc("A") To Asc("Z") Step 1
                            slPrefix = Left$(slName, 1) & Chr(ilLetter)
                            If slPrefix <> "X" Then
                                ilOk = True
                                For ilTest = 0 To ilLoop - 1 Step 1
                                    If Trim$(tmVef(ilTest).sCodeStn) = slPrefix Then
                                        ilOk = False
                                        Exit For
                                    End If
                                Next ilTest
                            End If
                            If ilOk Then
                                tmVef(ilTest).sCodeStn = slPrefix
                                Exit For
                            End If
                        Next ilLetter
                    End If
                End If
            End If
        Next ilLoop
        ''Adjust links for selling only points to primary airing.
        'For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
        '    If (tmVef(ilLoop1).sType = "S") Then
        '        ilVpfIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
        '        If ilVpfIndex >= 0 Then
        '            For ilLoop2 = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tgVpf(ilVpfIndex).iGLink) Step 1
        '                ilFound = False
        '                If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop2) > 0 Then
        '                    For ilLoop = LBound(tmVpfInfo) To UBound(tmVpfInfo) - 1 Step 1
        '                        If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop2) = tmVpfInfo(ilLoop).tVpf.iVefKCode Then
        '                            ilFound = True
        '                        End If
        '                        If Not ilFound Then
        '                            For ilIndex = LBound(tmVpfInfo(ilLoop).iVefLink) To tmVpfInfo(ilLoop).iNoVefLinks - 1 Step 1
        '                                If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop2) = tmVpfInfo(ilLoop).iVefLink(ilIndex) Then
        '                                    tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop2) = tmVpfInfo(ilLoop).tVpf.iVefKCode
        '                                    ilFound = True
        '                                End If
        '                            Next ilIndex
        '                        End If
        '                        If ilFound Then
        '                            'Remove from links that are associated with this vehicle
        '                            For ilLoop3 = ilLoop2 + 1 To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
        '                                ilFound = False
        '                                If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop3) > 0 Then
        '                                    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop3) = tmVpfInfo(ilLoop).tVpf.iVefKCode Then
        '                                        tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop3) = 0
        '                                    End If
        '                                    For ilIndex = LBound(tmVpfInfo(ilLoop).iVefLink) To tmVpfInfo(ilLoop).iNoVefLinks - 1 Step 1
        '                                        If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop3) = tmVpfInfo(ilLoop).iVefLink(ilIndex) Then
        '                                            tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop3) = 0
        '                                        End If
        '                                    Next ilIndex
        '                                End If
        '                            Next ilLoop3
        '                            Exit For
        '                        End If
        '                    Next ilLoop
        '                End If
        '            Next ilLoop2
        '        End If
        '    End If
        'Next ilLoop1
    Else
        For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
            lbcVeh.Selected(ilLoop) = False
            tmcRot.Enabled = False
        Next ilLoop
    End If
    'Only show Airing and Conventional
    For ilLoop2 = 0 To UBound(tmVef) - 1 Step 1
        If (tmVef(ilLoop2).sType = "A") Or (tmVef(ilLoop2).sType = "C") Then
            ilVpfIndex = mFindVpfIndex(tmVef(ilLoop2).iCode)
            If ilVpfIndex >= 0 Then
                slName = Trim$(tmVef(ilLoop2).sName)
                'For ilLoop1 = LBound(tmVpfInfo(ilVpfIndex).sVefName) To tmVpfInfo(ilVpfIndex).iNoVefLinks - 1 Step 1
                '    slName = slName & " " & Trim$(tmVpfInfo(ilVpfIndex).sVefName(ilLoop1))
                'Next ilLoop1
                ilNextLk = tmVpfInfo(ilVpfIndex).iFirstLkVehInfo
                Do While ilNextLk >= 0
                    slName = slName & " " & Trim$(tmLkVehInfo(ilNextLk).sVefName)
                    ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                Loop
                lbcVehicleCode.AddItem slName & "|" & "0:0" & "\" & Trim$(str$(ilVpfIndex))
                lbcVehicleCode.ItemData(lbcVehicleCode.NewIndex) = tmVef(ilLoop2).iCode
            End If
        End If
    Next ilLoop2
    DoEvents
    For ilLoop1 = 0 To lbcVehicleCode.ListCount - 1 Step 1
        slNameCode = lbcVehicleCode.List(ilLoop1)
        ilRet = gParseItem(slNameCode, 1, "\", slName)    'Get application name
        lbcVehicle.AddItem slName
        lbcVehicle.ItemData(lbcVehicle.NewIndex) = lbcVehicleCode.ItemData(ilLoop1)
    Next ilLoop1
    pbclbcVehicle_Paint
    DoEvents
End Sub

Private Sub lbcVehicle_Scroll()
    pbclbcVehicle_Paint
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
                If imDateBox = 1 Then
                    edcStartDate.Text = Format$(llDate, "m/d/yy")
                    edcStartDate.SelStart = 0
                    edcStartDate.SelLength = Len(edcStartDate.Text)
                    imBypassFocus = True
                    edcStartDate.SetFocus
                    Exit Sub
                ElseIf imDateBox = 2 Then
                    edcEndDate.Text = Format$(llDate, "m/d/yy")
                    edcEndDate.SelStart = 0
                    edcEndDate.SelLength = Len(edcEndDate.Text)
                    imBypassFocus = True
                    edcEndDate.SetFocus
                    Exit Sub
                ElseIf imDateBox = 3 Then
                    edcTranDate.Text = Format$(llDate, "m/d/yy")
                    edcTranDate.SelStart = 0
                    edcTranDate.SelLength = Len(edcTranDate.Text)
                    imBypassFocus = True
                    edcTranDate.SetFocus
                    Exit Sub
                End If
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If imDateBox = 1 Then
        edcStartDate.SetFocus
    ElseIf imDateBox = 2 Then
        edcEndDate.SetFocus
    ElseIf imDateBox = 3 Then
        edcTranDate.SetFocus
    End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcDateTab_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub pbcInterface_Paint()
    pbcInterface.CurrentX = 0
    pbcInterface.CurrentY = 0
    pbcInterface.Print "Interface"
End Sub

Private Sub pbcLbcRot_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRotEnd As Integer
    Dim ilField As Integer
    Dim slFields(0 To 10) As String
    Dim llFgColor As Long
    Dim llWidth As Long
    Dim ilFieldIndex As Integer

    ilRotEnd = lbcRot.TopIndex + lbcRot.Height \ fgListHtArial825
    If ilRotEnd > lbcRot.ListCount Then
        ilRotEnd = lbcRot.ListCount
    End If
    If lbcRot.ListCount <= lbcRot.Height \ fgListHtArial825 Then
        llWidth = lbcRot.Width - 30
    Else
        llWidth = lbcRot.Width - igScrollBarWidth - 30
    End If
    pbcLbcRot.Width = llWidth
    pbcLbcRot.Cls
    llFgColor = pbcLbcRot.ForeColor
    For ilLoop = lbcRot.TopIndex To ilRotEnd - 1 Step 1
        pbcLbcRot.ForeColor = llFgColor
        If lbcRot.MultiSelect = 0 Then
            If lbcRot.ListIndex = ilLoop Then
                gPaintArea pbcLbcRot, CSng(0), CSng((ilLoop - lbcRot.TopIndex) * fgListHtArial825), CSng(pbcLbcRot.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcRot.ForeColor = vbWhite
            End If
        Else
            If lbcRot.Selected(ilLoop) Then
                gPaintArea pbcLbcRot, CSng(0), CSng((ilLoop - lbcRot.TopIndex) * fgListHtArial825), CSng(pbcLbcRot.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcRot.ForeColor = vbWhite
            End If
        End If
        slStr = lbcRot.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = 1 To 11 Step 1
            pbcLbcRot.CurrentX = imListFieldRot(ilField)
            pbcLbcRot.CurrentY = (ilLoop - lbcRot.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcRot, slStr, imListFieldRot(ilField + 1) - imListFieldRot(ilField)
            pbcLbcRot.Print slStr
        Next ilField
        pbcLbcRot.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub pbclbcVehicle_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilVehicleEnd As Integer
    Dim ilField As Integer
    Dim slFields(0 To 1) As String
    Dim llFgColor As Long
    Dim llWidth As Long
    Dim ilFieldIndex As Integer

    ilVehicleEnd = lbcVehicle.TopIndex + lbcVehicle.Height \ fgListHtArial825
    If ilVehicleEnd > lbcVehicle.ListCount Then
        ilVehicleEnd = lbcVehicle.ListCount
    End If
    If lbcVehicle.ListCount <= lbcVehicle.Height \ fgListHtArial825 Then
        llWidth = lbcVehicle.Width - 30
    Else
        llWidth = lbcVehicle.Width - igScrollBarWidth - 30
    End If
    pbcLbcVehicle.Width = llWidth
    pbcLbcVehicle.Cls
    llFgColor = pbcLbcVehicle.ForeColor
    For ilLoop = lbcVehicle.TopIndex To ilVehicleEnd - 1 Step 1
        pbcLbcVehicle.ForeColor = llFgColor
        If lbcVehicle.MultiSelect = 0 Then
            If lbcVehicle.ListIndex = ilLoop Then
                gPaintArea pbcLbcVehicle, CSng(0), CSng((ilLoop - lbcVehicle.TopIndex) * fgListHtArial825), CSng(pbcLbcVehicle.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcVehicle.ForeColor = vbWhite
            End If
        Else
            If lbcVehicle.Selected(ilLoop) Then
                gPaintArea pbcLbcVehicle, CSng(0), CSng((ilLoop - lbcVehicle.TopIndex) * fgListHtArial825), CSng(pbcLbcVehicle.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcVehicle.ForeColor = vbWhite
            End If
        End If
        slStr = lbcVehicle.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = 1 To UBound(slFields) Step 1
            pbcLbcVehicle.CurrentX = imListFieldVeh(ilField)
            pbcLbcVehicle.CurrentY = (ilLoop - lbcVehicle.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcVehicle, slStr, imListFieldVeh(ilField + 1) - imListFieldVeh(ilField)
            pbcLbcVehicle.Print slStr
        Next ilField
        pbcLbcVehicle.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub rbcGen_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcGen(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    Dim ilRet As Integer
    If Value Then
        If Index = 0 Then
            ckcAll.Visible = True
            lbcRot.Visible = True
            lbcVehicle.Visible = True
            vbcRot.Visible = True
            cmcGetRot.Visible = True
            plcDates.Height = 4080
            cmcSuppress.Enabled = True
            cmcReSend.Enabled = True
            lbcVeh.Visible = True
            lbcRegVeh.Visible = False
            lacTranDate.Visible = True
            edcTranDate.Visible = True
            cmcTranDate.Visible = True
            lacTranDate.Left = 2940
            edcTranDate.Left = 4740
            cmcTranDate.Left = 5685
            lbcFrom.Visible = True
            plcFrom.Visible = True
            cmcFrom.Visible = True
            lacRunLetter.Visible = True
            edcRunLetter.Visible = True
            edcStartDate.Text = gObtainNextMonday(smTodaysDate)
        ElseIf Index = 3 Then
            ckcAll.Visible = True
            lbcRot.Visible = True
            lbcVehicle.Visible = False
            vbcRot.Visible = True
            plcDates.Height = 4080
            cmcSuppress.Enabled = False
            cmcReSend.Enabled = False
            lbcVeh.Visible = True
            cmcGetRot.Visible = True
            lbcRegVeh.Visible = False
            lacTranDate.Visible = True
            edcTranDate.Visible = True
            cmcTranDate.Visible = True
            lacTranDate.Left = 120
            edcTranDate.Left = 1920
            cmcTranDate.Left = 2865
            lbcFrom.Visible = False
            plcFrom.Visible = False
            cmcFrom.Visible = False
            lacRunLetter.Visible = False
            edcRunLetter.Visible = False
            edcStartDate.Text = gObtainNextMonday(smTodaysDate)
        Else
            ckcAll.Visible = False
            lbcRot.Visible = False
            lbcVehicle.Visible = False
            vbcRot.Visible = False
            cmcGetRot.Visible = False
            plcDates.Height = 1740
            cmcSuppress.Enabled = False
            cmcReSend.Enabled = False
            lbcRegVeh.Visible = True
            lbcVeh.Visible = False
            lacTranDate.Visible = False
            edcTranDate.Visible = False
            cmcTranDate.Visible = False
            lbcFrom.Visible = True
            plcFrom.Visible = True
            cmcFrom.Visible = True
            lacRunLetter.Visible = True
            edcRunLetter.Visible = True
            'If lbcRegVeh.ListCount <= 0 Then
                Screen.MousePointer = vbHourglass
                lbcRegVeh.Clear
                smStationFile = Trim$(edcFrom.Text)
                ilRet = mGetStnInfo(False)
                Screen.MousePointer = vbDefault
            'End If
            edcStartDate.Text = gIncOneDay(smTodaysDate)
        End If
        edcEndDate.Text = ""
        ReDim tgSortCrf(0 To 0) As SORTCRF
        lbcRot.Clear
        pbcLbcRot_Paint
        ReDim tgDuplCrf(0 To 0) As DUPLCRF
        ReDim tgCombineCrf(0 To 0) As COMBINECRF
        ckcAll.Value = vbUnchecked
        For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
            lbcVeh.Selected(ilLoop) = False
            tmcRot.Enabled = False
        Next ilLoop
        For ilLoop = 0 To lbcRegVeh.ListCount - 1 Step 1
            lbcRegVeh.Selected(ilLoop) = False
        Next ilLoop
    End If
End Sub

Private Sub rbcInterface_Click(Index As Integer)
    If rbcInterface(Index).Value Then
        cmcSuppress.Enabled = True
        cmcReSend.Enabled = True
        plcSelect(Index).Visible = True
        plcFormat(Index).Visible = True
        ckcAll.Value = vbUnchecked
        ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
        ReDim tgSortCrf(0 To 0) As SORTCRF
        ReDim tmPSAPromoSortCrf(0 To 0) As SORTCRF
        lbcRot.Clear
        pbcLbcRot_Paint
        If Index = 0 Then
            plcSelect(1).Visible = False
            plcFormat(1).Visible = False
            plcCmmlLog.Visible = False
        Else
            plcSelect(0).Visible = False
            plcFormat(0).Visible = False
            plcCmmlLog.Visible = True
        End If
    End If
    mSetCommands
End Sub

Private Sub tmcRot_Timer()
    tmcRot.Enabled = False
    Screen.MousePointer = vbHourglass
    ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
    'mVehPop False
    imExptPrevWeek = True
    If (rbcInterface(0).Value) And (rbcGen(3).Value) Then
        imExptPrevWeek = False
    End If
    mRotPop
    Screen.MousePointer = vbDefault
End Sub
Private Sub vbcRot_Change()
    Dim ilStartIndex As Integer
    Dim ilEndIndex As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    If imIgnoreVbcChg Then
        Exit Sub
    End If
    imIgnoreVbcChg = True
    ilStartIndex = vbcRot.Value
    ilEndIndex = ilStartIndex + vbcRot.LargeChange
    If ilEndIndex > UBound(tgSortCrf) - 1 Then
        ilEndIndex = UBound(tgSortCrf) - 1
    End If
    ilValue = False
    If UBound(tgSortCrf) < vbcRot.LargeChange + 1 Then
        llRg = CLng(UBound(tgSortCrf) - 1) * &H10000 Or 0
    Else
        llRg = CLng(vbcRot.LargeChange) * &H10000 Or 0
    End If
    llRet = SendMessageByNum(lbcRot.HWnd, LB_SELITEMRANGE, ilValue, llRg)
    ilIndex = 0
    For ilLoop = ilStartIndex To ilEndIndex Step 1
        lbcRot.List(ilIndex) = tgSortCrf(ilLoop).sKey
        ilIndex = ilIndex + 1
    Next ilLoop
    ilIndex = 0
    For ilLoop = ilStartIndex To ilEndIndex Step 1
        lbcRot.Selected(ilIndex) = tgSortCrf(ilLoop).iSelected
        ilIndex = ilIndex + 1
    Next ilLoop
    pbcLbcRot_Paint
    imIgnoreVbcChg = False
End Sub
Private Sub vbcRot_Scroll()
    vbcRot_Change
End Sub
Private Sub plcSelect_Paint(Index As Integer)
    plcSelect(Index).CurrentX = 0
    plcSelect(Index).CurrentY = 0
    If Index = 0 Then
        plcSelect(Index).Print "Export"
    Else
        plcSelect(Index).Print "Envelope Carts"
    End If
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Station Feed"
End Sub
Private Sub plcFormat_Paint(Index As Integer)
    plcFormat(Index).CurrentX = 0
    plcFormat(Index).CurrentY = 0
    If Index = 0 Then
        plcFormat(Index).Print "Output Format"
    Else
        plcFormat(Index).Print "Output Rotation/Copy"
    End If
End Sub

Public Sub mPSAPromoProcess()
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim ilStn As Integer
    Dim ilCyf As Integer
    Dim ilRet As Integer

    'Save and restore at end of this procedure.
    ReDim tmSvSortCrf(0 To UBound(tgSortCrf)) As SORTCRF
    For ilLoop = 0 To UBound(tgSortCrf) Step 1
        tmSvSortCrf(ilLoop) = tgSortCrf(ilLoop)
    Next ilLoop
    ReDim tgSortCrf(0 To UBound(tmPSAPromoSortCrf)) As SORTCRF
    For ilLoop = 0 To UBound(tmPSAPromoSortCrf) Step 1
        tgSortCrf(ilLoop) = tmPSAPromoSortCrf(ilLoop)
        tgSortCrf(ilLoop).iSelected = True
    Next ilLoop
    For ilStn = 0 To UBound(tmStnInfo) - 1 Step 1
        If Trim$(tmStnInfo(ilStn).sFileName) <> "" Then
            mBuildExpTable tmStnInfo(ilStn)
            ilUpper = UBound(tmAddCyf)
            If ilUpper > 0 Then
                'ArraySortTyp fnAV(tgSort(),0), ilUpper, 0, LenB(tgSort(0)), 0, -9, 0
                ArraySortTyp fnAV(tmAddCyf(), 0), ilUpper, 0, LenB(tmAddCyf(0)), 0, LenB(tmAddCyf(0).sKey), 0
            End If
            'Code taken from mMergeXRefCyf because can't alter tmXRefCyf.  Instead created a new array for mCraetCartFile (tmWemCyf)
            'Test for duplicates
            For ilLoop = LBound(tmAddCyf) To UBound(tmAddCyf) - 1 Step 1
                ilFound = False
                For ilCyf = LBound(tmWemCyf) To UBound(tmWemCyf) - 1 Step 1
                    If (tmAddCyf(ilLoop).tCyf.lCifCode = tmWemCyf(ilCyf).tCyf.lCifCode) And (tmAddCyf(ilLoop).tCyf.iVefCode = tmWemCyf(ilCyf).tCyf.iVefCode) Then   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) Then
                        If (tmAddCyf(ilLoop).tCyf.sTimeZone = tmWemCyf(ilCyf).tCyf.sTimeZone) And (tmAddCyf(ilLoop).tCyf.lRafCode = tmWemCyf(ilCyf).tCyf.lRafCode) Then
                            ilFound = True
                            If tmAddCyf(ilLoop).lRotStartDate < tmWemCyf(ilCyf).lRotStartDate Then
                                tmWemCyf(ilCyf).lRotStartDate = tmAddCyf(ilLoop).lRotStartDate
                            End If
                            If tmAddCyf(ilLoop).lRotEndDate > tmWemCyf(ilCyf).lRotEndDate Then
                                tmWemCyf(ilCyf).lRotEndDate = tmAddCyf(ilLoop).lRotEndDate
                            End If
                            Exit For
                        End If
                    End If
                Next ilCyf
                If Not ilFound Then
                    tmWemCyf(UBound(tmWemCyf)) = tmAddCyf(ilLoop)
                    ReDim Preserve tmWemCyf(0 To UBound(tmWemCyf) + 1) As SENDCOPYINFO
                End If
            Next ilLoop
            For ilLoop = 0 To UBound(tmRotInfo) - 1 Step 1
                Do
                    tmCrfSrchKey.lCode = tmRotInfo(ilLoop).lCrfCode
                    ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        Print #hmMsg, "Get CRF to Update Feed Dates Failed" & Str$(ilRet) & " processing terminated"
'                        mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
'                        Screen.MousePointer = vbDefault
'                        ilRet = MsgBox("File in Use [Re-press Export], GetEqual Crf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
'                        Exit Function
'                    End If
                    If rbcInterface(0).Value Then
                        tmCrf.sAffFdStatus = "S" '"S"
                        tmCrf.sAffXMitChar = smRunLetter
                        'tmCrf.iFeedDate(0) = ilTranDate0
                        'tmCrf.iFeedDate(1) = ilTranDate1
                        'gPackDateLong lmInputStartDate, tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1)
                        mSetAffFdDate
                    Else
                        tmCrf.sKCFdStatus = "S" '"S"
                        tmCrf.sKCXMitChar = smRunLetter
                        'tmCrf.iFeedDate(0) = ilTranDate0
                        'tmCrf.iFeedDate(1) = ilTranDate1
                        'gPackDateLong lmInputStartDate, tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1)
                        mSetKCFdDate
                    End If
                    ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
            Next ilLoop
        End If
    Next ilStn
    'Restore tgSortCrf
    ReDim tgSortCrf(0 To UBound(tmSvSortCrf)) As SORTCRF
    For ilLoop = 0 To UBound(tmSvSortCrf) Step 1
        tgSortCrf(ilLoop) = tmSvSortCrf(ilLoop)
    Next ilLoop
End Sub

Private Sub mCheckForDuplOrCombines(ilCurrent As Integer, llRotStartDate As Long, llRotEndDate As Long, slShortTitle As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llSifCode                     ilVsf                                                   *
'******************************************************************************************

'       tmCrf(I)- Rotation to be checked
    Dim ilDuplIndex As Integer
    Dim ilCombIndex As Integer
    Dim ilUpper As Integer
    Dim ilRet As Integer
    Dim ilMatch As Integer
    Dim ilDatesMatch As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilDupl As Integer
    'Dim slShortTitle As String
    Dim slStr As String
    Dim ilDay As Integer
    Dim ilTest As Integer
    Dim tlCrf As CRF
    Dim tlAdf As ADF
    Dim tlChf As CHF
    ReDim llCifCode(0 To 0) As Long
    ReDim tgDuplCrf(0 To 0) As DUPLCRF
    ReDim tgCombineCrf(0 To 0) As COMBINECRF

    gUnpackDateLong tmCrf.iStartDate(0), tmCrf.iStartDate(1), llRotStartDate
    gUnpackDateLong tmCrf.iEndDate(0), tmCrf.iEndDate(1), llRotEndDate
    ilDuplIndex = -1
    ilCombIndex = -1
    'Get instructions
    ilUpper = 0
    tmCnfSrchKey.lCrfCode = tmCrf.lCode
    tmCnfSrchKey.iInstrNo = 0
    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
        llCifCode(ilUpper) = tmCnf.lCifCode
        ilUpper = ilUpper + 1
        ReDim Preserve llCifCode(0 To ilUpper)
        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If tmAdf.iCode <> tmChf.iAdfCode Then
        tmAdfSrchKey.iCode = tmChf.iAdfCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    End If
'    llSifCode = 0
'    If tmChf.lVefCode < 0 Then
'        tmVsfSrchKey.lCode = -tmChf.lVefCode
'        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        Do While ilRet = BTRV_ERR_NONE
'            For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
'                If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
'                    If tmVsf.lFSComm(ilVsf) > 0 Then
'                        llSifCode = tmVsf.lFSComm(ilVsf)
'                    End If
'                    Exit For
'                End If
'            Next ilVsf
'            If llSifCode <> 0 Then
'                Exit Do
'            End If
'            If tmVsf.lLkVsfCode <= 0 Then
'                Exit Do
'            End If
'            tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
'            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'    End If
'    slShortTitle = gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf)
    For ilDupl = ilCurrent + 1 To UBound(tmRotInfo) - 1 Step 1
        If (tmRotInfo(ilDupl).iStatus = 1) And (tmRotInfo(ilDupl).iVefCode = tmRotInfo(ilCurrent).iVefCode) Then
            tmCrfSrchKey.lCode = tmRotInfo(ilDupl).lCrfCode
            ilRet = btrGetEqual(hmCrf, tlCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If tlChf.lCode <> tlCrf.lChfCode Then
                    If tmChf.lCode <> tlCrf.lChfCode Then
                        tmChfSrchKey.lCode = tlCrf.lChfCode
                        ilRet = btrGetEqual(hmCHF, tlChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Else
                        tlChf = tmChf
                        ilRet = BTRV_ERR_NONE
                    End If
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                If ilRet = BTRV_ERR_NONE Then
                    If tlAdf.iCode <> tmChf.iAdfCode Then
                        tmAdfSrchKey.iCode = tlChf.iAdfCode
                        ilRet = btrGetEqual(hmAdf, tlAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    End If
'                    llSifCode = 0
'                    If tlChf.lVefCode < 0 Then
'                        tmVsfSrchKey.lCode = -tlChf.lVefCode
'                        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                        Do While ilRet = BTRV_ERR_NONE
'                            For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
'                                If tmVsf.iFSCode(ilVsf) = tlCrf.iVefCode Then
'                                    If tmVsf.lFSComm(ilVsf) > 0 Then
'                                        llSifCode = tmVsf.lFSComm(ilVsf)
'                                    End If
'                                    Exit For
'                                End If
'                            Next ilVsf
'                            If llSifCode <> 0 Then
'                                Exit Do
'                            End If
'                            If tmVsf.lLkVsfCode <= 0 Then
'                                Exit Do
'                            End If
'                            tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
'                            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                        Loop
'                    End If
'                    slStr = gGetProdOrShtTitle(hmSif, llSifCode, tlChf, tlAdf)
                    ilMatch = True
                    ilDatesMatch = True
'                    If StrComp(Trim$(slStr), Trim$(slShortTitle), 1) <> 0 Then
'                        ilMatch = False
'                    End If
                    If tmCrf.iAdfCode <> tlCrf.iAdfCode Then
                        ilMatch = False
                    End If
                    If tmCrf.sRotType <> tlCrf.sRotType Then
                        ilMatch = False
                    End If
                    'Determine if Dates, times, days,... are the same
'                    If (tmCrf.iStartDate(0) <> tlCrf.iStartDate(0)) Or (tmCrf.iStartDate(1) <> tlCrf.iStartDate(1)) Then
'                        ilMatch = False
'                    End If
'                    If (tmCrf.iEndDate(0) <> tlCrf.iEndDate(0)) Or (tmCrf.iEndDate(1) <> tlCrf.iEndDate(1)) Then
'                        ilMatch = False
'                    End If
                    gUnpackDateLong tlCrf.iStartDate(0), tlCrf.iStartDate(1), llStartDate
                    gUnpackDateLong tlCrf.iEndDate(0), tlCrf.iEndDate(1), llEndDate
                    If (llRotStartDate <> llStartDate) Or (llRotEndDate <> llEndDate) Then
                        If (llEndDate < llRotStartDate) Then
                            ilMatch = False
                        ElseIf llStartDate > llRotEndDate Then
                            ilMatch = False
                        Else
                            ilDatesMatch = False
                        End If
                    End If
                    If (tmCrf.iStartTime(0) <> tlCrf.iStartTime(0)) Or (tmCrf.iStartTime(1) <> tlCrf.iStartTime(1)) Then
                        ilMatch = False
                    End If
                    If (tmCrf.iEndTime(0) <> tlCrf.iEndTime(0)) Or (tmCrf.iEndTime(1) <> tlCrf.iEndTime(1)) Then
                        ilMatch = False
                    End If
                    For ilDay = 0 To 6 Step 1
                        If (tmCrf.sDay(ilDay) <> tlCrf.sDay(ilDay)) Then
                            ilMatch = False
                            Exit For
                        End If
                    Next ilDay
                    If tmCrf.sZone <> tlCrf.sZone Then
                        ilMatch = False
                    Else
                        If Trim$(tmCrf.sZone) = "R" Then
                            If tmCrf.lRafCode <> tlCrf.lRafCode Then
                                ilMatch = False
                            End If
                        End If
                    End If
                    If tmCrf.iLen <> tlCrf.iLen Then
                        ilMatch = False
                    End If
                    If tmCrf.sInOut <> tlCrf.sInOut Then
                        ilMatch = False
                    End If
                    If tmCrf.ianfCode <> tlCrf.ianfCode Then
                        ilMatch = False
                    End If
                    If ilMatch Then
                        slStr = mGetShortTitle(tlChf, tlAdf, tlCrf.iVefCode)
                        If StrComp(Trim$(slStr), Trim$(slShortTitle), 1) <> 0 Then
                            ilMatch = False
                        End If
                    End If
                    If ilMatch Then
                        ilTest = 0
                        tmCnfSrchKey.lCrfCode = tlCrf.lCode
                        tmCnfSrchKey.iInstrNo = 0
                        ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tlCrf.lCode)
                            If ilTest >= UBound(llCifCode) Then
                                ilMatch = False
                                Exit Do
                            End If
                            If tmCnf.lCifCode <> llCifCode(ilTest) Then
                                ilMatch = False
                                Exit Do
                            End If
                            ilTest = ilTest + 1
                            ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If ilMatch Then
                            If Not ilDatesMatch Then
                                If llStartDate < llRotStartDate Then
                                    llRotStartDate = llStartDate
                                End If
                                If llEndDate > llRotEndDate Then
                                    llRotEndDate = llEndDate
                                End If
                            End If
                        End If
                        tmRotInfo(ilDupl).iStatus = 2
                        'Set combine flag
                        tgCombineCrf(UBound(tgCombineCrf)).lCntrNo = 0
                        tgCombineCrf(UBound(tgCombineCrf)).sVehName = ""
                        tgCombineCrf(UBound(tgCombineCrf)).tCrf = tlCrf
                        tgCombineCrf(UBound(tgCombineCrf)).lCrfRecPos = 0
                        tgCombineCrf(UBound(tgCombineCrf)).iCombineIndex = -1
                        tgCombineCrf(UBound(tgCombineCrf)).iVpfIndex = 0
                        ReDim Preserve tgCombineCrf(0 To UBound(tgCombineCrf) + 1) As COMBINECRF
                    End If
                End If
            End If
        End If
    Next ilDupl
End Sub

Private Function mGetShortTitle(tlChf As CHF, tlAdf As ADF, ilVefCode As Integer) As String
    Dim llSifCode As Long
    Dim ilRet As Integer
    Dim slShortTitle As String
    Dim ilVsf As Integer

    llSifCode = 0
    If tlChf.lVefCode < 0 Then
        tmVsfSrchKey.lCode = -tlChf.lVefCode
        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While ilRet = BTRV_ERR_NONE
            For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilVsf) = ilVefCode Then
                    If tmVsf.lFSComm(ilVsf) > 0 Then
                        llSifCode = tmVsf.lFSComm(ilVsf)
                    End If
                    Exit For
                End If
            Next ilVsf
            If llSifCode <> 0 Then
                Exit Do
            End If
            If tmVsf.lLkVsfCode <= 0 Then
                Exit Do
            End If
            tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    slShortTitle = gGetProdOrShtTitle(hmSif, llSifCode, tlChf, tlAdf)
    mGetShortTitle = slShortTitle
End Function

Private Function mCheckVehicles() As Integer
    Dim ilVeh As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slInDate As String
    Dim ilVef As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilStn As Integer
    Dim ilFound As Integer
    Dim tlVef As VEF
    ReDim ilSelVefCode(0 To 0) As Integer
    ReDim ilVefCode(0 To 0) As Integer


    mCheckVehicles = True
    slInDate = Format$(lmInputStartDate, "m/d/yy")
    For ilVeh = 0 To lbcVeh.ListCount - 1 Step 1
        If lbcVeh.Selected(ilVeh) Then
            slNameCode = tmVehCode(ilVeh).sKey    'Selling and conventional vehicles 'lbcVehCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
            ilCode = Val(slCode)
            ilVef = gBinarySearchVef(ilCode)
            If ilVef <> -1 Then
                tlVef = tgMVef(ilVef)
                If tlVef.sType = "S" Then
                    gBuildLinkArray hmVlf, tlVef, slInDate, ilVefCode()
                Else
                    ReDim ilVefCode(0 To 1) As Integer
                    ilVefCode(0) = ilCode
                End If
                For ilLoop = 0 To UBound(ilVefCode) - 1 Step 1
                    ilCode = ilVefCode(ilLoop)
                    ilFound = False
                    For ilTest = 0 To UBound(ilSelVefCode) - 1 Step 1
                        If ilCode = ilSelVefCode(ilTest) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilTest
                    If Not ilFound Then
                        ilSelVefCode(UBound(ilSelVefCode)) = ilCode
                        ReDim Preserve ilSelVefCode(0 To UBound(ilSelVefCode) + 1) As Integer
                    End If
                Next ilLoop
            End If
        End If
    Next ilVeh
    ReDim ilVefCode(0 To 0) As Integer
    For ilStn = 0 To UBound(tmStnInfo) - 1 Step 1
        ilFound = False
        For ilVeh = 0 To UBound(ilSelVefCode) - 1 Step 1
            If tmStnInfo(ilStn).iAirVeh = ilSelVefCode(ilVeh) Then
                ilFound = True
                Exit For
            End If
        Next ilVeh
        If Not ilFound Then
            For ilVeh = 0 To UBound(ilVefCode) - 1 Step 1
                If tmStnInfo(ilStn).iAirVeh = ilVefCode(ilVeh) Then
                    ilFound = True
                    Exit For
                End If
            Next ilVeh
            If Not ilFound Then
                ilVef = gBinarySearchVef(tmStnInfo(ilStn).iAirVeh)
                If ilVef <> -1 Then
                    'Print #hmMsg, "Vehicle in Import File but not Selected " & Trim$(tgMVef(ilVef).sName)
                    gAutomationAlertAndLogHandler "Vehicle in Import File but not Selected " & Trim$(tgMVef(ilVef).sName)
                End If
                ilVefCode(UBound(ilVefCode)) = tmStnInfo(ilStn).iAirVeh
                ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                mCheckVehicles = False
            End If
        End If
    Next ilStn
    For ilVeh = 0 To UBound(ilSelVefCode) - 1 Step 1
        ilFound = False
        For ilStn = 0 To UBound(tmStnInfo) - 1 Step 1
            If tmStnInfo(ilStn).iAirVeh = ilSelVefCode(ilVeh) Then
                ilFound = True
                Exit For
            End If
        Next ilStn
        If Not ilFound Then
            ilVef = gBinarySearchVef(ilSelVefCode(ilVeh))
            If ilVef <> -1 Then
                If tgMVef(ilVef).sType = "A" Then
                    'Print #hmMsg, "Selling Vehicle Selected but Airing Vehicle " & Trim$(tgMVef(ilVef).sName) & " Not in the Import File"
                    gAutomationAlertAndLogHandler "Selling Vehicle Selected but Airing Vehicle " & Trim$(tgMVef(ilVef).sName) & " Not in the Import File"
                Else
                    'Print #hmMsg, "Conventional Vehicle " & Trim$(tgMVef(ilVef).sName) & " Selected but not in the Import File"
                    gAutomationAlertAndLogHandler "Conventional Vehicle " & Trim$(tgMVef(ilVef).sName) & " Selected but not in the Import File"
                End If
            End If
            mCheckVehicles = False
        End If
    Next ilVeh
End Function

Private Sub mGetRotInfo(slISCI As String, slRotStartDate As String, slRotEndDate As String, slRotComment As String, slTimeRestrictions As String, slDayRestrictions As String, llCsfCode As Long)
    Dim ilRet As Integer
    Dim slTime As String
    Dim ilDayReq As Integer
    Dim ilDay As Integer
    Dim llRotStartDate As Long
    Dim llRotEndDate As Long
    Dim llDate As Long
    '5/31/06:
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilDupl As Integer
    Dim ilNoteNo As Integer
    '5/31/06
    Dim slStr As String
    ReDim ilDayOn(0 To 6) As Integer
    Dim tlCrf As CRF

    slRotStartDate = ""
    slRotEndDate = ""
    slRotComment = ""
    slTimeRestrictions = ""
    slDayRestrictions = ""
    llCsfCode = 0
    tmCrfSrchKey1.sRotType = "A"
    tmCrfSrchKey1.iEtfCode = 0
    tmCrfSrchKey1.iEnfCode = 0
    tmCrfSrchKey1.iAdfCode = tmSdf.iAdfCode
    tmCrfSrchKey1.lChfCode = tmSdf.lChfCode
    tmCrfSrchKey1.lFsfCode = 0
    tmCrfSrchKey1.iVefCode = 0
    tmCrfSrchKey1.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hmCrf, tlCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (tlCrf.sRotType = "A") And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tmSdf.iAdfCode) And (tlCrf.lChfCode = tmSdf.lChfCode)
        If (tlCrf.iRotNo = tmSdf.iRotNo) And (tlCrf.sState <> "D") Then
            gUnpackDate tlCrf.iStartDate(0), tlCrf.iStartDate(1), slRotStartDate
            gUnpackDate tlCrf.iEndDate(0), tlCrf.iEndDate(1), slRotEndDate
            If tlCrf.lCsfCode <> 0 Then
                tmCsfSrchKey.lCode = tlCrf.lCsfCode
                tmCsf.sComment = ""
                imCsfRecLen = Len(tmCsf) '5011
                ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'If tmCsf.iStrLen > 0 Then
                    slStr = gStripChr0(tmCsf.sComment)
                    If slStr <> "" Then
                        slRotComment = slStr    'Trim$(Left$(tmCsf.sComment, tmCsf.iStrLen))
                        llCsfCode = tlCrf.lCsfCode
                    End If
                End If
            End If
            'Special Instructions
            If (tlCrf.iStartTime(0) <> 0) Or (tlCrf.iStartTime(1) <> 0) Or (tlCrf.iEndTime(0) <> 0) Or (tlCrf.iEndTime(1) <> 0) Then
                slTimeRestrictions = slTimeRestrictions & "Air copy between "
                gUnpackTime tlCrf.iStartTime(0), tlCrf.iStartTime(1), "A", "1", slTime
                slTime = UCase(slTime)
                If slTime = "12AM" Then
                    slTime = "12M"
                ElseIf slTime = "12PM" Then
                    slTime = "12N"
                End If
                slTimeRestrictions = slTimeRestrictions & slTime & " and "
                gUnpackTime tlCrf.iEndTime(0), tlCrf.iEndTime(1), "A", "1", slTime
                slTime = UCase(slTime)
                If slTime = "12AM" Then
                    slTime = "12M"
                ElseIf slTime = "12PM" Then
                    slTime = "12N"
                End If
                slTimeRestrictions = slTimeRestrictions & slTime
            End If
            llRotStartDate = gDateValue(slRotStartDate)
            llRotEndDate = gDateValue(slRotEndDate)
            ilDayReq = False
            For ilDay = 0 To 6 Step 1
                ilDayOn(ilDay) = False
            Next ilDay
            If llRotEndDate - llRotStartDate >= 6 Then
                For ilDay = 0 To 6 Step 1
                    If tlCrf.sDay(ilDay) <> "Y" Then
                        ilDayReq = True
                        ilDayOn(ilDay) = False
                    Else
                        ilDayOn(ilDay) = True
                    End If
                Next ilDay
            Else
                For llDate = llRotStartDate To llRotEndDate Step 1
                    ilDay = gWeekDayLong(llDate)
                    If tlCrf.sDay(ilDay) <> "Y" Then
                        'ilDayReq = True
                        ilDayOn(ilDay) = False
                    Else
                        ilDayOn(ilDay) = True
                    End If
                Next llDate
                ilDayReq = True
            End If
            If ilDayReq Then
                slStr = ""
                If (ilDayOn(0) = True) And (ilDayOn(1) = True) And (ilDayOn(2) = True) And (ilDayOn(3) = True) And (ilDayOn(4) = True) And (ilDayOn(5) = False) And (ilDayOn(6) = False) Then
                    slStr = "Mon thru Fri"
                ElseIf (ilDayOn(0) = False) And (ilDayOn(1) = False) And (ilDayOn(2) = False) And (ilDayOn(3) = False) And (ilDayOn(4) = False) And (ilDayOn(5) = True) And (ilDayOn(6) = True) Then
                    slStr = "Sat and Sun"
                Else
                    For ilDay = 0 To 6 Step 1
                        If ilDayOn(ilDay) = True Then
                            Select Case ilDay
                                Case 0
                                    slStr = slStr & " Mon"
                                Case 1
                                    slStr = slStr & " Tue"
                                Case 2
                                    slStr = slStr & " Wed"
                                Case 3
                                    slStr = slStr & " Thu"
                                Case 4
                                    slStr = slStr & " Fri"
                                Case 5
                                    slStr = slStr & " Sat"
                                Case 6
                                    slStr = slStr & " Sun"
                            End Select
                        End If
                    Next ilDay
                End If
                slDayRestrictions = "Air copy on " & Trim$(slStr)
            End If
            '5/31/06:  Show comments on first occurrance only
            If (Trim$(slRotComment) <> "") Or (Trim$(slTimeRestrictions) <> "") Or (Trim$(slDayRestrictions) <> "") Then
                ilFound = False
                For ilLoop = 0 To UBound(smRotComment) - 1 Step 1
                    If StrComp(Trim$(slRotComment), Trim$(smRotComment(ilLoop)), vbTextCompare) = 0 Then
                        If StrComp(Trim$(slTimeRestrictions), Trim$(smTimeRestrictions(ilLoop)), vbTextCompare) = 0 Then
                            If StrComp(Trim$(slDayRestrictions), Trim$(smDayRestrictions(ilLoop)), vbTextCompare) = 0 Then
                                For ilDupl = 0 To UBound(tmDuplComment) - 1 Step 1
                                    If (tmDuplComment(ilDupl).iAdfCode = tlCrf.iAdfCode) Then
                                        If StrComp(Trim$(tmDuplComment(ilDupl).sISCI), Trim$(slISCI), vbTextCompare) = 0 Then
                                            ilFound = True
                                            ilNoteNo = tmDuplComment(ilDupl).iNoteNo
                                        ElseIf tmDuplComment(ilDupl).lChfCode = tlCrf.lChfCode Then
                                            ilFound = True
                                            ilNoteNo = tmDuplComment(ilDupl).iNoteNo
                                            tmDuplComment(UBound(tmDuplComment)).iNoteNo = tmDuplComment(ilDupl).iNoteNo
                                            tmDuplComment(UBound(tmDuplComment)).iAdfCode = tmDuplComment(ilDupl).iAdfCode
                                            tmDuplComment(UBound(tmDuplComment)).lChfCode = tmDuplComment(ilDupl).lChfCode
                                            tmDuplComment(UBound(tmDuplComment)).sISCI = Trim$(slISCI)
                                            ReDim Preserve tmDuplComment(0 To UBound(tmDuplComment) + 1) As DUPLCOMMENT
                                        End If
                                    End If
                                Next ilDupl
                                If ilFound Then
                                    slRotComment = " "
                                    Do While Len(slRotComment) < 10
                                        slRotComment = slRotComment & " "
                                    Loop
                                    slRotComment = slRotComment & "(" & "See Note " & Trim$(str$(ilNoteNo)) & ")"
                                    slTimeRestrictions = ""
                                    slDayRestrictions = ""
                                    llCsfCode = 0
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next ilLoop
                If Not ilFound Then
                    smRotComment(UBound(smRotComment)) = slRotComment
                    ReDim Preserve smRotComment(0 To UBound(smRotComment) + 1) As String
                    If (Trim$(slRotComment) <> "") Then
                        slStr = "Note " & Trim$(str$(UBound(smRotComment)))
                        Do While Len(slStr) < 10
                            slStr = slStr & " "
                        Loop
                        slRotComment = slStr & slRotComment
                    End If
                    smTimeRestrictions(UBound(smTimeRestrictions)) = slTimeRestrictions
                    ReDim Preserve smTimeRestrictions(0 To UBound(smTimeRestrictions) + 1) As String
                    If (Trim$(slTimeRestrictions) <> "") Then
                        If (Trim$(slRotComment) <> "") Then
                            slStr = String(10, " ")
                        Else
                            slStr = "Note " & Trim$(str$(UBound(smRotComment)))
                            Do While Len(slStr) < 10
                                slStr = slStr & " "
                            Loop
                        End If
                        slTimeRestrictions = slStr & slTimeRestrictions
                    End If
                    smDayRestrictions(UBound(smDayRestrictions)) = slDayRestrictions
                    ReDim Preserve smDayRestrictions(0 To UBound(smDayRestrictions) + 1) As String
                    If (Trim$(slDayRestrictions) <> "") Then
                        If (Trim$(slRotComment) <> "") Or (Trim$(slTimeRestrictions) <> "") Then
                            slStr = String(10, " ")
                        Else
                            slStr = "Note " & Trim$(str$(UBound(smRotComment)))
                            Do While Len(slStr) < 10
                                slStr = slStr & " "
                            Loop
                        End If
                        slDayRestrictions = slStr & slDayRestrictions
                    End If
                    tmDuplComment(UBound(tmDuplComment)).iNoteNo = Trim$(str$(UBound(smRotComment)))
                    tmDuplComment(UBound(tmDuplComment)).iAdfCode = tlCrf.iAdfCode
                    tmDuplComment(UBound(tmDuplComment)).lChfCode = tlCrf.lChfCode
                    tmDuplComment(UBound(tmDuplComment)).sISCI = Trim$(slISCI)
                    ReDim Preserve tmDuplComment(0 To UBound(tmDuplComment) + 1) As DUPLCOMMENT
                End If
            End If
            '5/31/06
            Exit Sub
        End If
        ilRet = btrGetNext(hmCrf, tlCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub

Private Function mWriteRestrictions(hlSch As Integer, slMsg As String, llCsfCode As Long) As Integer
    Dim ilRet As Integer
    Dim ilPos As Integer

    mWriteRestrictions = True
    If Trim$(slMsg) = "" Then
        Exit Function
    End If
    If (ckcCmmlLog(1).Value = vbChecked) Then
        'Print #hlSch, ",," & Trim$(slMsg)
        '5/31/06:  Show comments on first occurrance only
        ilPos = InStr(1, slMsg, "Note ", vbTextCompare)
        If ilPos = 1 Then
            ilPos = InStr(ilPos + 5, slMsg, " ", vbTextCompare)
            If ilPos > 0 Then
                Print #hlSch, "," & Trim$(Left$(slMsg, ilPos - 1)) & ":" & "," & Trim$(Mid$(slMsg, ilPos + 1))
            Else
                Print #hlSch, ",," & Trim$(slMsg)
            End If
        Else
            Print #hlSch, ",," & Trim$(slMsg)
        End If
        '5/31/06
    End If
    If ckcCmmlLog(0).Value = vbChecked Then
        gPackDate smPDFDate, tmTxr.iGenDate(0), tmTxr.iGenDate(1)
        tmTxr.lGenTime = gTimeToLong(smPDFTime, False)
        lmPDFSeqNo = lmPDFSeqNo + 1
        tmTxr.lSeqNo = lmPDFSeqNo
        tmTxr.iType = 3
        If llCsfCode > 0 Then
            tmTxr.sText = Left$(slMsg, 10)
        Else
            tmTxr.sText = slMsg
        End If
        tmTxr.lCsfCode = llCsfCode
        ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            'Print #hmMsg, "Insert TXR Failed" & str$(ilRet) & " processing terminated"
            gAutomationAlertAndLogHandler "Insert TXR Failed" & str$(ilRet) & " processing terminated"
            If ckcCmmlLog(1).Value = vbChecked Then
                Close #hlSch
            End If
            mClearTxr
            btrDestroy hmTxr
            mWriteRestrictions = False
            Exit Function
        End If
    End If

End Function

Private Function mExportXFerHeader(ilPrtFirstXRef As Integer, ilLineNo As Integer, ilPageNo As Integer, slRecord As String) As Integer
    If ilPrtFirstXRef Then
        ilPrtFirstXRef = False
        If ilPageNo = 0 Then
            slRecord = ""
            If Not mExportLine(slRecord, ilLineNo, -1) Then
                mExportXFerHeader = False
                Exit Function
            End If
        Else
            slRecord = Chr(12)  'Form Feed
            If Not mExportLine(slRecord, ilLineNo, -1) Then
                mExportXFerHeader = False
                Exit Function
            End If
        End If
        ilPageNo = ilPageNo + 1
        ilLineNo = 0
        slRecord = " "
        Do While Len(slRecord) < 35
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & Trim$(tgSpf.sGClient)
        If Not mExportLine(slRecord, ilLineNo, -1) Then
            mExportXFerHeader = False
            Exit Function
        End If
        slRecord = " "
        Do While Len(slRecord) < 35
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & "Cross Reference"
        If Not mExportLine(slRecord, ilLineNo, -1) Then
            mExportXFerHeader = False
            Exit Function
        End If
        slRecord = " "
        Do While Len(slRecord) < 35
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & smTranDate & "  "
        'slRecord = slRecord & "Page:"
        'slStr = Trim$(Str$(ilPageNo))
        'Do While Len(slStr) < 5
        '    slStr = " " & slStr
        'Loop
        'slRecord = slRecord & slStr
        If Not mExportLine(slRecord, ilLineNo, -1) Then
            mExportXFerHeader = False
            Exit Function
        End If
        slRecord = ""
        If Not mExportLine(slRecord, ilLineNo, -1) Then
            mExportXFerHeader = False
            Exit Function
        End If
        If Not mExportLine(slRecord, ilLineNo, -1) Then
            mExportXFerHeader = False
            Exit Function
        End If
        slRecord = "Short Title"
        Do While Len(slRecord) < 20
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & " Cart"
        Do While Len(slRecord) < 31
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & " ISCI"
        Do While Len(slRecord) < 52
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & " Creative Title"
        Do While Len(slRecord) < 83
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & " Vehicle"
        If Not mExportLine(slRecord, ilLineNo, -1) Then
            mExportXFerHeader = False
            Exit Function
        End If
    End If
    mExportXFerHeader = True
End Function

Private Sub mExportCopyHeader()
    Exit Sub
End Sub

Private Function mExportRotHeader(ilPrtFirstRot As Integer, ilNewHdRot As Integer, ilLineNo As Integer, ilPageNo As Integer, slVehName As String, tlStnInfo As STNINFO, slBlank As String, slRecord As String) As Integer
    If ilPrtFirstRot Then
        ilPrtFirstRot = False
        If ilPageNo > 0 Then
            If ((rbcInterface(0).Value) And (rbcFormat(1).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(3).Value)) Then
                slRecord = Chr(12)  'Form Feed
                If Not mExportLine(slRecord, ilLineNo, 5) Then
                    mExportRotHeader = False
                    Exit Function
                End If
            End If
        End If
        ilPageNo = ilPageNo + 1
        ilLineNo = 0
        'slRecord = "-"
        'Do While Len(slRecord) < 78
        '    slRecord = slRecord & "-"
        'Loop
        'If Not mExportLine(slRecord, ilLineNo) Then
        '    Exit Function
        'End If
        'If Not mExportLine(slRecord, ilLineNo) Then
        '    Exit Function
        'End If
        ''slRecord = " "
        ''Do While Len(slRecord) < 68
        ''    slRecord = slRecord & " "
        ''Loop
        ''slRecord = slRecord & "Page:"
        ''slStr = Trim$(Str$(ilPageNo))
        ''Do While Len(slStr) < 5
        ''    slStr = " " & slStr
        ''Loop
        ''slRecord = slRecord & slStr
        ''If Not mExportLine(slRecord, ilLineNo) Then
        ''    Exit Function
        ''End If
        'slRecord = ""
        'If Not mExportLine(slRecord, ilLineNo) Then
        '    Exit Function
        'End If
        'If Not mExportLine(slRecord, ilLineNo) Then
        '    Exit Function
        'End If
        'slRecord = Trim$(tgSpf.sGClient) & " " & slVehName & " Network Feed Instructions " & smTranDate
        'If Not mExportLine(slRecord, ilLineNo) Then
        '    Exit Function
        'End If
        If ((rbcInterface(0).Value) And (rbcFormat(0).Value)) Or ((rbcInterface(1).Value) And (rbcFormat(2).Value)) Then
            slRecord = smTranDate
            Do While Len(slRecord) < 11
                slRecord = slRecord & " "
            Loop
            If tlStnInfo.sType = "G" Then
                slRecord = slRecord & UCase$(slVehName & " Network Commercial Instructions")
            Else
                slRecord = slRecord & Trim$(tlStnInfo.sCallLetter) & "-" & tlStnInfo.sBand & ", " & UCase$(slVehName & " Network Commercial Instructions")
            End If
            If Not mExportLine(slRecord, ilLineNo, 1) Then
                mExportRotHeader = False
                Exit Function
            End If
        Else
            slRecord = UCase$(Trim$(tgSpf.sGClient))
            slRecord = slRecord & "          " & smTranDate
            If Not mExportLine(slRecord, ilLineNo, 5) Then
                mExportRotHeader = False
                Exit Function
            End If
            If Not mExportLine(slBlank, ilLineNo, 5) Then
                mExportRotHeader = False
                Exit Function
            End If
            If Not mExportLine(slBlank, ilLineNo, 5) Then
                mExportRotHeader = False
                Exit Function
            End If
            If tlStnInfo.sType = "G" Then
                slRecord = UCase$(slVehName & " Network Commercial Instructions")
            Else
                slRecord = Trim$(tlStnInfo.sCallLetter) & "-" & tlStnInfo.sBand & ", " & UCase$(slVehName & " Network Commercial Instructions")
            End If
            If Not mExportLine(slRecord, ilLineNo, 5) Then
                mExportRotHeader = False
                Exit Function
            End If
        End If
        If Not mExportLine(slBlank, ilLineNo, 5) Then
            mExportRotHeader = False
            Exit Function
        End If
        If Not mExportLine(slBlank, ilLineNo, 5) Then
            mExportRotHeader = False
            Exit Function
        End If
    Else
        If ilNewHdRot = True Then
            ilNewHdRot = False
            If Not mExportLine(slBlank, ilLineNo, 5) Then
                mExportRotHeader = False
                Exit Function
            End If
            'slRecord = "-"
            'Do While Len(slRecord) < 60
            '    slRecord = slRecord & "-"
            'Loop
            'If Not mExportLine(slRecord, ilLineNo) Then
            '    Exit Function
            'End If
            'If Not mExportLine(slBlank, ilLineNo) Then
            '    Exit Function
            'End If
        End If
    End If
    mExportRotHeader = True
End Function

Private Sub mExportSendMsg(slMsgLine As String, slMsgFileName As String, slMsgFile As String, ilMsgType As Integer, ilPrtFirstRot As Integer, ilNewHdRot As Integer, ilLineNo As Integer, ilPageNo As Integer, slVehName As String, tlStnInfo As STNINFO, slBlank As String, slRecord As String)
    Dim ilRet As Integer
    Dim hlMsg As Integer
    Dim ilPos As Integer
    
    ilRet = 0
    'On Error GoTo mExportSendMsgErr:
    'hlMsg = FreeFile
    slMsgFile = sgExportPath & slMsgFileName
    'Open slMsgFile For Input Access Read As hlMsg
    ilRet = gFileOpen(slMsgFile, "Input Access Read", hlMsg)
    If ilRet = 0 Then
        err.Clear
        Do
            'On Error GoTo mExportSendMsgErr:
            Line Input #hlMsg, slMsgLine
            On Error GoTo 0
            ilRet = err.Number
            If (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            End If
            If Len(slMsgLine) > 0 Then
                If (Asc(slMsgLine) = 26) Then    'Ctrl Z
                    Exit Do
                End If
                ilPos = InStr(UCase$(slMsgLine), "XX/XX/XXXX")
                If ilPos > 0 Then
                    Mid$(slMsgLine, ilPos) = smTranDate
                End If
            End If
            If ilMsgType = 0 Then
                If Not mExportLine(slMsgLine, ilLineNo, -1) Then
                    Exit Sub
                End If
                '6/3/16: Replaced GoSub
                'GoSub cmcExportCopyHeader
                mExportCopyHeader
            ElseIf ilMsgType = 1 Then
                If Not mExportLine(slMsgLine, ilLineNo, 5) Then
                    Exit Sub
                End If
                '6/3/16: Replaced GoSub
                'GoSub cmcExportRotHeader
                If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                    Exit Sub
                End If
            End If
        Loop
        Close hlMsg
    End If
    Exit Sub
'mExportSendMsgErr:
'    ilRet = Err.Number
'    Resume Next
End Sub

Private Function mSpecialInstructions(llRotStartDate As Long, llRotEndDate As Long, ilIncludeNewMessage As Integer, ilPrtFirstRot As Integer, ilNewHdRot As Integer, ilLineNo As Integer, ilPageNo As Integer, slVehName As String, tlStnInfo As STNINFO, slBlank As String, slRecord As String) As Integer
    Dim slStr As String
    Dim slTime As String
    Dim llDate As Long
    Dim ilDayReq As Integer
    Dim ilDay As Integer
    Dim ilPos As Integer
    Dim ilRet As Integer
    Dim slComment As String
    ReDim ilDayOn(0 To 6) As Integer
    
    If Not mExportLine(slBlank, ilLineNo, 5) Then
        mSpecialInstructions = False
        Exit Function
    End If
    slRecord = "Instructions:"
    Do While Len(slRecord) < 15
        slRecord = slRecord & " "
    Loop
    If (tmCrf.iStartTime(0) <> 0) Or (tmCrf.iStartTime(1) <> 0) Or (tmCrf.iEndTime(0) <> 0) Or (tmCrf.iEndTime(1) <> 0) Then
        'slRecord = "Air This Copy Between "
        slRecord = slRecord & "Air this copy between "
        gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slTime
        slTime = UCase(slTime)
        If slTime = "12AM" Then
            slTime = "12M"
        ElseIf slTime = "12PM" Then
            slTime = "12N"
        End If
        'If slTime = "12M" Then
        '    slTime = "12AM"
        'ElseIf slTime = "12N" Then
        '    slTime = "12PM"
        'End If
        slRecord = slRecord & slTime & " and "
        gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slTime
        slTime = UCase(slTime)
        If slTime = "12AM" Then
            slTime = "12M"
        ElseIf slTime = "12PM" Then
            slTime = "12N"
        End If
        'If slTime = "12M" Then
        '    slTime = "12AM"
        'ElseIf slTime = "12N" Then
        '    slTime = "12PM"
        'End If
        slRecord = slRecord & slTime
        If Not mExportLine(slRecord, ilLineNo, 5) Then
            mSpecialInstructions = False
            Exit Function
        End If
        '6/3/16: Replaced GoSub
        'GoSub cmcExportRotHeader
        If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
            mSpecialInstructions = False
            Exit Function
        End If
        slRecord = " "
        Do While Len(slRecord) < 15
            slRecord = slRecord & " "
        Loop
    End If
    ilDayReq = False
    For ilDay = 0 To 6 Step 1
        ilDayOn(ilDay) = False
    Next ilDay
    If llRotEndDate - llRotStartDate >= 6 Then
        For ilDay = 0 To 6 Step 1
            If tmCrf.sDay(ilDay) <> "Y" Then
                ilDayReq = True
                ilDayOn(ilDay) = False
            Else
                ilDayOn(ilDay) = True
            End If
        Next ilDay
    Else
        For llDate = llRotStartDate To llRotEndDate Step 1
            ilDay = gWeekDayLong(llDate)
            If tmCrf.sDay(ilDay) <> "Y" Then
                'ilDayReq = True
                ilDayOn(ilDay) = False
            Else
                ilDayOn(ilDay) = True
            End If
        Next llDate
        ilDayReq = True
    End If
    If ilDayReq Then
        slStr = ""
        If (ilDayOn(0) = True) And (ilDayOn(1) = True) And (ilDayOn(2) = True) And (ilDayOn(3) = True) And (ilDayOn(4) = True) And (ilDayOn(5) = False) And (ilDayOn(6) = False) Then
            slStr = "Mon thru Fri"
        ElseIf (ilDayOn(0) = False) And (ilDayOn(1) = False) And (ilDayOn(2) = False) And (ilDayOn(3) = False) And (ilDayOn(4) = False) And (ilDayOn(5) = True) And (ilDayOn(6) = True) Then
            slStr = "Sat and Sun"
        Else
            For ilDay = 0 To 6 Step 1
                If ilDayOn(ilDay) = True Then
                    Select Case ilDay
                        Case 0
                            slStr = slStr & " Mon"
                        Case 1
                            slStr = slStr & " Tue"
                        Case 2
                            slStr = slStr & " Wed"
                        Case 3
                            slStr = slStr & " Thu"
                        Case 4
                            slStr = slStr & " Fri"
                        Case 5
                            slStr = slStr & " Sat"
                        Case 6
                            slStr = slStr & " Sun"
                    End Select
                End If
            Next ilDay
        End If
        slRecord = slRecord & "Air this copy on" & slStr
        If Not mExportLine(slRecord, ilLineNo, 5) Then
            mSpecialInstructions = False
            Exit Function
        End If
        '6/3/16: Replaced GoSub
        'GoSub cmcExportRotHeader
        If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
            mSpecialInstructions = False
            Exit Function
        End If
        slRecord = " "
        Do While Len(slRecord) < 15
            slRecord = slRecord & " "
        Loop
    End If
    'If (Trim$(tmCrf.sZone) <> "") Then
    If (Trim$(tmCrf.sZone) <> "") And (Trim$(tmCrf.sZone) <> "R") Then
        Select Case Trim$(tmCrf.sZone)
            Case "EST"
                slStr = "EASTERN TIME ZONE"
            Case "CST"
                slStr = "CENTRAL TIME ZONE"
            Case "MST"
                slStr = "MOUNTAIN TIME ZONE"
            Case "PST"
                slStr = "PACIFIC TIME ZONE"
            Case "R"
                tmRafSrchKey.lCode = tlStnInfo.lRafCode
                ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                If ilRet = BTRV_ERR_NONE Then
                    slStr = Trim$(tmRaf.sName)
                Else
                    slStr = ""
                End If
        End Select
        slRecord = slRecord & "Air this copy only if you are in the " & slStr
        If Not mExportLine(slRecord, ilLineNo, 5) Then
            mSpecialInstructions = False
            Exit Function
        End If
        '6/3/16: Replaced GoSub
        'GoSub cmcExportRotHeader
        If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
            mSpecialInstructions = False
            Exit Function
        End If
        slRecord = " "
        Do While Len(slRecord) < 15
            slRecord = slRecord & " "
        Loop
    End If
    If tmCrf.lCsfCode <> 0 Then
        tmCsfSrchKey.lCode = tmCrf.lCsfCode
        tmCsf.sComment = ""
        imCsfRecLen = Len(tmCsf) '5011
        ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            'Output 70 characters per line
            'If tmCsf.iStrLen > 0 Then
            slStr = gStripChr0(tmCsf.sComment)
            If slStr <> "" Then
                slComment = slStr 'Trim$(Left$(tmCsf.sComment, tmCsf.iStrLen))
                If Not ilIncludeNewMessage Then
                    If InStr(1, slComment, "Revised For New", vbTextCompare) > 0 Then
                        slComment = ""
                    End If
                End If
                Do While Len(slComment) > 0
                    'Repeat all CR/LF with Space/LF
                    For ilPos = 1 To Len(slComment) Step 1
                        If Asc(Mid$(slComment, ilPos, 1)) = Asc(sgCR) Then
                            Mid$(slComment, ilPos, 1) = " "
                        End If
                    Next ilPos
                    ilPos = InStr(slComment, " ")
                    If ilPos > 0 Then
                        If Len(slRecord) + ilPos - 1 > 70 Then
                            If Not mExportLine(slRecord, ilLineNo, 5) Then
                                mSpecialInstructions = False
                                Exit Function
                            End If
                            '6/3/16: Replaced GoSub
                            'GoSub cmcExportRotHeader
                            If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                mSpecialInstructions = False
                                Exit Function
                            End If
                            slRecord = " "
                            Do While Len(slRecord) < 15
                                slRecord = slRecord & " "
                            Loop
                        End If
                        slRecord = slRecord & Left$(slComment, ilPos)
                        slComment = right$(slComment, Len(slComment) - ilPos)
                        If (Asc(slComment) = Asc(sgLF)) Then
                            If Not mExportLine(slRecord, ilLineNo, 5) Then
                                mSpecialInstructions = False
                                Exit Function
                            End If
                            '6/3/16: Replaced GoSub
                            'GoSub cmcExportRotHeader
                            If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                mSpecialInstructions = False
                                Exit Function
                            End If
                            slRecord = " "
                            Do While Len(slRecord) < 15
                                slRecord = slRecord & " "
                            Loop
                            slComment = right$(slComment, Len(slComment) - 1)
                        End If
                    Else
                        If Len(slRecord) + Len(slComment) > 70 Then
                            If Not mExportLine(slRecord, ilLineNo, 5) Then
                                mSpecialInstructions = False
                                Exit Function
                            End If
                            '6/3/16: Replaced GoSub
                            'GoSub cmcExportRotHeader
                            If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                                mSpecialInstructions = False
                                Exit Function
                            End If
                            slRecord = " "
                            Do While Len(slRecord) < 15
                                slRecord = slRecord & " "
                            Loop
                        End If
                        slRecord = slRecord & slComment
                        If Not mExportLine(slRecord, ilLineNo, 5) Then
                            mSpecialInstructions = False
                            Exit Function
                        End If
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportRotHeader
                        If Not mExportRotHeader(ilPrtFirstRot, ilNewHdRot, ilLineNo, ilPageNo, slVehName, tlStnInfo, slBlank, slRecord) Then
                            mSpecialInstructions = False
                            Exit Function
                        End If
                        slRecord = " "
                        Do While Len(slRecord) < 15
                            slRecord = slRecord & " "
                        Loop
                        slComment = ""
                        Exit Do
                    End If
                Loop
            End If
        End If
    End If
    If InStr(1, slRecord, "Instructions:", 1) <= 0 Then
        If Not mExportLine(slBlank, ilLineNo, 5) Then
            mSpecialInstructions = False
            Exit Function
        End If
    End If
    mSpecialInstructions = True
End Function

Private Sub mProcAdjDate(ilAirHour As Integer, ilLocalHour As Integer, slAdjDate As String)
    'Test if Air time is AM and Local Time is PM. If so, adjust date
    ilAirHour = tmAvail.iTime(1) \ 256  'Obtain month
    ilLocalHour = tmDlf.iLocalTime(1) \ 256  'Obtain month
    If (ilAirHour < 6) And (ilLocalHour > 17) Then
        'If monday convert to next sunday- this is wrong but the same spot
        'runs each sunday (the spot should have show on the previous week sunday)
        'If not monday, then subtract one day
        If gWeekDayStr(slAdjDate) = 0 Then
            slAdjDate = gObtainNextSunday(slAdjDate)
        Else
            slAdjDate = gDecOneDay(slAdjDate)
        End If
    End If
End Sub

Private Sub mProcSpot(ilVpfIndex As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, ilAirHour As Integer, ilLocalHour As Integer, slAdjDate As String, ilDlfFound As Integer, ilSeqNo As Integer, ilRet As Integer)
    tmCpr.iGenDate(0) = imGenDate(0)
    tmCpr.iGenDate(1) = imGenDate(1)
    'tmCpr.iGenTime(0) = imGenTime(0)
    'tmCpr.iGenTime(1) = imGenTime(1)
    gUnpackTimeLong imGenTime(0), imGenTime(1), False, tmCpr.lGenTime
    'Air Vehicle
    tmCpr.iVefCode = tmEVef.iCode
    'EDAS Time Window
    tmCpr.lHd1CefCode = 400
    If tgVpf(ilVpfIndex).lEDASWindow > 0 Then
        tmCpr.lHd1CefCode = tgVpf(ilVpfIndex).lEDASWindow
    End If
    'Air Date
    gUnpackDate ilLogDate0, ilLogDate1, slAdjDate
    '6/3/16:Replaced GoSub
    'GoSub lProcAdjDate  'Adjust dates prior to adjusting seq numbers
    mProcAdjDate ilAirHour, ilLocalHour, slAdjDate
    gPackDate slAdjDate, tmCpr.iSpotDate(0), tmCpr.iSpotDate(1)
    'Air Time
    tmCpr.iSpotTime(0) = tmDlf.iLocalTime(0)
    tmCpr.iSpotTime(1) = tmDlf.iLocalTime(1)
    'Spot Length
    tmCpr.iLen = tmSdf.iLen
    'Sdf Code
    tmCpr.lCntrNo = tmSdf.lCode
    'If (ilDlfFound) Or (tmSdf.sPtType <> "3") Then
    '    tmCpr.sZone = tmDlf.sZone
    '    ilRet = mObtainCopy(tmDlf.sZone)
    '    If ilRet Then
    '        If Trim$(tmCif.sCut) = "" Then
    '            tmCpr.sCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName)
    '        Else
    '            tmCpr.sCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut)
    '        End If
    '    Else
    '        tmCpr.sCartNo = ""
    '    End If
    '    ilSeqNo = ilSeqNo + 1
    '    tmCpr.iLineNo = ilSeqNo
    '    ilRet = btrUpdate(hmCpr, tmCpr, imCprRecLen)
    'Else
    '    For ilZone = 1 To 4 Step 1
    '        Select Case ilZone
    '            Case 1
    '                slZone = "EST"
    '            Case 2
    '                slZone = "MST"
    '            Case 3
    '                slZone = "CST"
    '            Case 4
    '                slZone = "PST"
    '        End Select
    '        tmCpr.sZone = slZone
    '        ilRet = mObtainCopy(slZone)
    '        If ilRet Then
    '            If Trim$(tmCif.sCut) = "" Then
    '                tmCpr.sCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName)
    '            Else
    '                tmCpr.sCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut)
    '            End If
    '        Else
    '            tmCpr.sCartNo = ""
    '        End If
    '        ilSeqNo = ilSeqNo + 1
    '        tmCpr.iLineNo = ilSeqNo
    '        ilRet = btrUpdate(hmCpr, tmCpr, imCprRecLen)
    '    Next ilZone
    'End If
    If ilDlfFound Then
        tmCpr.sZone = tmDlf.sZone
    Else
        tmCpr.sZone = ""
    End If
    tmCpr.sCartNo = ""
    ilSeqNo = ilSeqNo + 1
    tmCpr.iLineNo = ilSeqNo
    ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
End Sub
