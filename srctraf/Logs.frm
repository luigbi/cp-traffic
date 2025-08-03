VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Logs 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6030
   ClientLeft      =   465
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   9345
   Begin VB.ListBox lbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      ItemData        =   "Logs.frx":0000
      Left            =   180
      List            =   "Logs.frx":0002
      TabIndex        =   71
      Top             =   5880
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.CommandButton cmcSplitFill 
      Appearance      =   0  'Flat
      Caption         =   "&Split Fill"
      Height          =   285
      Left            =   7320
      TabIndex        =   53
      Top             =   5625
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   180
      Left            =   0
      Picture         =   "Logs.frx":0004
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox plcInfo 
      BackColor       =   &H00FFFF80&
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
      Height          =   720
      Left            =   135
      ScaleHeight     =   660
      ScaleWidth      =   9105
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   1035
      Visible         =   0   'False
      Width           =   9165
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Start Time xx:xx:xxam  Length xx:xx:xx"
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
         Left            =   105
         TabIndex        =   68
         Top             =   435
         Width           =   8925
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Start Time xx:xx:xxam  Length xx:xx:xx"
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
         Left            =   105
         TabIndex        =   67
         Top             =   225
         Width           =   8925
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Library Name xxxxxxxxxxxxxxxxxx  Version xx"
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
         Left            =   105
         TabIndex        =   66
         Top             =   30
         Width           =   8970
      End
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   7575
      Top             =   5850
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".Txt"
      Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
      FilterIndex     =   1
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.PictureBox plcLogMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
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
      Height          =   4050
      Left            =   1395
      ScaleHeight     =   4050
      ScaleWidth      =   6945
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   6945
      Begin VB.ListBox lbcLogMsg 
         Appearance      =   0  'Flat
         Height          =   3180
         Left            =   150
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   285
         Width           =   6630
      End
      Begin VB.CommandButton cmcLogMsgOk 
         Appearance      =   0  'Flat
         Caption         =   "&Ok"
         Height          =   285
         Left            =   2985
         TabIndex        =   31
         Top             =   3645
         Width           =   945
      End
   End
   Begin VB.CommandButton cmcBlackout 
      Appearance      =   0  'Flat
      Caption         =   "&Blackout"
      Height          =   285
      Left            =   6090
      TabIndex        =   52
      Top             =   5625
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CheckBox ckcCheckOn 
      Caption         =   "Set Gen Checks"
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
      Left            =   240
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5640
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin VB.PictureBox pbcRptSample 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1020
      Index           =   0
      Left            =   2790
      ScaleHeight     =   990
      ScaleWidth      =   6600
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2355
      Visible         =   0   'False
      Width           =   6630
      Begin VB.PictureBox pbcRptSample 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Index           =   1
         Left            =   0
         ScaleHeight     =   165
         ScaleWidth      =   4530
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   0
         Width           =   4530
      End
   End
   Begin VB.TextBox edcInfo 
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
      Height          =   1020
      Left            =   255
      MultiLine       =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.ListBox lbcOther 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   7815
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1545
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcLogo 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6810
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2145
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcTimeZ 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6105
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lbcCP 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4875
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1770
      Visible         =   0   'False
      Width           =   945
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
      Left            =   4410
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2655
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Logs.frx":030E
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   17
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
            TabIndex        =   18
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
         TabIndex        =   14
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
         TabIndex        =   16
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
         Left            =   315
         TabIndex        =   15
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.ListBox lbcLog 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1035
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   7485
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2430
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "Logs.frx":3128
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "Logs.frx":3DE6
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
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
      End
   End
   Begin VB.FileListBox lbcFile 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6555
      Pattern         =   "g???.bmp;g???.jpg;g???.gif"
      TabIndex        =   45
      Top             =   -15
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6750
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   4620
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.VScrollBar vbcLogs 
      Height          =   4275
      LargeChange     =   19
      Left            =   9000
      Min             =   1
      TabIndex        =   21
      Top             =   1125
      Value           =   1
      Width           =   255
   End
   Begin VB.CommandButton cmcLogChk 
      Appearance      =   0  'Flat
      Caption         =   "C&heck"
      Height          =   285
      Left            =   4995
      TabIndex        =   34
      Top             =   5625
      Width           =   945
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
      Left            =   -30
      ScaleHeight     =   165
      ScaleWidth      =   105
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4455
      Width           =   105
   End
   Begin VB.PictureBox pbcTab 
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
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   20
      Top             =   1485
      Width           =   60
   End
   Begin VB.PictureBox pbcSTab 
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
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   240
      Width           =   15
   End
   Begin VB.PictureBox pbcSelections 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   570
      ScaleHeight     =   210
      ScaleWidth      =   375
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox edcDropDown 
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
      Left            =   420
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcDropDown 
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
      Left            =   1440
      Picture         =   "Logs.frx":40F0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   540
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   540
   End
   Begin VB.CommandButton cmcGenerate 
      Appearance      =   0  'Flat
      Caption         =   "&Generate"
      Height          =   285
      Left            =   2415
      TabIndex        =   32
      Top             =   5625
      Width           =   1365
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3900
      TabIndex        =   33
      Top             =   5625
      Width           =   945
   End
   Begin VB.ListBox lbcFeedCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   7350
      Sorted          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5610
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ListBox lbcVehCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   7320
      Sorted          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5505
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox edcLinkSrceHelpMsg 
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
      Left            =   240
      TabIndex        =   39
      Top             =   1875
      Visible         =   0   'False
      Width           =   210
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
      Left            =   4800
      TabIndex        =   42
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
      Left            =   5310
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   -45
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox plcStatus 
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
      Height          =   210
      Left            =   105
      ScaleHeight     =   210
      ScaleWidth      =   9120
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   5415
      Width           =   9120
   End
   Begin VB.PictureBox pbcLogs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Index           =   0
      Left            =   255
      Picture         =   "Logs.frx":41EA
      ScaleHeight     =   4275
      ScaleWidth      =   8745
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   8745
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   43
         Top             =   345
         Visible         =   0   'False
         Width           =   8730
      End
   End
   Begin VB.PictureBox pbcLogs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Index           =   1
      Left            =   255
      Picture         =   "Logs.frx":7E77C
      ScaleHeight     =   4275
      ScaleWidth      =   8745
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   8745
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   51
         Top             =   345
         Visible         =   0   'False
         Width           =   8730
      End
   End
   Begin VB.PictureBox plcLogs 
      ForeColor       =   &H00000000&
      Height          =   4395
      Left            =   120
      ScaleHeight     =   4335
      ScaleWidth      =   9075
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1035
      Width           =   9135
   End
   Begin VB.ListBox lbcAdvt 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8205
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   -105
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.PictureBox plcLogInfo 
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
      Height          =   960
      Left            =   585
      ScaleHeight     =   900
      ScaleWidth      =   8235
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   15
      Width           =   8295
      Begin VB.CheckBox ckcAssignCopy 
         Caption         =   "Assign Copy"
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
         Left            =   4695
         TabIndex        =   29
         Top             =   540
         Width           =   1395
      End
      Begin VB.CommandButton cmcTime 
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
         Index           =   1
         Left            =   7935
         Picture         =   "Logs.frx":F8D0E
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   210
         Width           =   195
      End
      Begin VB.TextBox edcTime 
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
         Left            =   7110
         MaxLength       =   20
         TabIndex        =   27
         Top             =   210
         Width           =   825
      End
      Begin VB.CommandButton cmcTime 
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
         Index           =   0
         Left            =   6435
         Picture         =   "Logs.frx":F8E08
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   210
         Width           =   195
      End
      Begin VB.TextBox edcTime 
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
         Left            =   5595
         MaxLength       =   20
         TabIndex        =   24
         Top             =   210
         Width           =   825
      End
      Begin VB.Frame frcLog 
         Caption         =   "Logs"
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
         Height          =   900
         Left            =   105
         TabIndex        =   54
         Top             =   -15
         Width           =   1890
         Begin VB.OptionButton rbcLogType 
            Caption         =   "Internal Reprint"
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
            Left            =   60
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   645
            Width           =   1650
         End
         Begin VB.OptionButton rbcLogType 
            Caption         =   "Alert"
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
            Left            =   1020
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   420
            Width           =   720
         End
         Begin VB.OptionButton rbcLogType 
            Caption         =   "Prelim"
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
            Left            =   60
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   195
            Width           =   900
         End
         Begin VB.OptionButton rbcLogType 
            Caption         =   "Final"
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
            Left            =   1020
            TabIndex        =   56
            Top             =   195
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton rbcLogType 
            Caption         =   "Reprint"
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
            Left            =   60
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   420
            Width           =   960
         End
      End
      Begin VB.Frame frcOutput 
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
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   2070
         TabIndex        =   58
         Top             =   60
         Width           =   2520
         Begin VB.CheckBox ckcOutput 
            Caption         =   "Display"
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
            Left            =   60
            TabIndex        =   59
            Top             =   255
            Width           =   990
         End
         Begin VB.CheckBox ckcOutput 
            Caption         =   "Print"
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
            Left            =   1095
            TabIndex        =   60
            Top             =   255
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox ckcOutput 
            Caption         =   "File"
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
            Left            =   60
            TabIndex        =   61
            Top             =   510
            Width           =   645
         End
         Begin VB.ComboBox cbcFile 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   705
            TabIndex        =   62
            Top             =   480
            Width           =   1740
         End
      End
      Begin VB.Label lacTime 
         Appearance      =   0  'Flat
         Caption         =   "End"
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
         Left            =   6750
         TabIndex        =   26
         Top             =   210
         Width           =   450
      End
      Begin VB.Label lacTime 
         Appearance      =   0  'Flat
         Caption         =   "Time: Start"
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
         Left            =   4635
         TabIndex        =   23
         Top             =   210
         Width           =   1005
      End
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   30
      Picture         =   "Logs.frx":F8F02
      Top             =   480
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8895
      Top             =   5625
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Logs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Logs.frm on Wed 6/17/09 @ 12:56 PM **
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmGhfSrchKey0                 tmGsfSrchKey0                 imRafRecLen               *
'*  smTeamCodeTag                                                                         *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'************************************************************
' File Name: Logs.Frm
'
' Release: 1.0
'               Created: ?          By: D. LeVine
'               Modified : 5/4/94  By: D. Hannifan
'
' Description:
'   This file contains the Log input screen code
'************************************************************
Option Explicit
Option Compare Text
Dim hmMsg As Integer   'From file hanle
Dim imFirstActivate As Integer
'Btrieve files
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0     'VEF key 0 image
Dim imVefRecLen As Integer      'VEF record length
Dim hmVef As Integer            'Vehicle file handle
Dim tmLogGen() As LOGGEN
Dim tmVpf As VPF                'VPF record image
Dim tmVpfSrchKey As VPFKEY0     'VPF key 0 image
Dim imVpfRecLen As Integer      'VPF record length
Dim hmVpf As Integer            'Vehicle preference file handle
Dim tmVlf As VLF                'VLF record image
Dim imVlfRecLen As Integer      'VLF record length
Dim hmVLF As Integer            'Vehicle Link file handle
Dim tmVlfSrchKey1 As VLFKEY1
Dim tmStf As STF                'STF record image
Dim tmStfSrchKey As STFKEY0     'STF key 0 image
Dim imStfRecLen As Integer      'STF record length
Dim hmStf As Integer            'Spot Tracking file handle
Dim lmStfRecPos As Long
'One day file (ODF)
Dim hmOdf As Integer        'One day file
Dim imOdfRecLen As Integer  'ODF record length
Dim tmOdf As ODF            'ODF record image
Dim tmOdfSrchKey2 As ODFKEY2    'GSF key record image
'Log Spot record
Dim hmLst As Integer        'Log Spots file
Dim tmLst As LST
Dim imLstRecLen As Integer
Dim tmLstSrchKey2 As LSTKEY2    'LST key record image
Dim imLstExist As Integer       'False=LST Does NOT exist, don't update
'Contract header record information
Dim hmCHF As Integer        'Contract file handle
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey0 As LONGKEY0
'Contract line record information
Dim hmClf As Integer        'Contract line file handle
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image
'Blackout record information
Dim hmBof As Integer        'Blackout file handle
Dim imBofRecLen As Integer  'BOF record length
Dim tmBof As BOF            'BOF record image
'Prduct code file
Dim hmPrf As Integer 'Product file handle
Dim tmPrf As PRF        'PRF record image
Dim imPrfRecLen As Integer
'Short Title record information
Dim hmSif As Integer        'Short Title file handle
Dim imSifRecLen As Integer  'SIF record length
Dim tmSif As SIF            'SIF record image
'Avail Name record information
Dim hmAnf As Integer        'Avail Name file handle
Dim imAnfRecLen As Integer  'MCF record length
Dim tmAnf As ANF            'MCF record image
Dim tmAnfSrchKey0 As INTKEY0
'Media code record information
Dim hmMcf As Integer        'Contract line file handle
Dim imMcfRecLen As Integer  'MCF record length
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey0 As INTKEY0
'Copy inventory record information
Dim hmCif As Integer        'Copy line file handle
Dim imCifRecLen As Integer  'CIF record length
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey0 As LONGKEY0
'Copy Product
Dim hmCpf As Integer        'Copy Product file handle
Dim imCpfRecLen As Integer  'CPF record length
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey0 As LONGKEY0
'Copy Rotation
Dim hmCrf As Integer        'Copy Rotation file handle
Dim imCrfRecLen As Integer  'CNF record length
Dim tmCrf As CRF            'CNF record image
Dim tmCrfSrchKey As LONGKEY0 'CIF key record image
'Copy
Dim hmCnf As Integer        'Copy file handle
Dim imCnfRecLen As Integer  'CNF record length
Dim tmCnf As CNF            'CNF record image
'Alerts
Dim tmAuf As AUF        'Rvf record image
Dim tmAufSrchKey1 As AUFKEY1    'Rvf key record image
Dim imAufRecLen As Integer        'RvF record length
'Regional or Blackout copy
Dim hmRsf As Integer        'Regional or Blackout copy file handle
Dim imRsfRecLen As Integer  'RSF record length
Dim tmRsf As RSF            'RSF record image
'11/6/14: add key
Dim tmRsfSrchKey1 As LONGKEY0 'RSF key record image
Dim tmRnf As RNF                'RNF record image
Dim imRnfRecLen As Integer      'RnF record length
Dim hmRnf As Integer            'Report Name file handle
Dim hmSsf As Integer            'Spot Summary file handle
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim imSsfRecLen As Integer
Dim hmSdf As Integer
'11/6/14
Dim tmSdfSrchKey3 As LONGKEY0 'SDF key record image (code)
Dim imSdfRecLen As Integer     'SDF record length
Dim tmSdf As SDF

Dim hmGhf As Integer
Dim tmGhf As GHF        'GHF record image
Dim tmCombineGhf As GHF        'GHF record image
Dim tmGhfSrchKey1 As GHFKEY1    'GHF key record image
Dim imGhfRecLen As Integer        'GHF record length

Dim hmGsf As Integer
Dim tmGsf As GSF        'GSF record image
Dim tmCombineGsf As GSF        'GSF record image
Dim tmGsfSrchKey1 As GSFKEY1    'GSF key record image
Dim tmGsfSrchKey2 As GSFKEY2    'GSF key record image
Dim tmGsfSrchKey3 As GSFKEY3    'GSF key record image
Dim tmGsfSrchKey4 As GSFKEY4    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length

Dim hmRaf As Integer
Dim tmRaf As RAF
Dim tmRafSrchKey0 As LONGKEY0

Dim hmCvf As Integer

'AST Queue record information
Dim hmAbf As Integer        'Avail Name file handle
Dim imAbfRecLen As Integer  'MCF record length
Dim tmAbf As ABF            'MCF record image
Dim tmAbfSrchKey0 As LONGKEY0
Dim tmAbfSrchKey2 As ABFKEY2

Dim tmRBofRec() As SPLITBOFREC
Dim tmSplitNetLastFill() As SPLITNETLASTFILL

Dim tmTeam() As MNF
Dim smTeamTag As String

Dim tmLang() As MNF
Dim smLangTag As String

Dim tmAdvertiser() As SORTCODE
Dim smAdvertiserTag As String
'Dim tmRec As LPOPREC
Dim imGettingMkt As Integer
'Affiliate Btrieve files
Dim imATTExist As Integer
Dim tmAtt As ATT                'ATT record image
Dim tmVATT() As ATT
Dim tmLogVATT() As ATT          'tmVATT build for log vehicles
Dim tmATTSrchKey1 As INTKEY0     'ATT key 1 image
Dim imAttRecLen As Integer      'ATT record length
Dim hmAtt As Integer            'Agreement file handle
Dim tmCPTT As CPTT                'CPTT record image
Dim tmCPTTSrchKey As LONGKEY0     'CPTT key 0 image
Dim tmCPTTSrchKey1 As CPTTKEY1     'CPTT key 1 image
Dim tmCPTTSrchKey2 As CPTTKEY2     'CPTT key 2 image
Dim imCPTTRecLen As Integer      'CPTT record length
Dim hmCPTT As Integer            'Cert. of Perfor file handle
Dim tmCPTTInfo() As CPTTINFO
Dim tmSHTT As SHTT                'SHTT record image
Dim tmSHTTSrchKey As INTKEY0     'SHTT key 0 image
Dim imSHTTRecLen As Integer      'SHTT record length
Dim hmSHTT As Integer            'Station file handle

'11/7/14
'Split Entity
Dim tmSef As SEF            'SEF record image
Dim tmSefSrchKey1 As SEFKEY1  'SEF key record image
Dim hmSef As Integer        'SEF Handle
Dim imSefRecLen As Integer      'SEF record length

'Audio Vault To X-Digital Indicator ID
Dim hmAxf As Integer
Dim tmAxf As AXF
Dim tmAxfSrchKey1 As INTKEY0 'ANF key record image
Dim imAxfRecLen As Integer  'ANF record length

'LOGS flags
Dim imTerminate As Integer      'True=terminate  False = OK
Dim imFirstTime As Integer
Dim imChgMode As Integer        'True=value changed
Dim imLcfFound As Integer       'True=valid Lcf date found
Dim imCpyFlag As Integer        'True=don't allow in field ; default Assign copy to yes
'LOGS modular variables
Dim smLogType As String         'P=Preliminary; F=Final; R=Reprint
Dim imPFAllowed As Integer      'True=Preliminary/Final Logs; False=Final only
Dim imAssCopy As Integer        'True=Allowed to assign copy; False=Not allowed to assign copy
Dim imVehCode As Integer        'vehicle code
Dim imFeedCode As Integer       'Feed code
Dim smDate As String            'now +1
Dim smNowDate As String
Dim lmNowDate As Long
Dim smDefaultDate As String     'default start date
Dim smDefaultTime As String     'default start time
Dim imDelivery As Integer       'True=Delivery file exist; False=No delivery file
Dim imComboBoxIndex As Integer  'Previous ListIndex value
Dim imLbcArrowSetting As Integer 'True = Processing arrow key- retain current list box visible
                                'False= Make list box invisible
Dim imGeneratingLog As Integer
Dim imRPGen As Integer          'Reprint Generated
Dim imAlertGen As Integer
Dim imLogType As Integer        'rbcLogType Index
Dim imPbcIndex As Integer
Dim imTmeIndex As Integer
Dim imCurrentIndex As Integer
Dim imShiftKey As Integer   'Bit 0=Shift; 1=Ctrl; 2=Alt
Dim imButton As Integer
Dim imButtonIndex As Integer
Dim imIgnoreRightMove As Integer
Dim imButtonRow As Integer
'Calendar variables
Dim tmCDCtrls(0 To 7) As FIELDAREA  'Field area image
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer        'Month of displayed calendar
Dim imCalMonth As Integer       'Year of displayed calendar
Dim lmCalStartDate As Long      'Start date of displayed calendar
Dim lmCalEndDate As Long        'End date of displayed calendar
Dim imCalType As Integer        'Calendar type
Dim imBSMode As Integer         'Backspace flag
Dim imBypassFocus As Integer    'Bypass gotfocus
Dim imShowHelpMsg As Integer    'True=Show help message; False=Ignore help message system
Dim smNewLines(0 To 0) As String * 72   'required as parameter to gBlackoutTest
Dim smStatusCaption As String
'Tabs
Dim tmCtrls(0 To 13)  As FIELDAREA   'Field area image
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current event name Box
Dim imRowNo As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imSettingValue As Integer

Dim imCurSort As Integer    '0=Sort by Work Day; 1=Sort by Vehicle name
Dim imUpdateAllowed As Integer    'User can update records

Dim tmRnfList() As RNFLIST

Dim tmTeamCode() As SORTCODE

Dim hmTo As Integer

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Dim tmRegionInfo() As REGIONINFO '5-16-08 list of regions generated in the log

Dim rst_cptt As ADODB.Recordset

'Constants
Const CHKINDEX = 1
Const WRKDATEINDEX = 2
Const VEHINDEX = 3           'Log or Commercial Scgedule format control/field
Const LLDINDEX = 4
Const LEADTIMEINDEX = 5
Const CYCLEINDEX = 6
Const SDATEINDEX = 7            'Start Date control/field
Const EDATEINDEX = 8            'Start Date control/field
Const LOGINDEX = 9
Const CPINDEX = 10
Const LOGOINDEX = 11
Const OTHERINDEX = 12
Const ZONEINDEX = 13
'Const SDATEINDEX = 2            'Start Date control/field
Const ENDTIMEINDEX = 5          'End time control/field

Private Type EXPORTCOPYINFO
    iRotNo As Integer
    lRafCode As Long
    lCrfCode As Long
    sPtType As String * 1
    lCopyCode As Long
End Type
Dim tmExportCopyInfo() As EXPORTCOPYINFO
Private Type STATIONEXPORTINFO
    iShttCode As Integer
    iRotNo As Integer
    sSource As String * 1   'I=Include; E=Exclude
    sExport As String * 1    'Y=Yes; N=No
    iExportCopyInfoIndex As Integer
End Type
Dim tmStationExportInfo() As STATIONEXPORTINFO

Private Sub ckcAssignCopy_GotFocus()
    plcTme.Visible = False
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub ckcCheckOn_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcCheckOn.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilLoop As Integer
    If Value Then
        For ilLoop = 0 To UBound(tgSel) - 1 Step 1
            If tgSel(ilLoop).iStatus = 0 Then
                tgSel(ilLoop).iChk = tgSel(ilLoop).iInitChk
            Else
                tgSel(ilLoop).iChk = 0
            End If
        Next ilLoop
    Else
        For ilLoop = 0 To UBound(tgSel) - 1 Step 1
            tgSel(ilLoop).iChk = 0
        Next ilLoop
    End If
    pbcLogs(imPbcIndex).Cls
    pbcLogs_Paint imPbcIndex
End Sub
Private Sub ckcCheckOn_GotFocus()
    plcTme.Visible = False
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub ckcOutput_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcOutput(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    If Index = 2 Then
        If Value Then
            cbcFile.Enabled = True
            If cbcFile.ListIndex < 0 Then
                'If (Trim$(tgUrf(0).sPDFDrvChar) <> "") And (tgUrf(0).iPDFDnArrowCnt >= 0) And (Trim$(tgUrf(0).sPrtDrvChar) <> "") And (tgUrf(0).iPrtDnArrowCnt >= 0) Then
                    'cbcFile.ListIndex = 7
                    cbcFile.ListIndex = 0   'adobe 10-19-01
                'Else
                '    'cbcFile.ListIndex = 6
                '    cbcFile.ListIndex = 5       'rtf 10-19-01
                'End If
            End If
        Else
            cbcFile.Enabled = False
        End If
    Else
        If ckcOutput(2).Value = vbUnchecked Then
            cbcFile.Enabled = False
        End If
    End If
End Sub
Private Sub cmcBlackout_Click()
    'igShowHelpMsg = imShowHelpMsg
    'Blackout.Show vbModal
    Dim slStr As String
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Logs^Test\" & sgUserName
        Else
            slStr = "Logs^Prod\" & sgUserName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Logs^Test^NOHELP\" & sgUserName
    '    Else
    '        slStr = "Logs^Prod^NOHELP\" & sgUserName
    '    End If
    'End If
    slStr = slStr & "\Log"
    'lgShellRet = Shell(sgExePath & "Blackout.Exe " & slStr, 1)
    'Logs.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    Blackout.Show vbModal
    slStr = sgDoneMsg
    'Logs.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCancel_Click()
    If imGeneratingLog Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    If imFirstTime Then
        imFirstTime = False
    End If
    plcTme.Visible = False
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcDropDown_Click()

    Select Case imBoxNo
        Case CHKINDEX
        Case LLDINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case LEADTIMEINDEX
        Case CYCLEINDEX
        Case SDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case LOGINDEX
            lbcLog.Visible = Not lbcLog.Visible
        Case CPINDEX
            lbcCP.Visible = Not lbcCP.Visible
        Case LOGOINDEX
            lbcLogo.Visible = Not lbcLogo.Visible
        Case OTHERINDEX
            lbcOther.Visible = Not lbcOther.Visible
        Case ZONEINDEX
            lbcTimeZ.Visible = Not lbcTimeZ.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcGenerate_Click()
    Dim ilRet As Integer       'Call return value
    Dim slDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilDeleteSTF As Integer
    Dim ilLoop As Integer
    Dim ilVef As Integer
    Dim ilLink As Integer
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim ilAirCopy As Integer
    Dim ilAnySelected As Integer
    Dim ilFound As Integer
    Dim ilValue As Integer
    Dim slVefName As String
    
    
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    Screen.MousePointer = vbHourglass
    ilAnySelected = False
    For ilLoop = 0 To UBound(tgSel) - 1 Step 1
        If tgSel(ilLoop).iChk = 1 Then
            ilAnySelected = True
        End If
    Next ilLoop
    If Not ilAnySelected Then
        Screen.MousePointer = vbDefault
        Beep
        MsgBox "No Vehicle Selected", 48, "Log Generation"
        Exit Sub
    End If
    sgVpfStamp = "~"    'Force read
    ilRet = gVpfRead()
    slStartTime = edcTime(0).Text
    slEndTime = edcTime(1).Text
    If gTimeToCurrency(slEndTime, True) < gTimeToCurrency(slStartTime, False) Then
        Screen.MousePointer = vbDefault
        Beep
        MsgBox "End Time earlier than Start Time.", 48, "Invalid Time Value"
        imBoxNo = ENDTIMEINDEX
        mEnableBox imBoxNo
        Exit Sub
    End If
    If Not gRecLengthOk("odf.btr", Len(tmOdf)) Then
        Screen.MousePointer = vbDefault
        cmcCancel.SetFocus
        Exit Sub
    End If
    If tgSpf.sGUseAffSys = "Y" Then
        If Not gRecLengthOk("att.mkd", Len(tmAtt)) Then
            Screen.MousePointer = vbDefault
            cmcCancel.SetFocus
            Exit Sub
        End If
        If Not gRecLengthOk("CPTT.Mkd", Len(tmCPTT)) Then
            Screen.MousePointer = vbDefault
            cmcCancel.SetFocus
            Exit Sub
        End If
        If Not gRecLengthOk("SHTT.Mkd", Len(tmSHTT)) Then
            Screen.MousePointer = vbDefault
            cmcCancel.SetFocus
            Exit Sub
        End If
        If Not gRecLengthOk("Lst.Mkd", Len(tmLst)) Then
            Screen.MousePointer = vbDefault
            cmcCancel.SetFocus
            Exit Sub
        End If
        
    End If
    ilAirCopy = False
    For ilLoop = LBound(tgVpf) To UBound(tgVpf) Step 1
        If tgVpf(ilLoop).sCopyOnAir = "Y" Then
            ilAirCopy = True
            Exit For
        End If
    Next ilLoop
    ilValue = Asc(tgSpf.sUsingFeatures2)  'Option Fields in Orders/Proposals
    If (((ilValue And REGIONALCOPY) = REGIONALCOPY) Or ((ilValue And SPLITCOPY) = SPLITCOPY)) Or (tgSpf.sCBlackoutLog = "Y") Or (ilAirCopy) Then
        If Not gRecLengthOk("Rsf.btr", Len(tmRsf)) Then
            Screen.MousePointer = vbDefault
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    ilDeleteSTF = -1
    If (rbcLogType(0).Value) Then   'Prel
        smLogType = "P"
    ElseIf (rbcLogType(2).Value) Or (rbcLogType(3).Value) Then  'Reprint
        smLogType = "R"
    ElseIf (rbcLogType(3).Value) Then  'Alert
        smLogType = "A"
    ElseIf (rbcLogType(4).Value) Then   'Internal
        smLogType = "I"
    Else    'Final
        smLogType = "F"
        'Dan 7/14/2010 ask if want to do a log check--for 'final' only
        ilRet = MsgBox("Would you like to run a Log Check first?", vbQuestion + vbYesNo, "Run Log Check?")
        If ilRet = vbYes Then
            cmcLogChk_Click
            DoEvents
            ilRet = MsgBox("Press Ok to continue log generation, or Cancel to first fix any issues?", vbQuestion + vbOKCancel, "Continue Log Generation?")
            If ilRet = vbCancel Then
                Screen.MousePointer = vbDefault
                cmcCancel.SetFocus
                Exit Sub
            End If
        End If
    End If
    slVefName = ""
    If (tgSpf.sGUseAffSys = "Y") And (rbcLogType(4).Value = False) Then
        For ilLoop = 0 To UBound(tgSel) - 1 Step 1
            If tgSel(ilLoop).iChk = 1 Then
                slStartDate = gObtainPrevMonday(Format(tgSel(ilLoop).lStartDate, "mm/dd/yyyy"))
                slEndDate = Format(tgSel(ilLoop).lEndDate, "mm/dd/yyyy")
                SQLQuery = "SELECT Count(1) as AnyPosted FROM cptt WHERE"
                SQLQuery = SQLQuery & " cpttVefCode = " & tgSel(ilLoop).iVefCode
                SQLQuery = SQLQuery & " And cpttStartDate >= " & "'" & Format$(gAdjYear(slStartDate), sgSQLDateForm) & "'"
                SQLQuery = SQLQuery & " And cpttStartDate <= " & "'" & Format$(gAdjYear(slEndDate), sgSQLDateForm) & "'"
                SQLQuery = SQLQuery & " And cpttPostingStatus > 0"
                Set rst_cptt = gSQLSelectCall(SQLQuery)
                If Not rst_cptt.EOF Then
                    If rst_cptt!AnyPosted > 0 Then
                        ilVef = gBinarySearchVef(tgSel(ilLoop).iVefCode)
                        If ilVef <> -1 Then
                            If slVefName = "" Then
                                slVefName = Trim$(tgMVef(ilVef).sName) & ":" & Format(tgSel(ilLoop).lStartDate, "mm/dd/yyyy") & "-" & Format(tgSel(ilLoop).lEndDate, "mm/dd/yyyy")
                            Else
                                slVefName = slVefName & ", " & Trim$(tgMVef(ilVef).sName) & ":" & Format(tgSel(ilLoop).lStartDate, "mm/dd/yyyy") & "-" & Format(tgSel(ilLoop).lEndDate, "mm/dd/yyyy")
                            End If
                        End If
                    End If
                End If
            End If
        Next ilLoop
        If slVefName <> "" Then
            ilRet = MsgBox("Warning: Weeks previously posted in Affiliate System. See LogAffiliatePosted in Messages folder for list of Posted Vehicle. Press Cancel and select Internal Reprint or press Ok to continue reprinting posted week", vbQuestion + vbOKCancel, "Continue Log Generation?")
            If ilRet = vbCancel Then
                Screen.MousePointer = vbDefault
                cmcCancel.SetFocus
                Exit Sub
            End If
            gLogMsg "    Vehicles: " & slVefName, "LogAffiliatePosted.Txt", False
            gLogMsg "User re-generated the Logs for posted vehicle/weeks shown above", "LogAffiliatePosted.Txt", False
        End If
    End If
    
    bgReprintLogType = False
    If (rbcLogType(2).Value) Then
        bgReprintLogType = True
    End If
    'Convert Selling to Airing alerts as Logs are by Airing and Alerts entered by Selling
    gAlertVehicleReplace hmVLF
    'Check if this would remove alert- if so ask message
    ilFound = False
    If (rbcLogType(2).Value) Then
        For ilLoop = 0 To UBound(tgSel) - 1 Step 1
            If tgSel(ilLoop).iChk = 1 Then
                ilRet = gBinarySearchVef(tgSel(ilLoop).iVefCode)
                If ilRet <> -1 Then
                    tmVef = tgMVef(ilRet)
                    slStartDate = Format$(tgSel(ilLoop).lStartDate, "m/d/yy")
                    If tmVef.sType = "L" Then
                        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            '7/27/12: Include Sports within Log vehicles
                            'If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                            If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G")) And (tgMVef(ilVef).iVefCode = tmVef.iCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                                If gAlertFound("L", "S", 0, tgMVef(ilVef).iVefCode, slStartDate) Then
                                    ilFound = True
                                    Exit For
                                Else
                                    If gAlertFound("L", "C", 0, tgMVef(ilVef).iVefCode, slStartDate) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                End If
                            End If
                        Next ilVef
                        If ilFound Then
                            Exit For
                        End If
                    ElseIf tmVef.sType = "A" Then
                        'For ilLink = LBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) To UBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) Step 1
                        '    If tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink) > 0 Then
                        gBuildLinkArray hmVLF, tmVef, slStartDate, igSVefCode()
                        For ilLink = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                            If gAlertFound("L", "S", 0, igSVefCode(ilLink), slStartDate) Then
                                ilFound = True
                                Exit For
                            Else
                                If gAlertFound("L", "C", 0, igSVefCode(ilLink), slStartDate) Then
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        Next ilLink
                    Else
                        If gAlertFound("L", "S", 0, tmVef.iCode, slStartDate) Then
                            ilFound = True
                            Exit For
                        Else
                            If gAlertFound("L", "C", 0, tmVef.iCode, slStartDate) Then
                                ilFound = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next ilLoop
    End If
    If ilFound Or (rbcLogType(3).Value) Then
        ilRet = MsgBox("Have you verified that all spot and copy changes have been completed for the selected Logs", vbYesNo + vbQuestion, "All Changes Made")
        If ilRet = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    If (rbcLogType(1).Value) Or (rbcLogType(2).Value) Or (rbcLogType(3).Value) Then
        For ilLoop = 0 To UBound(tgSel) - 1 Step 1
            If tgSel(ilLoop).iChk = 1 Then
                'tmVefSrchKey.iCode = tgSel(ilLoop).iVefCode
                'ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                'TTP 10496 - Affiliate alerts created when log is generated even if there's no spots
                sgLogStartDate = Format$(tgSel(ilLoop).lStartDate, "m/d/yy")
                sgLogEndDate = Format$(tgSel(ilLoop).lEndDate, "m/d/yy")
                ilRet = gBinarySearchVef(tgSel(ilLoop).iVefCode)
                If ilRet <> -1 Then
                    tmVef = tgMVef(ilRet)
                    slStartDate = Format$(tgSel(ilLoop).lStartDate, "m/d/yy")
                    If tmVef.sType = "L" Then
                        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            '7/27/12: Include Sports within Log vehicles
                            'If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                            If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G")) And (tgMVef(ilVef).iVefCode = tmVef.iCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                                'tmVpfSrchKey.iVefKCode = tgMVef(ilVef).iCode
                                'ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                ilRet = gBinarySearchVpfPlus(tgMVef(ilVef).iCode)
                                If ilRet <> -1 Then
                                    tmVpf = tgVpf(ilRet)
                                    If (tmVpf.sMoveLLD = "Y") And (tgSpf.sCmmlSchStatus = "A") Then
                                        gUnpackDate tmVpf.iLLD(0), tmVpf.iLLD(1), slDate
                                        If gDateValue(slStartDate) <= gDateValue(slDate) Then
                                            ilDeleteSTF = MsgBox("**** Delete Commercial Changes ****", vbYesNoCancel + vbQuestion, "Spot Tracking File")
                                            If ilDeleteSTF = vbCancel Then
                                                Screen.MousePointer = vbDefault
                                                Exit Sub
                                            End If
                                            Exit For
                                        End If
                                    Else
                                        ilDeleteSTF = vbNo
                                    End If
                                End If
                            End If
                        Next ilVef
                        If ilDeleteSTF <> -1 Then
                            Exit For
                        End If
                    ElseIf tmVef.sType = "A" Then
                        'For ilLink = LBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) To UBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) Step 1
                        '    If tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink) > 0 Then
                        gBuildLinkArray hmVLF, tmVef, slStartDate, igSVefCode()
                        For ilLink = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                                tmVpfSrchKey.iVefKCode = igSVefCode(ilLink) 'tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink)
                                ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If (tmVpf.sMoveLLD = "Y") And (tgSpf.sCmmlSchStatus = "A") Then
                                    gUnpackDate tmVpf.iLLD(0), tmVpf.iLLD(1), slDate
                                    If gDateValue(slStartDate) <= gDateValue(slDate) Then
                                        ilDeleteSTF = MsgBox("**** Delete Commercial Changes ****", vbYesNoCancel + vbQuestion, "Spot Tracking File")
                                        If ilDeleteSTF = vbCancel Then
                                            Screen.MousePointer = vbDefault
                                            Exit Sub
                                        End If
                                        Exit For
                                    End If
                                Else
                                    ilDeleteSTF = vbNo
                                End If
                        '    End If
                        Next ilLink
                        If ilDeleteSTF <> -1 Then
                            Exit For
                        End If
                    Else
                        tmVpfSrchKey.iVefKCode = tgSel(ilLoop).iVefCode
                        ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If (tmVpf.sMoveLLD = "Y") And (tgSpf.sCmmlSchStatus = "A") Then
                            gUnpackDate tmVpf.iLLD(0), tmVpf.iLLD(1), slDate
                            If gDateValue(slStartDate) <= gDateValue(slDate) Then
                                ilDeleteSTF = MsgBox("**** Delete Commercial Changes ****", vbYesNoCancel + vbQuestion, "Spot Tracking File")
                                If ilDeleteSTF = vbCancel Then
                                    Screen.MousePointer = vbDefault
                                    Exit Sub
                                End If
                                Exit For
                            End If
                        Else
                            ilDeleteSTF = vbNo
                        End If
                    End If
                End If
            End If
        Next ilLoop
    End If
    'Warning message for copy
    If (ckcAssignCopy.Value = vbChecked) And (Not rbcLogType(4).Value) Then
        For ilLoop = 0 To UBound(tgSel) - 1 Step 1
            If tgSel(ilLoop).iChk = 1 Then
                slStartDate = Format$(tgSel(ilLoop).lStartDate, "m/d/yy")
                'TTP 10496 - Affiliate alerts created when log is generated even if there's no spots
                sgLogStartDate = slStartDate
                sgLogEndDate = Format(tgSel(ilLoop).lEndDate, "mm/dd/yyyy")
                If ((Asc(tgSpf.sUsingFeatures4) And ALLOWMOVEONTODAY) = ALLOWMOVEONTODAY) Then
                    If (gDateValue(slStartDate) < gDateValue(smNowDate)) Then
                        Screen.MousePointer = vbDefault
                        Beep
                        ilRet = MsgBox("Copy will only be assigned starting on Today's Date, use Assign Copy in Copy to Assign in the Past", vbOKOnly + vbExclamation, "Copy Assign Date")
                        Exit For
                    End If
                Else
                    If (gDateValue(slStartDate) <= gDateValue(smNowDate)) Then
                        Screen.MousePointer = vbDefault
                        Beep
                        ilRet = MsgBox("Copy will only be assigned starting after Today's Date, use Assign Copy in Copy to Assign in the Past", vbOKOnly + vbExclamation, "Copy Assign Date")
                        Exit For
                    End If
                End If
            End If
        Next ilLoop
    End If
    Screen.MousePointer = vbHourglass
    imGeneratingLog = True
    lbcLogMsg.Clear
    If tgSpf.sCBlackoutLog = "Y" Then
        If Not mOpenMsgFile() Then
            Screen.MousePointer = vbDefault
            cmcCancel.SetFocus
            Exit Sub
        End If
        'Determine earliest start date
        llStartDate = 99999999
        For ilLoop = 0 To UBound(tgSel) - 1 Step 1
            If tgSel(ilLoop).iChk = 1 Then
                slStartDate = Format$(tgSel(ilLoop).lStartDate, "m/d/yy")
                If tgSel(ilLoop).lStartDate < llStartDate Then
                    llStartDate = tgSel(ilLoop).lStartDate
                End If
            End If
        Next ilLoop
        slStartDate = Format$(llStartDate, "m/d/yy")
        ilRet = gReadBofRec(1, hmBof, hmCif, hmPrf, hmSif, hmCHF, "B", slStartDate, 1)
        igStartBofIndex = LBound(tgRBofRec) - 1
        ig30StartBofIndex = igStartBofIndex
        ig60StartBofIndex = igStartBofIndex
    End If
    'If imLogType < 2 Then
    '    For ilLoop = 0 To UBound(tgSel) Step 1
    '        tgLogSel(ilLoop) = tgSel(ilLoop)
    '    Next ilLoop
    'Else
    '    For ilLoop = 0 To UBound(tgSel) Step 1
    '        tgRPSel(ilLoop) = tgSel(ilLoop)
    '    Next ilLoop
    'End If
    ilRet = mGenLog(ilDeleteSTF)
    '11/26/17
    gFileChgdUpdate "vpf.btr", False
    If Not rbcLogType(4).Value Then
        gFileChgdUpdate "cptt.mkd", True
    End If
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        If tgSpf.sCBlackoutLog = "Y" Then
            Print #hmMsg, "Error during Log Generation: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
            Close #hmMsg
        End If
        imGeneratingLog = False
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        imTerminate = False
        Exit Sub
    End If
    imGeneratingLog = False
    If imTerminate Then
        If tgSpf.sCBlackoutLog = "Y" Then
            Print #hmMsg, "Log Generation Terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
            Close #hmMsg
        End If
        Screen.MousePointer = vbDefault
        imTerminate = False
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    If tgSpf.sCBlackoutLog = "Y" Then
        Print #hmMsg, "Log Generation Completed: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        Close #hmMsg
    End If
    If (lbcLogMsg.ListCount > 0) Then
        plcLogMsg.Visible = True
    End If
    If (imLogType = 1) Then
        pbcLogs(imPbcIndex).Cls
        mLogPop
        If imRPGen Then
            imRPGen = False
            mRPPop
        End If
        If imAlertGen Then
            imAlertGen = False
            mAlertPop
        End If
        pbcLogs_Paint imPbcIndex
        If ckcCheckOn.Value <> vbChecked Then
            ckcCheckOn.Value = vbChecked
        End If
    End If
    If (imLogType = 2) Then
        If imAlertGen Then
            imAlertGen = False
            mAlertPop
        End If
    End If
    If (imLogType = 3) Then
        pbcLogs(imPbcIndex).Cls
        'mLogPop
        'If imRPGen Then
        '    mRPPop
        'End If
        If imAlertGen Then
            imAlertGen = False
            mAlertPop
        End If
        ReDim tgSel(0 To UBound(tgAlertSel)) As LOGSEL
        For ilLoop = 0 To UBound(tgAlertSel) Step 1
            tgSel(ilLoop) = tgAlertSel(ilLoop)
        Next ilLoop
        If ckcCheckOn.Value <> vbChecked Then
            ckcCheckOn.Value = vbChecked
        End If
        imSettingValue = True
        vbcLogs.Value = vbcLogs.Min
        imSettingValue = True
        If UBound(tgSel) <= vbcLogs.LargeChange + 1 Then
            vbcLogs.Max = vbcLogs.Min
        Else
            vbcLogs.Max = UBound(tgSel) - vbcLogs.LargeChange + 1   'Show one extra line
        End If
        pbcLogs_Paint imPbcIndex
    End If
    'igRptCallType = LOGSJOB
    'igRptType = imSave(1)
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
    '    If igTestSystem Then
    '        slStr = "Logs^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode)) & "\" & smSave(2) & "\" & smSave(3) & "\" & smSave(4) & "\" & smSave(5) & "\" & Trim$(Str$(ilNoCodes)) & slCodeStr
    '    Else
    '        slStr = "Logs^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode)) & "\" & smSave(2) & "\" & smSave(3) & "\" & smSave(4) & "\" & smSave(5) & "\" & Trim$(Str$(ilNoCodes)) & slCodeStr
    '    End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Logs^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode)) & "\" & smSave(2) & "\" & smSave(3) & "\" & smSave(4) & "\" & smSave(5) & "\" & Trim$(Str$(ilNoCodes)) & slCodeStr
    '    Else
    '        slStr = "Logs^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode)) & "\" & smSave(2) & "\" & smSave(3) & "\" & smSave(4) & "\" & smSave(5) & "\" & Trim$(Str$(ilNoCodes)) & slCodeStr
    '    End If
    'End If
    'If imSave(1) = 0 Then
    '    ilShell = Shell(sgExePath & "RptSel.Exe " & slStr & "\||" & Trim$(Str$(LOGSJOB)) & "\Log", 1)
    'ElseIf imSave(1) = 1 Then
    '    ilShell = Shell(sgExePath & "RptSel.Exe " & slStr & "\||" & Trim$(Str$(LOGSJOB)) & "\Commercial Schedule", 1)
    'Else
    '    ilShell = Shell(sgExePath & "RptSel.Exe " & slStr & "\||" & Trim$(Str$(LOGSJOB)) & "\Commercial Summary", 1)
    'End If
    'Screen.MousePointer = vbDefault  'Wait
    'While GetModuleUsage(ilShell) > 0
    '    ilRet = DoEvents()
    'Wend
    'Screen.MousePointer = vbDefault
End Sub
Private Sub cmcGenerate_GotFocus()

    If imFirstTime Then
        imFirstTime = False
    End If
    plcTme.Visible = False
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
    'mSetShow imBoxNo  'Process last field if user moused here
    'imBoxNo = -1
    'gCtrlGotFocus ActiveControl
    'For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
    '    If mTestFields(ilLoop, ALLMANDEFINED + SHOWMSG) = NO Then
    '        Beep
    '        imBoxNo = ilLoop
    '        mEnableBox imBoxNo
    '        Exit Sub
    '    End If
    'Next ilLoop
    '
    'clTime1 = gTimeToCurrency(smSave(4), False)
    'clTime2 = gTimeToCurrency(smSave(5), True)
    'If (clTime2 < clTime1) Then
    '    Beep
    '    MsgBox "End Time earlier than Start Time.", 48, "Invalid Time Value"
    '    imBoxNo = ENDTIMEINDEX
    '    mEnableBox imBoxNo
    '    Exit Sub
    'End If


    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcLogChk_Click()
    Dim ilLoop As Integer
    If imLogType < 2 Then
        For ilLoop = 0 To UBound(tgSel) Step 1
            tgLogSel(ilLoop) = tgSel(ilLoop)
        Next ilLoop
        ReDim tgChkMsg(0 To UBound(tgLogChkMsg)) As LOGSEL
        For ilLoop = 0 To UBound(tgLogChkMsg) Step 1
            tgChkMsg(ilLoop) = tgLogChkMsg(ilLoop)
        Next ilLoop
    Else
        ReDim tgChkMsg(0 To 0) As LOGSEL
    End If

    LogChk.Show vbModal
End Sub
Private Sub cmcLogChk_GotFocus()
    plcTme.Visible = False
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcLogMsgOk_Click()
    plcLogMsg.Visible = False
End Sub

Private Sub cmcSplitFill_Click()
    Dim slStr As String
    Dim ilRet As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If igTestSystem Then
        slStr = "Logs^Test\" & sgUserName
    Else
        slStr = "Logs^Prod\" & sgUserName
    End If
    slStr = slStr & "\SplitFill"
    sgCommandStr = slStr
    Blackout.Show vbModal
    slStr = sgDoneMsg
    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        ilRet = mObtainSplitReplacments()
    End If
End Sub

Private Sub cmcTime_Click(Index As Integer)
    Select Case Index
        Case 0
            plcTme.Visible = Not plcTme.Visible
        Case 1
            plcTme.Visible = Not plcTme.Visible
    End Select
    edcTime(Index).SelStart = 0
    edcTime(Index).SelLength = Len(edcTime(Index).Text)
    edcTime(Index).SetFocus
End Sub
Private Sub cmcTime_GotFocus(Index As Integer)
    If imTmeIndex <> Index Then
        plcTme.Visible = False
        If Index = 0 Then
            plcTme.Move plcLogInfo.Left + edcTime(0).Left, plcLogInfo.Top + edcTime(0).Top + edcTime(0).height ' - plcTme.Height
        Else
            plcTme.Move plcLogInfo.Left + edcTime(1).Left, plcLogInfo.Top + edcTime(1).Top + edcTime(1).height ' - plcTme.Height
        End If
        imTmeIndex = Index
    End If
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_Change()
    Dim slStr As String
    Select Case imBoxNo
        Case CHKINDEX
        Case LLDINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                Exit Sub
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case LEADTIMEINDEX
        Case CYCLEINDEX
        Case SDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                Exit Sub
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case LOGINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcLog, imBSMode, imComboBoxIndex
        Case CPINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcCP, imBSMode, imComboBoxIndex
        Case LOGOINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcLogo, imBSMode, imComboBoxIndex
        Case OTHERINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcOther, imBSMode, imComboBoxIndex
        Case ZONEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcTimeZ, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
End Sub
Private Sub edcDropDown_GotFocus()
    If imFirstTime Then
        imFirstTime = False
    End If
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNo
        Case CHKINDEX
        Case LLDINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case LEADTIMEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case CYCLEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case SDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case LOGINDEX
        Case CPINDEX
        Case LOGOINDEX
        Case OTHERINDEX
        Case ZONEINDEX
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case CHKINDEX
            Case LLDINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case LEADTIMEINDEX
            Case CYCLEINDEX
            Case SDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case LOGINDEX
                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                    gProcessArrowKey Shift, KeyCode, lbcLog, imLbcArrowSetting
                End If
            Case CPINDEX
                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                    gProcessArrowKey Shift, KeyCode, lbcCP, imLbcArrowSetting
                End If
            Case LOGOINDEX
                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                    gProcessArrowKey Shift, KeyCode, lbcLogo, imLbcArrowSetting
                End If
            Case OTHERINDEX
                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                    gProcessArrowKey Shift, KeyCode, lbcOther, imLbcArrowSetting
                End If
            Case ZONEINDEX
                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                    gProcessArrowKey Shift, KeyCode, lbcTimeZ, imLbcArrowSetting
                End If
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
            Case CHKINDEX
            Case LLDINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case LEADTIMEINDEX
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case CYCLEINDEX
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case SDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case LOGINDEX
            Case CPINDEX
            Case LOGOINDEX
            Case OTHERINDEX
            Case ZONEINDEX
        End Select
    End If
End Sub
Private Sub edcDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Dim ilRnf As Integer
    'imButton = Button
    'If Button = 2 Then  'Right Mouse
    '    edcInfo.Visible = False
    '    Select Case imBoxNo
    '        Case CHKINDEX
    '        Case LEADTIMEINDEX
    '        Case CYCLEINDEX
    '        Case SDATEINDEX
    '        Case LOGINDEX
    '            imButtonIndex = lbcLog.ListIndex
    '            If (imButtonIndex > 0) And (imButtonIndex <= lbcLog.ListCount - 1) Then
    '                imIgnoreRightMove = True
    '                Screen.MousePointer = vbHourGlass
    '                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
    '                    If lbcLog.List(imButtonIndex) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
    '                        edcInfo.Text = Left$(tmRnfList(ilRnf).tRnf.sDescription, tmRnfList(ilRnf).tRnf.iStrLen)
    '                        edcInfo.Visible = True
    '                        On Error GoTo edcDropDownMDErr:
    '                        pbcRptSample(1).Picture = LoadPicture(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample)
    '                        pbcRptSample(0).Visible = True
    '                        Exit For
    '                    End If
    '                Next ilRnf
    '                imIgnoreRightMove = False
    '                Screen.MousePointer = vbDefault
    '            Else
    '                edcInfo.Visible = False
    '                pbcRptSample(0).Visible = False
    '            End If
    '        Case CPINDEX
    '            imButtonIndex = lbcCP.ListIndex
    '            If (imButtonIndex > 0) And (imButtonIndex <= lbcCP.ListCount - 1) Then
    '                imIgnoreRightMove = True
    '                Screen.MousePointer = vbHourGlass
    '                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
    '                    If lbcCP.List(imButtonIndex) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
    '                        edcInfo.Text = Left$(tmRnfList(ilRnf).tRnf.sDescription, tmRnfList(ilRnf).tRnf.iStrLen)
    '                        edcInfo.Visible = True
    '                        On Error GoTo edcDropDownMDErr:
    '                        pbcRptSample(1).Picture = LoadPicture(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample)
    '                        pbcRptSample(0).Visible = True
    '                        Exit For
    '                    End If
    '                Next ilRnf
    '                imIgnoreRightMove = False
    '                Screen.MousePointer = vbDefault
    '            Else
    '                edcInfo.Visible = False
    '                pbcRptSample(0).Visible = False
    '            End If
    '        Case LOGOINDEX
    '            imButtonIndex = lbcLogo.ListIndex
    '            If (imButtonIndex > 0) And (imButtonIndex <= lbcLogo.ListCount - 1) Then
    '                imIgnoreRightMove = True
    '                Screen.MousePointer = vbHourGlass
    '                On Error GoTo edcDropDownMDErr:
    '                pbcRptSample(1).Picture = LoadPicture(sgLogoPath & lbcLogo.List(imButtonIndex) & ".Bmp")
    '                pbcRptSample(0).Visible = True
    '                imIgnoreRightMove = False
    '                Screen.MousePointer = vbDefault
    '            Else
    '                pbcRptSample(0).Visible = False
    '            End If
    '        Case OTHERINDEX
    '            imButtonIndex = lbcOther.ListIndex
    '            If (imButtonIndex > 0) And (imButtonIndex <= lbcOther.ListCount - 1) Then
    '                imIgnoreRightMove = True
    '                Screen.MousePointer = vbHourGlass
    '                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
    '                    If lbcOther.List(imButtonIndex) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
    '                        edcInfo.Text = Left$(tmRnfList(ilRnf).tRnf.sDescription, tmRnfList(ilRnf).tRnf.iStrLen)
    '                        edcInfo.Visible = True
    '                        On Error GoTo edcDropDownMDErr:
    '                        pbcRptSample(1).Picture = LoadPicture(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample)
    '                        pbcRptSample(0).Visible = True
    '                        Exit For
    '                    End If
    '                Next ilRnf
    '                imIgnoreRightMove = False
    '                Screen.MousePointer = vbDefault
    '            Else
    '                edcInfo.Visible = False
    '                pbcRptSample(0).Visible = False
    '            End If
    '        Case ZONEINDEX
    '    End Select
    'End If
    'Exit Sub
'edcDropDownMDErr:
    'pbcRptSample(1).Picture = LoadPicture()
    'Resume Next
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 2 Then
    '    edcInfo.Visible = False
    '    pbcRptSample(0).Visible = False
    'End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcTime_GotFocus(Index As Integer)
    If imTmeIndex <> Index Then
        plcTme.Visible = False
    End If
    If Not imBypassFocus Then
        If Index = 0 Then
            plcTme.Move plcLogInfo.Left + edcTime(0).Left, plcLogInfo.Top + edcTime(0).Top + edcTime(0).height ' - plcTme.Height
        Else
            plcTme.Move plcLogInfo.Left + edcTime(1).Left, plcLogInfo.Top + edcTime(1).Top + edcTime(1).height ' - plcTme.Height
        End If
        gCtrlGotFocus ActiveControl
    End If
    imTmeIndex = Index
    imBypassFocus = False
End Sub
Private Sub edcTime_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTime_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKey As Integer
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
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If ActiveControl.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub edcTime_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case Index
            Case 0
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
                edcTime(Index).SelStart = 0
                edcTime(Index).SelLength = Len(edcTime(Index).Text)
            Case 1
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
                edcTime(Index).SelStart = 0
                edcTime(Index).SelLength = Len(edcTime(Index).Text)
        End Select
    End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True  'To get Alt J and Alt L keys
        Exit Sub
    End If
    imFirstActivate = False
    DoEvents    'Process events so pending keys are not sent to this
                'form when keypreview turn on
    'Logs.KeyPreview = True  'To get Alt J and Alt L keys
    pbcSelections.Enabled = True
    pbcSTab.Enabled = True
    pbcTab.Enabled = True
    If (igWinStatus(LOGSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
        cmcBlackout.Enabled = False
    Else
        imUpdateAllowed = True
    End If
    pbcLogs_Paint imPbcIndex
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Me.ZOrder 0 'Send to front
    Logs.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub Form_Deactivate()
    'Logs.KeyPreview = False
    Me.KeyPreview = False
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
    End If
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100) / Me.height
        Me.height = (lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100
    End If
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    rst_cptt.Close
    
    Erase tmTeamCode
    Erase tmRBofRec
    Erase tmSplitNetLastFill

    Erase tgLSTUpdateInfo
    Erase tmAdvertiser
    Erase tgSpotSum
    Erase tgOdfSdfCodes
    Erase tgRBofRec
    Erase tgSel
    Erase tgLogSel
    Erase tgRPSel
    Erase tmLogGen
    Erase igSVefCode
    Erase tmRnfList
    Erase tmTeam
    Erase tmLang
    Erase tgLogExportLoc
    'Close btrieve files
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmRnf)
    btrDestroy hmRnf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmStf)
    btrDestroy hmStf
    ilRet = btrClose(hmVLF)
    btrDestroy hmVLF
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    If tgSpf.sGUseAffSys = "Y" Then
        ilRet = btrClose(hmAtt)
        btrDestroy hmAtt
        ilRet = btrClose(hmCPTT)
        btrDestroy hmCPTT
        ilRet = btrClose(hmSHTT)
        btrDestroy hmSHTT
        ilRet = btrClose(hmLst)
        btrDestroy hmLst
        ilRet = btrClose(hmAbf)
        btrDestroy hmAbf
        ilRet = btrClose(hmSef)
        btrDestroy hmSef
        ilRet = btrClose(hmAxf)
        btrDestroy hmAxf
    End If
    If (tgSpf.sCBlackoutLog = "Y") Or ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        'If tgSpf.sGUseAffSys <> "Y" Then
        '    ilRet = btrClose(hmMcf)
        '    btrDestroy hmMcf
        'End If
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        ilRet = btrClose(hmBof)
        btrDestroy hmBof
        ilRet = btrClose(hmPrf)
        btrDestroy hmPrf
        ilRet = btrClose(hmSif)
        btrDestroy hmSif
        ilRet = btrClose(hmCrf)
        btrDestroy hmCrf
        ilRet = btrClose(hmCnf)
        btrDestroy hmCnf
        ilRet = btrClose(hmRsf)
        btrDestroy hmRsf
    End If
    'Delete arrays
    Erase tmCDCtrls
    Erase tmCtrls
    Erase tmVATT
    Erase tmLogVATT
    Erase tmCPTTInfo
    igJobShowing(LOGSJOB) = False
    
    Set Logs = Nothing
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub



Private Sub imcKey_Click()
    lbcKey.Visible = Not lbcKey.Visible
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lbcKey.Visible = True
    lbcKey.ZOrder
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lbcKey.Visible = False
End Sub

Private Sub lbcCP_Click()
    gProcessLbcClick lbcCP, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcCP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRnf As Integer
    imCurrentIndex = Y \ fgListHtArial825 'fgListHtSerif825
    imButton = Button
    If Button = 2 Then  'Right Mouse
        edcInfo.Visible = False
        pbcRptSample(0).Visible = False
        imButtonIndex = imCurrentIndex + lbcCP.TopIndex
        If (imButtonIndex > 0) And (imButtonIndex <= lbcCP.ListCount - 1) Then
            imIgnoreRightMove = True
            Screen.MousePointer = vbHourglass
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If lbcCP.List(imButtonIndex) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                    'edcInfo.Text = Left$(tmRnfList(ilRnf).tRnf.sDescription, tmRnfList(ilRnf).tRnf.iStrLen)
                    edcInfo.Text = gStripChr0(tmRnfList(ilRnf).tRnf.sDescription)
                    edcInfo.Visible = True
                    On Error GoTo lbcMDCPErr:
                    If gFileExist(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample) = 0 Then
                        pbcRptSample(1).Picture = LoadPicture(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample)
                    End If
                    pbcRptSample(0).Visible = True
                    Exit For
                End If
            Next ilRnf
            imIgnoreRightMove = False
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
lbcMDCPErr:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub lbcCP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRnf As Integer
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        If (Y < 60) Or (Y >= lbcCP.height - 135) Then
            imButtonIndex = 0
            edcInfo.Visible = False
            pbcRptSample(0).Visible = False
            Exit Sub
        End If
        If (X < 0) Or (X > lbcCP.Width) Then
            imButtonIndex = 0
            edcInfo.Visible = False
            pbcRptSample(0).Visible = False
            Exit Sub
        End If
        'If imButtonIndex <> (Y \ fgListHtSerif825) + lbcCP.TopIndex Then
        If imButtonIndex <> (Y \ fgListHtArial825) + lbcCP.TopIndex Then
            imIgnoreRightMove = True
            Screen.MousePointer = vbHourglass
            'imButtonIndex = Y \ fgListHtSerif825 + lbcCP.TopIndex
            imButtonIndex = Y \ fgListHtArial825 + lbcCP.TopIndex
            If (imButtonIndex > 0) And (imButtonIndex <= lbcCP.ListCount - 1) Then
                edcInfo.Visible = False
                pbcRptSample(0).Visible = False
                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                    If lbcCP.List(imButtonIndex) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                        'edcInfo.Text = Left$(tmRnfList(ilRnf).tRnf.sDescription, tmRnfList(ilRnf).tRnf.iStrLen)
                        edcInfo.Text = gStripChr0(tmRnfList(ilRnf).tRnf.sDescription)
                        edcInfo.Visible = True
                        On Error GoTo lbcMMCPErr:
                        If gFileExist(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample) = 0 Then
                            pbcRptSample(1).Picture = LoadPicture(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample)
                            pbcRptSample(0).Visible = True
                        End If
                        Exit For
                    End If
                Next ilRnf
            Else
                edcInfo.Visible = False
                pbcRptSample(0).Visible = False
            End If
            imIgnoreRightMove = False
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
lbcMMCPErr:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub lbcCP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        edcInfo.Visible = False
        pbcRptSample(0).Visible = False
    End If
End Sub
Private Sub lbcLog_Click()
    gProcessLbcClick lbcLog, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRnf As Integer
    imCurrentIndex = Y \ fgListHtArial825 'fgListHtSerif825
    imButton = Button
    If Button = 2 Then  'Right Mouse
        edcInfo.Visible = False
        pbcRptSample(0).Visible = False
        imButtonIndex = imCurrentIndex + lbcLog.TopIndex
        If (imButtonIndex > 0) And (imButtonIndex <= lbcLog.ListCount - 1) Then
            imIgnoreRightMove = True
            Screen.MousePointer = vbHourglass
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If lbcLog.List(imButtonIndex) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                    'edcInfo.Text = Left$(tmRnfList(ilRnf).tRnf.sDescription, tmRnfList(ilRnf).tRnf.iStrLen)
                    edcInfo.Text = gStripChr0(tmRnfList(ilRnf).tRnf.sDescription)
                    edcInfo.Visible = True
                    On Error GoTo lbcMDLogErr:
                    If gFileExist(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample) = 0 Then
                        pbcRptSample(1).Picture = LoadPicture(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample)
                        pbcRptSample(0).Visible = True
                    End If
                    Exit For
                End If
            Next ilRnf
            imIgnoreRightMove = False
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
lbcMDLogErr:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub lbcLog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRnf As Integer
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        If (Y < 60) Or (Y >= lbcLog.height - 135) Then
            imButtonIndex = 0
            edcInfo.Visible = False
            pbcRptSample(0).Visible = False
            Exit Sub
        End If
        If (X < 0) Or (X > lbcLog.Width) Then
            imButtonIndex = 0
            edcInfo.Visible = False
            pbcRptSample(0).Visible = False
            Exit Sub
        End If
        'If imButtonIndex <> (Y \ fgListHtSerif825) + lbcLog.TopIndex Then
        If imButtonIndex <> (Y \ fgListHtArial825) + lbcLog.TopIndex Then
            imIgnoreRightMove = True
            Screen.MousePointer = vbHourglass
            'imButtonIndex = Y \ fgListHtSerif825 + lbcLog.TopIndex
            imButtonIndex = Y \ fgListHtArial825 + lbcLog.TopIndex
            If (imButtonIndex > 0) And (imButtonIndex <= lbcLog.ListCount - 1) Then
                edcInfo.Visible = False
                pbcRptSample(0).Visible = False
                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                    If lbcLog.List(imButtonIndex) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                        'edcInfo.Text = Left$(tmRnfList(ilRnf).tRnf.sDescription, tmRnfList(ilRnf).tRnf.iStrLen)
                        edcInfo.Text = gStripChr0(tmRnfList(ilRnf).tRnf.sDescription)
                        edcInfo.Visible = True
                        On Error GoTo lbcMMLogErr:
                        If gFileExist(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample) = 0 Then
                            pbcRptSample(1).Picture = LoadPicture(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample)
                            pbcRptSample(0).Visible = True
                        End If
                        Exit For
                    End If
                Next ilRnf
            Else
                edcInfo.Visible = False
                pbcRptSample(0).Visible = False
            End If
            imIgnoreRightMove = False
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
lbcMMLogErr:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub lbcLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        edcInfo.Visible = False
        pbcRptSample(0).Visible = False
    End If
End Sub
Private Sub lbcLogo_Click()
    gProcessLbcClick lbcLogo, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRet As Integer

    imCurrentIndex = Y \ fgListHtArial825 'fgListHtSerif825
    imButton = Button
    If Button = 2 Then  'Right Mouse
        pbcRptSample(0).Visible = False
        pbcRptSample(1).Move 0, 0
        imButtonIndex = imCurrentIndex + lbcLog.TopIndex
        If (imButtonIndex > 0) And (imButtonIndex <= lbcLogo.ListCount - 1) Then
            imIgnoreRightMove = True
            Screen.MousePointer = vbHourglass
            'On Error GoTo lbcMDLogoErr:
            'pbcRptSample(1).Picture = LoadPicture(sgLogoPath & lbcLogo.List(imButtonIndex) & ".Bmp")
            'On Error GoTo mRetryPicture:
            'ilRet = 0
            If gFileExist(sgLogoPath & lbcLogo.List(imButtonIndex) & ".jpg") = 0 Then
                pbcRptSample(1).Picture = LoadPicture(sgLogoPath & lbcLogo.List(imButtonIndex) & ".jpg")
            ElseIf gFileExist(sgLogoPath & lbcLogo.List(imButtonIndex) & ".gif") = 0 Then
                pbcRptSample(1).Picture = LoadPicture(sgLogoPath & lbcLogo.List(imButtonIndex) & ".gif")
            ElseIf gFileExist(sgLogoPath & lbcLogo.List(imButtonIndex) & ".Bmp") = 0 Then
                pbcRptSample(1).Picture = LoadPicture(sgLogoPath & lbcLogo.List(imButtonIndex) & ".Bmp")
            End If
            pbcRptSample(0).Visible = True
            imIgnoreRightMove = False
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
mRetryPicture:
    ilRet = 1
    Resume Next
mNoPicture:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next

    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub lbcLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRet As Integer

    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        If (Y < 60) Or (Y >= lbcLogo.height - 135) Then
            imButtonIndex = 0
            edcInfo.Visible = False
            pbcRptSample(0).Visible = False
            Exit Sub
        End If
        If (X < 0) Or (X > lbcLogo.Width) Then
            imButtonIndex = 0
            edcInfo.Visible = False
            pbcRptSample(0).Visible = False
            Exit Sub
        End If
        'If imButtonIndex <> (Y \ fgListHtSerif825) + lbcLogo.TopIndex Then
        If imButtonIndex <> (Y \ fgListHtArial825) + lbcLogo.TopIndex Then
            imIgnoreRightMove = True
            Screen.MousePointer = vbHourglass
            'imButtonIndex = Y \ fgListHtSerif825 + lbcLogo.TopIndex
            imButtonIndex = Y \ fgListHtArial825 + lbcLogo.TopIndex
            If (imButtonIndex > 0) And (imButtonIndex <= lbcLogo.ListCount - 1) Then
                pbcRptSample(0).Visible = False
                pbcRptSample(1).Move 0, 0
                'On Error GoTo lbcMMLogoErr:
                'pbcRptSample(1).Picture = LoadPicture(sgLogoPath & lbcLogo.List(imButtonIndex) & ".Bmp")
                'On Error GoTo mRetryPicture:
                'ilRet = 0
                If gFileExist(sgLogoPath & lbcLogo.List(imButtonIndex) & ".jpg") = 0 Then
                    pbcRptSample(1).Picture = LoadPicture(sgLogoPath & lbcLogo.List(imButtonIndex) & ".jpg")
                ElseIf gFileExist(sgLogoPath & lbcLogo.List(imButtonIndex) & ".gif") = 0 Then
                    pbcRptSample(1).Picture = LoadPicture(sgLogoPath & lbcLogo.List(imButtonIndex) & ".gif")
                ElseIf gFileExist(sgLogoPath & lbcLogo.List(imButtonIndex) & ".Bmp") = 0 Then
                    pbcRptSample(1).Picture = LoadPicture(sgLogoPath & lbcLogo.List(imButtonIndex) & ".Bmp")
                End If
                pbcRptSample(0).Visible = True
            Else
                pbcRptSample(0).Visible = False
            End If
            imIgnoreRightMove = False
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
mRetryPicture:
    ilRet = 1
    Resume Next
mNoPicture:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next

    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub lbcLogo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        pbcRptSample(0).Visible = False
    End If
End Sub
Private Sub lbcOther_Click()
    gProcessLbcClick lbcOther, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcOther_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRnf As Integer
    imCurrentIndex = Y \ fgListHtArial825 'fgListHtSerif825
    imButton = Button
    If Button = 2 Then  'Right Mouse
        edcInfo.Visible = False
        pbcRptSample(0).Visible = False
        imButtonIndex = imCurrentIndex + lbcOther.TopIndex
        If (imButtonIndex > 0) And (imButtonIndex <= lbcOther.ListCount - 1) Then
            imIgnoreRightMove = True
            Screen.MousePointer = vbHourglass
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If lbcOther.List(imButtonIndex) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                    'edcInfo.Text = Left$(tmRnfList(ilRnf).tRnf.sDescription, tmRnfList(ilRnf).tRnf.iStrLen)
                    edcInfo.Text = gStripChr0(tmRnfList(ilRnf).tRnf.sDescription)
                    edcInfo.Visible = True
                    On Error GoTo lbcMDPlayErr:
                    If gFileExist(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample) = 0 Then
                        pbcRptSample(1).Picture = LoadPicture(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample)
                        pbcRptSample(0).Visible = True
                    End If
                    Exit For
                End If
            Next ilRnf
            imIgnoreRightMove = False
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
lbcMDPlayErr:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub lbcOther_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRnf As Integer
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        If (Y < 60) Or (Y >= lbcOther.height - 135) Then
            imButtonIndex = 0
            edcInfo.Visible = False
            pbcRptSample(0).Visible = False
            Exit Sub
        End If
        If (X < 0) Or (X > lbcOther.Width) Then
            imButtonIndex = 0
            edcInfo.Visible = False
            pbcRptSample(0).Visible = False
            Exit Sub
        End If
        'If imButtonIndex <> (Y \ fgListHtSerif825) + lbcOther.TopIndex Then
        If imButtonIndex <> (Y \ fgListHtArial825) + lbcOther.TopIndex Then
            imIgnoreRightMove = True
            Screen.MousePointer = vbHourglass
            'imButtonIndex = Y \ fgListHtSerif825 + lbcOther.TopIndex
            imButtonIndex = Y \ fgListHtArial825 + lbcOther.TopIndex
            If (imButtonIndex > 0) And (imButtonIndex <= lbcOther.ListCount - 1) Then
                edcInfo.Visible = False
                pbcRptSample(0).Visible = False
                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                    If lbcOther.List(imButtonIndex) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                        'edcInfo.Text = Left$(tmRnfList(ilRnf).tRnf.sDescription, tmRnfList(ilRnf).tRnf.iStrLen)
                        edcInfo.Text = gStripChr0(tmRnfList(ilRnf).tRnf.sDescription)
                        edcInfo.Visible = True
                        On Error GoTo lbcMMPlayErr:
                        If gFileExist(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample) = 0 Then
                            pbcRptSample(1).Picture = LoadPicture(sgRptPath & tmRnfList(ilRnf).tRnf.sRptSample)
                            pbcRptSample(0).Visible = True
                        End If
                        Exit For
                    End If
                Next ilRnf
            Else
                edcInfo.Visible = False
                pbcRptSample(0).Visible = False
            End If
            imIgnoreRightMove = False
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
lbcMMPlayErr:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub lbcOther_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        edcInfo.Visible = False
        pbcRptSample(0).Visible = False
    End If
End Sub
Private Sub lbcTimeZ_Click()
    gProcessLbcClick lbcTimeZ, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcVehicle_Click()
    'Vehicles with same date can only be picked
    'If Not imChgMode Then
    '    imChgMode = True
    '    slCommDate = "|"
    '    ilNotMatching = False
    '    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
    '        If lbcVehicle.Selected(ilLoop) Then
    '            slNameDate = lbcVehicle.List(ilLoop)
    '            ilPos = InStr(slNameDate, "(L")
    '            If ilPos > 0 Then
    '                slNameDate = Left$(slNameDate, ilPos - 1)
    '            End If
    '            ilPos = 2
    '            ilRet = gParseItem(slNameDate, ilPos, " ", slDate)
    '            Do While ilRet = CP_MSG_NONE
    '                If gValidDate(slDate) Then
    '                    Exit Do
    '                End If
    '                ilPos = ilPos + 1
    '                ilRet = gParseItem(slNameDate, ilPos, " ", slDate)
    '            Loop
    '            If ilRet <> CP_MSG_NONE Then
    '                slDate = ""
    '            End If
    '            If slCommDate = "|" Then
    '                slCommDate = slDate
    '            Else
    '                If slCommDate = "" Then
    '                    If slDate <> "" Then
    '                        ilNotMatching = True
    '                        Exit For
    '                    End If
    '                Else
    '                    If slDate = "" Then
    '                        ilNotMatching = True
    '                        Exit For
    '                    Else
    '                        If gDateValue(slCommDate) <> gDateValue(slDate) Then
    '                            ilNotMatching = True
    '                            Exit For
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Else
    '            imVefSelected(ilLoop) = False
    '        End If
    '    Next ilLoop
    '    If ilNotMatching Then
    '        For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
    '            If lbcVehicle.Selected(ilLoop) Then
    '                If Not imVefSelected(ilLoop) Then
    '                    lbcVehicle.Selected(ilLoop) = False
    '                End If
    '            End If
    '        Next ilLoop
    '    Else
    '        For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
    '            imVefSelected(ilLoop) = lbcVehicle.Selected(ilLoop)
    '        Next ilLoop
    '    End If
    '    imChgMode = False
    '    mSetCommands
    'End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mATTPop                         *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:5/4/94       By:                *
'*                                                     *
'*            Comments: Populatiom tmVATT given tmVef  *
'*                                                     *
'*******************************************************
Private Sub mATTPop(ilVefCode As Integer, slLSTOnly As String)
    Dim ilRet As Integer
    Dim llUpperBound As Long
    Dim llNoRec As Long
    Dim ilExtLen As Integer
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim ilVpf As Integer

    slLSTOnly = "N"
    ReDim tmVATT(0 To 0) As ATT
    If Not imATTExist Then
        Exit Sub
    End If
    ilVpf = gBinarySearchVpf(ilVefCode)
    If ilVpf <> -1 Then
        'If (tgVpf(ilVpf).sWegenerExport = "Y") Or (tgVpf(ilVpf).sOLAExport = "Y") Then
        If (tgVpf(ilVpf).sOLAExport = "Y") Then
            slLSTOnly = "Y"
            Exit Sub
        End If
    End If
    llUpperBound = UBound(tmVATT)
    ilExtLen = Len(tmVATT(llUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAnf) 'Obtain number of records
    btrExtClear hmAtt   'Clear any previous extend operation
    tmATTSrchKey1.iCode = ilVefCode
    ilRet = btrGetEqual(hmAtt, tmAtt, imAttRecLen, tmATTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    Call btrExtSetBounds(hmAtt, llNoRec, -1, "UC", "ATT", "") 'Set extract limits (all records)
    tlIntTypeBuff.iType = ilVefCode
    'ilOffset = GetOffSetForInt(tmAtt, tmAtt.ivefCode)
    ilOffSet = gFieldOffset("ATT", "ATTVEFCODE")
    ilRet = btrExtAddLogicConst(hmAtt, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    ilOffSet = 0
    ilRet = btrExtAddField(hmAtt, ilOffSet, ilExtLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilRet = btrExtGetNext(hmAtt, tmVATT(llUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            Exit Sub
        End If
        llUpperBound = UBound(tmVATT)
        ilExtLen = Len(tmVATT(llUpperBound))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmAtt, tmVATT(llUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If (tmVATT(llUpperBound).iCarryCmml = 0) Then
                tmSHTTSrchKey.iCode = tmVATT(llUpperBound).iShfCode
                ilRet = btrGetEqual(hmSHTT, tmSHTT, imSHTTRecLen, tmSHTTSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) And (tmSHTT.iType = 0) Then
                    llUpperBound = llUpperBound + 1
                    ReDim Preserve tmVATT(0 To llUpperBound) As ATT
                End If
            End If
            ilRet = btrExtGetNext(hmAtt, tmVATT(llUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmAtt, tmVATT(llUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
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

    slStr = edcDropDown.Text
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
'*      Procedure Name:mBuildBlackouts                 *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan     *
'*                                                     *
'*            Comments: Build ODF and suppress/replace *
'*                      as required
'*                                                     *
'*******************************************************
Private Function mBuildBlackouts() As Integer
    Dim ilRet As Integer
    If tgSpf.sCBlackoutLog = "Y" Then
        imOdfRecLen = Len(tmOdf)  'Get and save ADF record length
        hmOdf = CBtrvTable(TEMPHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        hmCvf = CBtrvTable(TEMPHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        lgEndIndex = UBound(tgSpotSum)
        gBlackoutTest 1, hmCif, hmMcf, hmOdf, hmRsf, hmCpf, hmCrf, hmCnf, hmClf, hmLst, hmCvf, smNewLines(), hmMsg, lbcLogMsg
        ilRet = btrClose(hmCvf)
        btrDestroy hmCvf
        ilRet = btrClose(hmOdf)
        btrDestroy hmOdf
    End If
    mBuildBlackouts = True
    Exit Function
End Function
'************************************************************
'*                                                          *
'*      Procedure Name:mEnableBox                           *
'*                                                          *
'*             Created:5/17/93       By:D. LeVine           *
'*            Modified:5/4/94       By:D. Hannifan          *
'*                                                          *
'*            D.S. 1/31/02 Added code to allow 3 digit      *
'*                         cycle number if using affiliate  *
'*                         system and doing a reprint       *
'*                                                          *
'*            Comments: Enable specified control            *
'*                                                          *
'************************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer       'For loop control parameter
    Dim slStr As String         'Parse string
    Dim ilIndex As Integer
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    If (imRowNo < vbcLogs.Value) Or (imRowNo >= vbcLogs.Value + vbcLogs.LargeChange + 1) Then
        mSetShow ilBoxNo
        Exit Sub
    End If
    lacFrame(imPbcIndex).Move 0, tmCtrls(CHKINDEX).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15) - 30
    lacFrame(imPbcIndex).Visible = True
    pbcArrow.Move pbcArrow.Left, plcLogs.Top + tmCtrls(CHKINDEX).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True

    Select Case ilBoxNo 'Branch on box type (control)
        Case CHKINDEX
            If tgSel(imRowNo - 1).iChk < 0 Then
                tgSel(imRowNo - 1).iChk = 0
                tgSel(imRowNo - 1).iInitChk = 0
            End If
            pbcSelections.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveTableCtrl pbcLogs(imPbcIndex), pbcSelections, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15)
            pbcSelections_Paint
            pbcSelections.Visible = True
            pbcSelections.SetFocus
        Case LLDINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcLogs(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcLogs.Value <= vbcLogs.LargeChange \ 2 Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.height
            End If
            slStr = tgSel(imRowNo - 1).sLLD
            If Trim$(slStr) = "" Then
                If tgSel(imRowNo - 1).lStartDate > 0 Then
                    slStr = Format$(tgSel(imRowNo - 1).lStartDate - tgSel(imRowNo - 1).iCycle, "m/d/yy")
                Else
                    slStr = smDate
                End If
            End If
            pbcCalendar_Paint
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            plcCalendar.Visible = True
            edcDropDown.SetFocus
        Case LEADTIMEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 2
            gMoveTableCtrl pbcLogs(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15)
            If tgSel(imRowNo - 1).iLeadTime < 0 Then
                tgSel(imRowNo - 1).iLeadTime = 1
            End If
            edcDropDown.Text = Trim$(str$(tgSel(imRowNo - 1).iLeadTime))
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case CYCLEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            'D.S. 1/31/02
            'Allow 3 digit cycle number if doing a reprint and using the Affiliate System
            If ((rbcLogType(2).Value) Or (rbcLogType(4).Value)) And (tgSpf.sGUseAffSys = "Y") Then
                edcDropDown.MaxLength = 3
            Else
                edcDropDown.MaxLength = 2
            End If

            gMoveTableCtrl pbcLogs(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15)
            If tgSel(imRowNo - 1).iCycle < 0 Then
                tgSel(imRowNo - 1).iCycle = 1
            End If
            edcDropDown.Text = Trim$(str$(tgSel(imRowNo - 1).iCycle))
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case SDATEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcLogs(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcLogs.Value <= vbcLogs.LargeChange \ 2 Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.height
            End If
            If tgSel(imRowNo - 1).lStartDate <= 0 Then     'initialize
                slStr = ""
            Else
                slStr = Format$(tgSel(imRowNo - 1).lStartDate, "m/d/yy")
            End If
            pbcCalendar_Paint
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            plcCalendar.Visible = True
            edcDropDown.SetFocus
        Case LOGINDEX
            lbcLog.height = gListBoxHeight(lbcLog.ListCount, 8)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcLogs(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcLogs.Value <= vbcLogs.LargeChange \ 2 Then
                lbcLog.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                lbcLog.Move edcDropDown.Left, edcDropDown.Top - lbcLog.height
            End If
            imChgMode = True
            imComboBoxIndex = lbcLog.ListIndex  'save gotfocus list index if any
            If tgSel(imRowNo - 1).iLog < 0 Then
                lbcLog.ListIndex = 0
                tgSel(imRowNo - 1).iLog = 0
                imComboBoxIndex = lbcLog.ListIndex 'save gotfocus list index if above save was invalid
            Else
                lbcLog.ListIndex = tgSel(imRowNo - 1).iLog
                imComboBoxIndex = lbcLog.ListIndex 'save gotfocus list index if above save was invalid
            End If
            edcDropDown.Text = lbcLog.List(lbcLog.ListIndex)
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case CPINDEX
            lbcCP.height = gListBoxHeight(lbcCP.ListCount, 8)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcLogs(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcLogs.Value <= vbcLogs.LargeChange \ 2 Then
                lbcCP.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                lbcCP.Move edcDropDown.Left, edcDropDown.Top - lbcCP.height
            End If
            imChgMode = True
            imComboBoxIndex = lbcLog.ListIndex  'save gotfocus list index if any
            If tgSel(imRowNo - 1).iCP < 0 Then
                lbcCP.ListIndex = 0
                tgSel(imRowNo - 1).iCP = 0
                imComboBoxIndex = lbcCP.ListIndex 'save gotfocus list index if above save was invalid
            Else
                lbcCP.ListIndex = tgSel(imRowNo - 1).iCP
                imComboBoxIndex = lbcCP.ListIndex 'save gotfocus list index if above save was invalid
            End If
            edcDropDown.Text = lbcCP.List(lbcCP.ListIndex)
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case LOGOINDEX
            lbcLogo.height = gListBoxHeight(lbcLogo.ListCount, 8)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
            edcDropDown.MaxLength = 4
            gMoveTableCtrl pbcLogs(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcLogs.Value <= vbcLogs.LargeChange \ 2 Then
                lbcLogo.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                lbcLogo.Move edcDropDown.Left, edcDropDown.Top - lbcLogo.height
            End If
            imChgMode = True
            imComboBoxIndex = lbcLogo.ListIndex  'save gotfocus list index if any
            If tgSel(imRowNo - 1).iLogo < 0 Then
                lbcLogo.ListIndex = 0
                tgSel(imRowNo - 1).iLogo = 0
                imComboBoxIndex = lbcLogo.ListIndex 'save gotfocus list index if above save was invalid
            Else
                lbcLogo.ListIndex = tgSel(imRowNo - 1).iLogo
                imComboBoxIndex = lbcLogo.ListIndex 'save gotfocus list index if above save was invalid
            End If
            edcDropDown.Text = lbcLogo.List(lbcLogo.ListIndex)
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case OTHERINDEX
            lbcOther.height = gListBoxHeight(lbcOther.ListCount, 8)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcLogs(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcLogs.Value <= vbcLogs.LargeChange \ 2 Then
                lbcOther.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                lbcOther.Move edcDropDown.Left, edcDropDown.Top - lbcOther.height
            End If
            imChgMode = True
            imComboBoxIndex = lbcOther.ListIndex  'save gotfocus list index if any
            If tgSel(imRowNo - 1).iOther < 0 Then
                lbcOther.ListIndex = 0
                tgSel(imRowNo - 1).iOther = 0
                imComboBoxIndex = lbcOther.ListIndex 'save gotfocus list index if above save was invalid
            Else
                lbcOther.ListIndex = tgSel(imRowNo - 1).iOther
                imComboBoxIndex = lbcOther.ListIndex 'save gotfocus list index if above save was invalid
            End If
            edcDropDown.Text = lbcOther.List(lbcOther.ListIndex)
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ZONEINDEX
            lbcTimeZ.Clear
            For ilLoop = LBound(tgVpf(tgSel(imRowNo - 1).iVpfIndex).sGZone) To UBound(tgVpf(tgSel(imRowNo - 1).iVpfIndex).sGZone) Step 1
               If Trim$(tgVpf(tgSel(imRowNo - 1).iVpfIndex).sGZone(ilLoop)) <> "" Then
                  lbcTimeZ.AddItem Trim$(tgVpf(tgSel(imRowNo - 1).iVpfIndex).sGZone(ilLoop))
               End If
            Next ilLoop
            lbcTimeZ.AddItem "[All]", 0
            lbcTimeZ.height = gListBoxHeight(lbcTimeZ.ListCount, 5)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW ' - cmcDropDown.Width
            edcDropDown.MaxLength = 6
            gMoveTableCtrl pbcLogs(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX - cmcDropDown.Width, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcLogs.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcLogs.Value <= vbcLogs.LargeChange \ 2 Then
                lbcTimeZ.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcTimeZ.Width, edcDropDown.Top + edcDropDown.height
            Else
                lbcTimeZ.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcTimeZ.Width, edcDropDown.Top - lbcTimeZ.height
            End If
            imChgMode = True
            imComboBoxIndex = lbcTimeZ.ListIndex  'save gotfocus list index if any
            If tgSel(imRowNo - 1).iZone <= 0 Then
                lbcTimeZ.ListIndex = 0
                tgSel(imRowNo - 1).iZone = 0
                imComboBoxIndex = lbcTimeZ.ListIndex 'save gotfocus list index if above save was invalid
            Else
                Select Case tgSel(imRowNo - 1).iZone
                    Case 1  'EST
                        slStr = "EST"
                    Case 2  'CST
                        slStr = "CST"
                    Case 3
                        slStr = "MST"
                    Case 4
                        slStr = "PST"
                    Case Else
                        slStr = ""
                End Select
                ilIndex = 0
                For ilLoop = 0 To lbcTimeZ.ListCount - 1 Step 1
                    If StrComp(slStr, lbcTimeZ.List(ilLoop), 1) = 0 Then
                        ilIndex = ilLoop
                        Exit For
                    End If
                Next ilLoop
                lbcTimeZ.ListIndex = ilIndex
                imComboBoxIndex = lbcTimeZ.ListIndex 'save gotfocus list index if above save was invalid
            End If
            edcDropDown.Text = lbcTimeZ.List(lbcTimeZ.ListIndex)
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
    mSetCommands   'Check mandatory fields and set controls
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGenLog                         *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Generate ODF and set Last Log  *
'*                      Date                           *
'*                                                     *
'*      1-28-05 change pdf filename to use all 5 characters
'               of the vhicle station code, exceeding
'               8char filename
'*******************************************************
Private Function mGenLog(ilDeleteSTF As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilGen                         ilGameMergeFd             *
'*                                                                                        *
'******************************************************************************************

'
'   ilRet = mGenLog(ilNoCodes, slCodeStr)
'   Where:
'       ilNoCodes(O)- Number of codes stored into slCodeStr
'       slCodeStr(O)- Vehicle code selected
'
    Dim ilRet As Integer       'Call return value
    Dim ilCRet As Integer
    Dim ilType As Integer       'SSF type
    Dim sLCP As String         'Vehicle status
    Dim slStartDate As String  'Start Date String
    Dim slEndDate As String     'End Date String
    Dim slStartTime As String  'Start Time String
    Dim slEndTime As String    'End Time String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slMonStartDate As String  'Start Date String
    Dim llDate1 As Long        'Start date value
    Dim llDate2 As Long
    Dim slDate As String
    Dim ilVpfIndex As Integer
    Dim ilLoop As Integer
    Dim llLoop1 As Long
    Dim llLoop2 As Long
    Dim ilCycle As Integer
    Dim llCycleDate As Long
    Dim slCycleDate As String
    Dim ilLogAlertExisted As Integer
    Dim ilCopyAlertExisted As Integer
    Dim ilLink As Integer
    Dim llTestDate As Long
    Dim ilVef As Integer
    Dim tlVef As VEF
    Dim llCDate As Long 'Copy Last date assigned
    Dim slCopyStartDate As String
    Dim ilPass As Integer
    Dim ilEPass As Integer
    Dim ilLogGen As Integer
    Dim ilGenLST As Integer
    Dim ilZone As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    ReDim ilEvtAllowed(0 To 14) As Integer
    Dim slTZStartDate As String
    Dim slTZEndDate As String
    Dim slTZStartTime As String
    Dim slTZEndTime As String
    Dim ilZoneExist As Integer
    Dim llIndex As Long
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilProcVefCode As Integer
    Dim ilExportType As Integer
    Dim ilODFVefCode As Integer
    Dim llDate As Long
    Dim ilValue As Integer
    Dim slGameDate As String
    Dim tlSvSel As LOGSEL
    Dim slLSTOnly As String
    Dim ilCombineVpfIndex As Integer
    Dim ilLSTForLogVeh As Integer
    Dim llLstDate As Long
    Dim slGenVehName As String
    Dim ilLogPass As Integer
    Dim ilMergeExist As Integer
    Dim llSeasonStart As Long
    Dim llSeasonEnd As Long
    Dim blGenNTR As Boolean
    Dim ilRnf As Integer
    Dim ihARf As Integer
    Dim tlArf As ARF
    Dim tlARFSrchKey As INTKEY0
    Dim ilVff As Integer
    Dim slStr As String
    lgAssignTotal = 0
    
    mGenLog = True
    ilType = 0
    sLCP = "C"
    slStartTime = Trim$(edcTime(0).Text)
    slEndTime = Trim$(edcTime(1).Text)
    If ckcOutput(2).Value = vbChecked Then
        ilEPass = 5
    Else
        ilEPass = 2
    End If
    For ilLoop = LBound(ilEvtAllowed) To UBound(ilEvtAllowed) Step 1
        ilEvtAllowed(ilLoop) = True
    Next ilLoop
    ilEvtAllowed(0) = False 'Don't include library names
    DoEvents
    If imTerminate Then
        Exit Function
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    plcStatus.Cls
    DoEvents
    If imTerminate Then
        Exit Function
    End If

    ReDim tgAbfInfo(0 To 0) As ABFINFO

    'Create Billboard spots
    '4/18/14: Moved within loop
    'ilRet = mBBSpots()

    ReDim tgLogExportLoc(0 To 0) As LOGEXPORTLOC
    If ckcOutput(2).Value = vbChecked Then          'find all the vehicles that have export loc defined
        ihARf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(ihARf, "", sgDBPath & "Arf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        For ilLoop = 0 To UBound(tgSel)
            If tgSel(ilLoop).iChk = 1 Then
                ilVff = gBinarySearchVff(tgSel(ilLoop).iVefCode)
                If ilVff <> -1 Then
                    If tgVff(ilVff).iLogExptArfCode > 0 Then            'separate export folder for this log
                        tlARFSrchKey.iCode = tgVff(ilVff).iLogExptArfCode
                        ilRet = btrGetEqual(ihARf, tlArf, Len(tlArf), tlARFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            'tlArf.sftp contains the export path for this log
                            tgLogExportLoc(UBound(tgLogExportLoc)).sExportPath = Trim$(tlArf.sFTP)     'vehicle export folder
                            'ensure theres a backslash (\) at the end of path
                            If right$(Trim$(tgLogExportLoc(UBound(tgLogExportLoc)).sExportPath), 1) <> "\" Then
                                tgLogExportLoc(UBound(tgLogExportLoc)).sExportPath = Trim$(tgLogExportLoc(UBound(tgLogExportLoc)).sExportPath) & "\"
                            End If

                            tgLogExportLoc(UBound(tgLogExportLoc)).iVefCode = tgSel(ilLoop).iVefCode
                            slStr = Trim$(str$(tgSel(ilLoop).iVefCode))
                            Do While Len(slStr) < 5
                                slStr = "0" & slStr
                            Loop
                            tgLogExportLoc(UBound(tgLogExportLoc)).sKey = slStr         'key for sorting
                            ReDim Preserve tgLogExportLoc(0 To UBound(tgLogExportLoc) + 1) As LOGEXPORTLOC
                        End If
                    End If
                End If
            End If
        Next ilLoop
        ilRet = btrClose(ihARf)
        btrDestroy ihARf
        'sort the list of vehicles with export locations
         If UBound(tgLogExportLoc) - 1 > 0 Then
            ArraySortTyp fnAV(tgLogExportLoc(), 0), UBound(tgLogExportLoc), 0, LenB(tgLogExportLoc(0)), 0, LenB(tgLogExportLoc(0).sKey), 0
        End If
    End If
    
    ReDim tgLSTUpdateInfo(0 To 0) As LSTUPDATEINFO
    For ilLoop = 0 To UBound(tgSel) - 1 Step 1
        If tgSel(ilLoop).iChk = 1 Then
            '4/18/14: Create BB's just prior to gathering spots to min time if running report that removes BB spots
            ilRet = mBBSpots(ilLoop)
            ilType = 0
            sLCP = "C"
            tmVefSrchKey.iCode = tgSel(ilLoop).iVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            'TTP 10496 - Affiliate alerts created when log is generated even if there's no spots
            sgLogStartDate = Format$(tgSel(ilLoop).lStartDate, "m/d/yy")
            sgLogEndDate = Format$(tgSel(ilLoop).lEndDate, "m/d/yy")
            If bgLogFirstCallToVpfFind Then
                ilVpfIndex = gVpfFind(Logs, tmVef.iCode)
                bgLogFirstCallToVpfFind = False
            Else
                ilVpfIndex = gVpfFindIndex(tmVef.iCode)
            End If
            slGenVehName = Trim$(tmVef.sName)
            gUserActivityLog "S", slGenVehName & ": Log Generation"
            smStatusCaption = "Clearing Log File for " & Trim$(tmVef.sName)
            plcStatus.Cls
            plcStatus_Paint
            ilZoneExist = False
            For ilZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
                If Trim$(tgVpf(ilVpfIndex).sGZone(ilZone)) <> "" Then
                    ilZoneExist = True
                    Exit For
                End If
            Next ilZone
            ReDim tmLogGen(0 To 1) As LOGGEN
            tmLogGen(0).iGenVefCode = tmVef.iCode
            tmLogGen(0).iSimVefCode = tmVef.iCode
            ReDim tgSpotSum(0 To 0) As SPOTSUM
            ReDim tgOdfSdfCodes(0 To 0) As ODFSDFCODES
            lgStartIndex = UBound(tgSpotSum)
            'Returning to use btrDelete instead of copy odf_blk
            'Jim-3/26/01
            'If tgSpf.sGUseAffSys = "Y" Then
            '    'gDeleteOdf "G", slType, slCP, tgSel(ilLoop).iVefCode
            'Else
            '    ilRet = 0
            '    On Error GoTo mFileCopyErr
            '    If igRetrievalDB = 1 Then
            '        FileCopy sgSDBPath & "Odf_Blk.Btr", sgSDBPath & "Odf.Btr"
            '    Else
            '        If Len(sgMDBPath) <= 2 Then
            '            FileCopy sgDBPath & "Odf_Blk.Btr", sgDBPath & "Odf.Btr"
            '        Else
            '            FileCopy sgMDBPath & "Odf_Blk.Btr", sgMDBPath & "Odf.Btr"
            '        End If
            '    End If
            'End If
            sgGenDate = Format$(gNow(), "m/d/yy")
            sgGenTime = Format$(gNow(), "h:mm:ssAM/PM")
            gPackDate sgGenDate, igGenDate(0), igGenDate(1)
            gPackTime sgGenTime, igGenTime(0), igGenTime(1)
            '10-9-01
            gUnpackTimeLong igGenTime(0), igGenTime(1), False, lgGenTime
            On Error GoTo 0
            If ilRet <> 0 Then
                hmOdf = CBtrvTable(ONEHANDLE)          'Save VEF handle
                ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                On Error GoTo mGenLogErr
                gBtrvErrorMsg ilRet, "mGenLog (btrOpen)", Logs
                On Error GoTo 0
                ilRet = btrClear(hmOdf) 'Remove all records
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrClose(hmOdf)
                    btrDestroy hmOdf
                    gUserActivityLog "E", slGenVehName & ": Log Generation"
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Clearing Odf error" & str$(ilRet), vbOKOnly + vbCritical, "Error")
                    mGenLog = False
                    Exit Function
                End If
                ilRet = btrClose(hmOdf)
                btrDestroy hmOdf
            End If
            llStartDate = tgSel(ilLoop).lStartDate
            llEndDate = tgSel(ilLoop).lEndDate
            slStartDate = Format$(tgSel(ilLoop).lStartDate, "m/d/yy")
            slEndDate = Format$(tgSel(ilLoop).lEndDate, "m/d/yy")
            slMonStartDate = gObtainPrevMonday(slStartDate)
            If ((Asc(tgSpf.sUsingFeatures4) And ALLOWMOVEONTODAY) = ALLOWMOVEONTODAY) Then
                If (gDateValue(slStartDate) < gDateValue(smNowDate)) Then
                    slCopyStartDate = smNowDate
                Else
                    slCopyStartDate = slStartDate
                End If
            Else
                If (gDateValue(slStartDate) <= gDateValue(smNowDate)) Then
                    slCopyStartDate = smDate
                Else
                    slCopyStartDate = slStartDate
                End If
            End If
            slTZStartDate = Format$(tgSel(ilLoop).lEndDate + 1, "m/d/yy")
            slTZEndDate = slTZStartDate
            slTZStartTime = "12M"
            slTZEndTime = "3A"
            
            'If tgVpf(ilVpfIndex).sGMedium = "P" Then            'podcast vehicle, read all NTRs for this period for the log
            'Determine logs to gen the NTR items in memory
                blGenNTR = False
                If UCase(Trim$(lbcLog.List(tgSel(ilLoop).iLog))) = "L87" Then
                    blGenNTR = True
                End If

                If UCase(Trim$(lbcOther.List(tgSel(ilLoop).iOther))) = "L87" Then
                    blGenNTR = True
                End If
                    
                ilRet = gGetNTRForLog(Logs, ilVpfIndex, slStartDate, slEndDate, blGenNTR)
                If Not ilRet Then
                    imTerminate = True
                    mGenLog = False
                    Exit Function
                End If
            'End If
            'Determine if LST required to be generated
            ilGenLST = False
            If (imATTExist) And (rbcLogType(1).Value Or rbcLogType(2).Value Or rbcLogType(3).Value) Then
                'Log tested added when changed to generate Affiliate with each Conventional Vehicle instead of
                'the Log Vehicle.  11/20/03
                If tmVef.sType = "L" Then
                    '11/4/09:  Reinstall Log vehicle to generate Affiliate spots if agreements exist
                    'ReDim tmLogVATT(1 To 1) As ATT
                    ilLSTForLogVeh = 0
                    'If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
                        mATTPop tmVef.iCode, slLSTOnly
                        If slLSTOnly <> "Y" Then
                            ilGenLST = mSetGenLST(ilLoop, slMonStartDate, ilExportType)
                            If Not ilGenLST Then
                                ReDim tmLogVATT(0 To 0) As ATT
                            Else
                                ilLSTForLogVeh = 1
                                ReDim Preserve tmLogVATT(0 To UBound(tmVATT)) As ATT
                                For llLoop1 = LBound(tmVATT) To UBound(tmVATT) Step 1
                                    tmLogVATT(llLoop1) = tmVATT(llLoop1)
                                Next llLoop1
                            End If
                        Else
                            ilGenLST = True
                            ReDim tmLogVATT(0 To 0) As ATT
                        End If
                    'Else
                    '    ReDim tmLogVATT(1 To 1) As ATT
                    'End If
                'ElseIf (tmVef.sType = "G") And (tgVpf(ilVpfIndex).sGenLog <> "M") Then
                '    ReDim tmLogVATT(1 To 1) As ATT
                Else
                    mATTPop tmVef.iCode, slLSTOnly
                    If slLSTOnly <> "Y" Then
                        ilGenLST = mSetGenLST(ilLoop, slMonStartDate, ilExportType)
                    Else
                        ilGenLST = True
                        ReDim tmLogVATT(0 To 0) As ATT
                    End If
                End If
                'End of Add.  Before it was just the Two calls mATTPop and mSetGenLST without test for "L"
            End If
            For ilPass = 0 To 1 Step 1
                If ilPass = 1 Then
                    If tmVef.iCombineVefCode <= 0 Then
                        Exit For
                    End If
                    ilFound = False
                    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If (tgMVef(ilVef).iCode = tmVef.iCombineVefCode) Then
                            If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "A") Then
                                ilFound = True
                                ilODFVefCode = tmVef.iCode  'Retain veCode that ODF should be created within
                                tmVef = tgMVef(ilVef)
                            End If
                            Exit For
                        End If
                    Next ilVef
                    If Not ilFound Then
                        Exit For
                    End If
                    tlSvSel = tgSel(ilLoop)
                    tgSel(ilLoop).iVefCode = tmVef.iCode
                    If bgLogFirstCallToVpfFind Then
                        tgSel(ilLoop).iVpfIndex = gVpfFind(Logs, tgSel(ilLoop).iVefCode)
                        bgLogFirstCallToVpfFind = False
                    Else
                        tgSel(ilLoop).iVpfIndex = gVpfFindIndex(tgSel(ilLoop).iVefCode)
                    End If
                    tgSel(ilLoop).sVehicle = Trim$(tmVef.sName)
                    If bgLogFirstCallToVpfFind Then
                        ilVpfIndex = gVpfFind(Logs, tmVef.iCode)
                        bgLogFirstCallToVpfFind = False
                    Else
                        ilVpfIndex = gVpfFindIndex(tmVef.iCode)
                    End If
                    'If tmVef.sType = "C" Then
                    '    If (imATTExist) And (rbcLogType(1).Value Or rbcLogType(2).Value Or rbcLogType(3).Value) Then
                    '        mATTPop tmVef.iCode
                    '        ilGenLST = mSetGenLST(ilLoop, slMonStartDate, ilExportType)
                    '    End If
                    'End If
                Else
                    ilODFVefCode = 0
                End If
                If tmVef.sType = "L" Then
                    If (ckcAssignCopy.Value = vbChecked) And (Not rbcLogType(4).Value) Then
                        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            '7/27/12: Include Sports within Log vehicles
                            'If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                            If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G")) And (tgMVef(ilVef).iVefCode = tmVef.iCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                                If ilDeleteSTF = vbYes Then
                                    smStatusCaption = "Deleting Commercial Changes for " & Trim$(tgMVef(ilVef).sName)
                                    plcStatus.Cls
                                    plcStatus_Paint
                                    tmStfSrchKey.iVefCode = tgMVef(ilVef).iCode
                                    gPackDate slStartDate, tmStfSrchKey.iLogDate(0), tmStfSrchKey.iLogDate(1)
                                    tmStfSrchKey.iLogTime(0) = 0
                                    tmStfSrchKey.iLogTime(1) = 0
                                    ilRet = btrGetGreaterOrEqual(hmStf, tmStf, imStfRecLen, tmStfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmStf.iVefCode = tgMVef(ilVef).iCode)
                                        gUnpackDateLong tmStf.iLogDate(0), tmStf.iLogDate(1), llTestDate
                                        If (llTestDate >= gDateValue(slStartDate)) And (llTestDate <= gDateValue(slEndDate)) Then
                                            If tmStf.sPrint = "R" Then
                                                ilRet = btrGetPosition(hmStf, lmStfRecPos)
                                                Do
                                                    'tmRec = tmStf
                                                    'ilRet = gGetByKeyForUpdate("STF", hmStf, tmRec)
                                                    'tmStf = tmRec
                                                    'On Error GoTo mGenLogErr
                                                    'gBtrvErrorMsg ilRet, "mGenLog (Get by Key)", Logs
                                                    'On Error GoTo 0
                                                    tmStf.sPrint = "D"
                                                    ilRet = btrUpdate(hmStf, tmStf, imStfRecLen)
                                                    If ilRet = BTRV_ERR_CONFLICT Then
                                                        ilCRet = btrGetDirect(hmStf, tmStf, imStfRecLen, lmStfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                    End If
                                                Loop While ilRet = BTRV_ERR_CONFLICT
                                                On Error GoTo mGenLogErr
                                                gBtrvErrorMsg ilRet, "mGenLog (btrUpdate)", Logs
                                                On Error GoTo 0
                                            End If
                                        Else
                                            Exit Do
                                        End If
                                        ilRet = btrGetNext(hmStf, tmStf, imStfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                End If
                                smStatusCaption = "Assigning Copy for " & Trim$(tgMVef(ilVef).sName)
                                plcStatus.Cls
                                plcStatus_Paint
                                'Assign copy
                                If rbcLogType(0).Value Then
                                    ilRet = gAssignCopyToSpots(ilType, tgMVef(ilVef).iCode, 0, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                                Else
                                    ilRet = gAssignCopyToSpots(ilType, tgMVef(ilVef).iCode, 1, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                                End If
                                If Not ilRet Then
                                    gUserActivityLog "E", slGenVehName & ": Log Generation"
                                    imTerminate = True
                                    mGenLog = False
                                    Exit Function
                                End If
                                Do
                                    tmVpfSrchKey.iVefKCode = tgMVef(ilVef).iCode
                                    ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                    On Error GoTo mGenLogErr
                                    gBtrvErrorMsg ilRet, "mGenLog (btrGetEqual)", Logs
                                    On Error GoTo 0
                                    gUnpackDateLong tmVpf.iLLastDateCpyAsgn(0), tmVpf.iLLastDateCpyAsgn(1), llCDate
                                    If gDateValue(slEndDate) > llCDate Then
                                        gPackDate slEndDate, tmVpf.iLLastDateCpyAsgn(0), tmVpf.iLLastDateCpyAsgn(1)
                                        ilRet = btrUpdate(hmVpf, tmVpf, imVpfRecLen)
                                        ilRet = gBinarySearchVpf(tmVpf.iVefKCode)
                                        If ilRet <> -1 Then
                                            tgVpf(ilRet) = tmVpf
                                        End If
                                    Else
                                        ilRet = BTRV_ERR_NONE
                                    End If
                                Loop While ilRet = BTRV_ERR_CONFLICT

                                ilRet = gBinarySearchVpf(tmVpf.iVefKCode)
                                If ilRet <> -1 Then
                                    tgVpf(ilRet) = tmVpf
                                End If
                                If ilZoneExist Then
                                    'Assign copy
                                    If rbcLogType(0).Value Then
                                        ilRet = gAssignCopyToSpots(ilType, tgMVef(ilVef).iCode, 0, slTZStartDate, slTZEndDate, slTZStartTime, slTZEndTime)
                                    Else
                                        ilRet = gAssignCopyToSpots(ilType, tgMVef(ilVef).iCode, 1, slTZStartDate, slTZEndDate, slTZStartTime, slTZEndTime)
                                    End If
                                    If Not ilRet Then
                                        gUserActivityLog "E", slGenVehName & ": Log Generation"
                                        imTerminate = True
                                        mGenLog = False
                                        Exit Function
                                    End If
                                End If
                            End If
                        Next ilVef
                    End If
                    DoEvents
                    If imTerminate Then
                        gUserActivityLog "E", slGenVehName & ": Log Generation"
                        Exit Function
                    End If
                    ilMergeExist = False
                    'Pass 1 was designed to handle vehicles that are associated with Log vehicles but are to have separate Traffic Logs
                    'They appear on the Log screen as separate items and will be generated with Conventional or Airing logic below
                    'For ilLogPass = 0 To 1 Step 1
                    For ilLogPass = 0 To 1 Step 1
                        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "A")) And (tgMVef(ilVef).iVefCode = tmVef.iCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                                If (mMergeWithAffiliate(tgMVef(ilVef).iCode) And (ilLogPass = 0)) Or (Not mMergeWithAffiliate(tgMVef(ilVef).iCode) And (ilLogPass = 1)) Then
                                    If ilLSTForLogVeh = 0 Then
                                        'Added when changed to generate Affiliate with each Conventional Vehicle instead of
                                        'the Log Vehicle.  11/20/03
                                        If (imATTExist) And (rbcLogType(1).Value Or rbcLogType(2).Value Or rbcLogType(3).Value) Then
                                            mATTPop tgMVef(ilVef).iCode, slLSTOnly
                                            If slLSTOnly <> "Y" Then
                                                ilGenLST = mSetGenLST(ilLoop, slMonStartDate, ilExportType)
                                                If ilGenLST Then
                                                    llIndex = UBound(tmLogVATT)
                                                    ReDim Preserve tmLogVATT(0 To llIndex + UBound(tmVATT) - 1) As ATT
                                                    For llLoop1 = LBound(tmVATT) To UBound(tmVATT) - 1 Step 1
                                                        tmLogVATT(llIndex) = tmVATT(llLoop1)
                                                        llIndex = llIndex + 1
                                                    Next llLoop1
                                                End If
                                            Else
                                                ilGenLST = True
                                            End If
                                        End If
                                        'End of Add
                                    End If
                                End If
                            End If
                            If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G")) And (tgMVef(ilVef).iVefCode = tmVef.iCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                                If (mMergeWithAffiliate(tgMVef(ilVef).iCode) And (ilLogPass = 0)) Or (Not mMergeWithAffiliate(tgMVef(ilVef).iCode) And (ilLogPass = 1)) Then
                                    smStatusCaption = "Generating Log for " & Trim$(tgMVef(ilVef).sName)
                                    plcStatus.Cls
                                    plcStatus_Paint
                                    '12/13/12: Move above so that Airing type vehicles include
                                    ''10/20/12:  Handle case where Vehicle is not mergeed with Log vehicle  and agreemnent exist for Log vehicle
                                    '''11/4/09: Add creating LST for Log vehicle
                                    'If ilLSTForLogVeh = 0 Then
                                    '    'Added when changed to generate Affiliate with each Conventional Vehicle instead of
                                    '    'the Log Vehicle.  11/20/03
                                    '    If (imATTExist) And (rbcLogType(1).Value Or rbcLogType(2).Value Or rbcLogType(3).Value) Then
                                    '        mATTPop tgMVef(ilVef).iCode, slLSTOnly
                                    '        If slLSTOnly <> "Y" Then
                                    '            ilGenLST = mSetGenLST(ilLoop, slMonStartDate, ilExportType)
                                    '            If ilGenLST Then
                                    '                llIndex = UBound(tmLogVATT)
                                    '                ReDim Preserve tmLogVATT(1 To llIndex + UBound(tmVATT) - 1) As ATT
                                    '                For llLoop1 = LBound(tmVATT) To UBound(tmVATT) - 1 Step 1
                                    '                    tmLogVATT(llIndex) = tmVATT(llLoop1)
                                    '                    llIndex = llIndex + 1
                                    '                Next llLoop1
                                    '            End If
                                    '        Else
                                    '            ilGenLST = True
                                    '        End If
                                    '    End If
                                    '    'End of Add
                                    'End If
                                    ''7/27/12: Include Sports within Log vehicles
                                    ''ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, ilLSTForLogVeh)
                                    If tgMVef(ilVef).sType <> "G" Then
                                        If ilLogPass = 0 Then
                                            ilMergeExist = True
                                            ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, ilLSTForLogVeh)
                                        ElseIf (ilLogPass = 1) Then
                                            ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, ilLSTForLogVeh)
                                        Else
                                            ilRet = True
                                        End If
                                    Else
                                        ilRet = True
                                        tmGsfSrchKey3.iVefCode = tgMVef(ilVef).iCode
                                        tmGsfSrchKey3.iGameNo = 0
                                        ilCRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
                                        Do While (ilCRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = tgMVef(ilVef).iCode)
                                            gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
                                            If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                                                If tmGsf.sGameStatus <> "C" Then
                                                    ilType = tmGsf.iGameNo
                                                    If ilLogPass = 0 Then
                                                        ilMergeExist = True
                                                        ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmGsf.iVefCode, Format(llDate, "ddddd"), Format(llDate, "ddddd"), slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, ilLSTForLogVeh)
                                                    Else
                                                        ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmGsf.iVefCode, Format(llDate, "ddddd"), Format(llDate, "ddddd"), slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, tmGsf.lCode, ilLSTForLogVeh)
                                                    End If
                                                    If Not ilRet Then
                                                        Exit Do
                                                    End If
                                                    If ilLSTForLogVeh > 0 Then
                                                        ilLSTForLogVeh = 2
                                                    End If
                                                End If
                                            End If
                                            ilCRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                        Loop
                                        ilType = 0
                                    End If
                                    If ilLSTForLogVeh > 0 Then
                                        ilLSTForLogVeh = 2
                                    End If
                                    If Not ilRet Then
                                        gUserActivityLog "E", slGenVehName & ": Log Generation"
                                        imTerminate = True
                                        mGenLog = False
                                        Exit Function
                                    End If
                                    If rbcLogType(0).Value Or rbcLogType(1).Value Then
                                        Do
                                            tmVpfSrchKey.iVefKCode = tgMVef(ilVef).iCode
                                            ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                            On Error GoTo mGenLogErr
                                            gBtrvErrorMsg ilRet, "mGenLog (btrGetEqual)", Logs
                                            On Error GoTo 0
                                            If rbcLogType(1).Value Then
                                                gUnpackDate tmVpf.iLLD(0), tmVpf.iLLD(1), slDate
                                                If gDateValue(slEndDate) > gDateValue(slDate) Then
                                                    gPackDate slEndDate, tmVpf.iLLD(0), tmVpf.iLLD(1)
                                                    'gPackDate slSyncDate, tmVpf.iSyncDate(0), tmVpf.iSyncDate(1)
                                                    'gPackTime slSyncTime, tmVpf.iSyncTime(0), tmVpf.iSyncTime(1)
                                                    ''tmVpf.iSourceID = tgUrf(0).iRemoteUserID
                                                End If
                                            End If
                                            gPackDate slEndDate, tmVpf.iLPD(0), tmVpf.iLPD(1)
                                            ilRet = btrUpdate(hmVpf, tmVpf, imVpfRecLen)
                                        Loop While ilRet = BTRV_ERR_CONFLICT
                                        ilRet = gBinarySearchVpf(tmVpf.iVefKCode)
                                        If ilRet <> -1 Then
                                            tgVpf(ilRet) = tmVpf
                                        End If
                                    End If
                                End If
                            ElseIf (tgMVef(ilVef).sType = "A") And (tgMVef(ilVef).iVefCode = tmVef.iCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                                If (mMergeWithAffiliate(tgMVef(ilVef).iCode) And (ilLogPass = 0)) Or (Not mMergeWithAffiliate(tgMVef(ilVef).iCode) And (ilLogPass = 1)) Then
                                    ilRet = mGenLogForAirVehicle(tgMVef(ilVef), ilDeleteSTF, slGenVehName, slStartDate, slEndDate, slCopyStartDate, ilZoneExist, slTZStartDate, slTZEndDate, slTZStartTime, slTZEndTime, ilGenLST, ilExportType, ilODFVefCode, ilLSTForLogVeh, ilMergeExist, False)
                                    If Not ilRet Then
                                        mGenLog = False
                                        Exit Function
                                    End If
                                    '1/22/13
                                    If ilLSTForLogVeh > 0 Then
                                        ilLSTForLogVeh = 2
                                    End If
                                End If
                            End If
                        Next ilVef
                        If (ilLogPass = 0) And (ilMergeExist) Then
                            'Clear up BuildODFSpotDay information
                            ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 3)
                            ilLSTForLogVeh = 0
                        End If
                    Next ilLogPass
                    ''11/4/09:  Add creation of LST to Log Vehicle
                    'If ilLSTForLogVeh = 0 Then
                        'Added when changed to generate Affiliate with each Conventional Vehicle instead of
                        'the Log Vehicle.  11/20/03
                        'Place all ATT for each Conventional vehicle that belonged to the Log Vehicle into tmVATT so that CPTT will be created
                        If (imATTExist) And (rbcLogType(1).Value Or rbcLogType(2).Value Or rbcLogType(3).Value) Then
                            ReDim Preserve tmVATT(0 To UBound(tmLogVATT)) As ATT
                            For llLoop1 = LBound(tmLogVATT) To UBound(tmLogVATT) - 1 Step 1
                                tmVATT(llLoop1) = tmLogVATT(llLoop1)
                            Next llLoop1
                        End If
                        'End of Add
                    'Else
                    '    'Clear up BuildODFSpotDay information
                    '    ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 3)
                    'End If
                ElseIf tmVef.sType = "A" Then
                    'gBuildLinkArray hmVLF, tmVef, slStartDate, igSVefCode() 'Build igSVefCode so that gBuildODFSpotDay can use it
                    'If (ckcAssignCopy.Value = vbChecked) And (Not rbcLogType(4).Value) Then
                    '    'For ilLink = LBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) To UBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) Step 1
                    '    '    If tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink) > 0 Then
                    '    For ilLink = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                    '            tmVefSrchKey.iCode = igSVefCode(ilLink) 'tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink)
                    '            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    '            On Error GoTo mGenLogErr
                    '            gBtrvErrorMsg ilRet, "mGenLog (btrGetEqual)", Logs
                    '            On Error GoTo 0
                    '            If ilDeleteSTF = vbYes Then
                    '                smStatusCaption = "Deleting Commercial Changes for " & tlVef.sName
                    '                plcStatus.Cls
                    '                plcStatus_Paint
                    '                tmStfSrchKey.iVefCode = tlVef.iCode
                    '                gPackDate slStartDate, tmStfSrchKey.iLogDate(0), tmStfSrchKey.iLogDate(1)
                    '                tmStfSrchKey.iLogTime(0) = 0
                    '                tmStfSrchKey.iLogTime(1) = 0
                    '                ilRet = btrGetGreaterOrEqual(hmStf, tmStf, imStfRecLen, tmStfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    '                Do While (ilRet = BTRV_ERR_NONE) And (tmStf.iVefCode = tlVef.iCode)
                    '                    gUnpackDateLong tmStf.iLogDate(0), tmStf.iLogDate(1), llTestDate
                    '                    If (llTestDate >= gDateValue(slStartDate)) And (llTestDate <= gDateValue(slEndDate)) Then
                    '                        If tmStf.sPrint = "R" Then
                    '                            ilRet = btrGetPosition(hmStf, lmStfRecPos)
                    '                            Do
                    '                                'tmRec = tmStf
                    '                                'ilRet = gGetByKeyForUpdate("STF", hmStf, tmRec)
                    '                                'tmStf = tmRec
                    '                                'On Error GoTo mGenLogErr
                    '                                'gBtrvErrorMsg ilRet, "mGenLog (Get by Key)", Logs
                    '                                'On Error GoTo 0
                    '                                tmStf.sPrint = "D"
                    '                                ilRet = btrUpdate(hmStf, tmStf, imStfRecLen)
                    '                                If ilRet = BTRV_ERR_CONFLICT Then
                    '                                    ilCRet = btrGetDirect(hmStf, tmStf, imStfRecLen, lmStfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    '                                End If
                    '                            Loop While ilRet = BTRV_ERR_CONFLICT
                    '                            On Error GoTo mGenLogErr
                    '                            gBtrvErrorMsg ilRet, "mGenLog (btrUpdate)", Logs
                    '                            On Error GoTo 0
                    '                        End If
                    '                    Else
                    '                        Exit Do
                    '                    End If
                    '                    ilRet = btrGetNext(hmStf, tmStf, imStfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    '                Loop
                    '            End If
                    '            smStatusCaption = "Assigning Copy for " & Trim$(tlVef.sName)
                    '            plcStatus.Cls
                    '            plcStatus_Paint
                    '            'Assign copy
                    '            If rbcLogType(0).Value Then
                    '                'ilRet = gAssignCopyToSpots(slType, tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink), 0, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                    '                ilRet = gAssignCopyToSpots(ilType, igSVefCode(ilLink), 0, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                    '            Else
                    '                'ilRet = gAssignCopyToSpots(slType, tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink), 1, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                    '                ilRet = gAssignCopyToSpots(ilType, igSVefCode(ilLink), 1, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                    '            End If
                    '            If Not ilRet Then
                    '                gUserActivityLog "E", slGenVehName & ": Log Generation"
                    '                imTerminate = True
                    '                mGenLog = False
                    '                Exit Function
                    '            End If
                    '            Do
                    '                tmVpfSrchKey.iVefKCode = igSVefCode(ilLink)  'tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink)
                    '                ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    '                On Error GoTo mGenLogErr
                    '                gBtrvErrorMsg ilRet, "mGenLog (btrGetEqual)", Logs
                    '                On Error GoTo 0
                    '                gUnpackDateLong tmVpf.iLLastDateCpyAsgn(0), tmVpf.iLLastDateCpyAsgn(1), llCDate
                    '                If gDateValue(slEndDate) > llCDate Then
                    '                    gPackDate slEndDate, tmVpf.iLLastDateCpyAsgn(0), tmVpf.iLLastDateCpyAsgn(1)
                    '                    ilRet = btrUpdate(hmVpf, tmVpf, imVpfRecLen)
                    '                Else
                    '                    ilRet = BTRV_ERR_NONE
                    '                End If
                    '            Loop While ilRet = BTRV_ERR_CONFLICT
                    '            ilRet = gBinarySearchVpf(tmVpf.iVefKCode)
                    '            If ilRet <> -1 Then
                    '                tgVpf(ilRet) = tmVpf
                    '            End If

                    '            If ilZoneExist Then
                    '                'Assign copy
                    '                If rbcLogType(0).Value Then
                    '                    ilRet = gAssignCopyToSpots(ilType, igSVefCode(ilLink), 0, slTZStartDate, slTZEndDate, slTZStartTime, slTZEndTime)
                    '                Else
                    '                    ilRet = gAssignCopyToSpots(ilType, igSVefCode(ilLink), 1, slTZStartDate, slTZEndDate, slTZStartTime, slTZEndTime)
                    '                End If
                    '                If Not ilRet Then
                    '                    gUserActivityLog "E", slGenVehName & ": Log Generation"
                    '                    imTerminate = True
                    '                    mGenLog = False
                    '                    Exit Function
                    '                End If
                    '            End If
                    '    '    End If
                    '    Next ilLink
                    '    gAlertVehicleReplace hmVLF
                    'End If
                    'DoEvents
                    'If imTerminate Then
                    '    gUserActivityLog "E", slGenVehName & ": Log Generation"
                    '    Exit Function
                    'End If
                    'smStatusCaption = "Generating Log for " & Trim$(tmVef.sName)
                    'plcStatus.Cls
                    'plcStatus_Paint
                    'ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)
                    'If Not ilRet Then
                    '    gUserActivityLog "E", slGenVehName & ": Log Generation"
                    '    imTerminate = True
                    '    mGenLog = False
                    '    Exit Function
                    'End If
                    'If rbcLogType(0).Value Or rbcLogType(1).Value Then
                    '    'For ilLink = LBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) To UBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) Step 1
                    '    '    If tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink) > 0 Then
                    '    For ilLink = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                    '            Do
                    '                tmVpfSrchKey.iVefKCode = igSVefCode(ilLink) 'tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink)
                    '                ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    '                On Error GoTo mGenLogErr
                    '                gBtrvErrorMsg ilRet, "mGenLog (btrGetEqual)", Logs
                    '                On Error GoTo 0
                    '                If rbcLogType(1).Value Then
                    '                    gUnpackDate tmVpf.iLLD(0), tmVpf.iLLD(1), slDate
                    '                    If gDateValue(slEndDate) > gDateValue(slDate) Then
                    '                        gPackDate slEndDate, tmVpf.iLLD(0), tmVpf.iLLD(1)
                    '                        'gPackDate slSyncDate, tmVpf.iSyncDate(0), tmVpf.iSyncDate(1)
                    '                        'gPackTime slSyncTime, tmVpf.iSyncTime(0), tmVpf.iSyncTime(1)
                    '                        ''tmVpf.iSourceID = tgUrf(0).iRemoteUserID
                    '                    End If
                    '                End If
                    '                gPackDate slEndDate, tmVpf.iLPD(0), tmVpf.iLPD(1)
                    '                ilRet = btrUpdate(hmVpf, tmVpf, imVpfRecLen)
                    '            Loop While ilRet = BTRV_ERR_CONFLICT
                    '            ilRet = gBinarySearchVpf(tmVpf.iVefKCode)
                    '            If ilRet <> -1 Then
                    '                tgVpf(ilRet) = tmVpf
                    '            End If
                    '        'End If
                    '    Next ilLink
                    'End If
                    If (tmVef.iVefCode <= 0) Or (Not mMergeWithLog(tmVef.iCode)) Then
                        ilRet = mGenLogForAirVehicle(tmVef, ilDeleteSTF, slGenVehName, slStartDate, slEndDate, slCopyStartDate, ilZoneExist, slTZStartDate, slTZEndDate, slTZStartTime, slTZEndTime, ilGenLST, ilExportType, ilODFVefCode, 0, ilMergeExist, True)
                        If Not ilRet Then
                            mGenLog = False
                            Exit Function
                        End If
                    End If
                Else
                    If (ilDeleteSTF = vbYes) And (Not rbcLogType(4).Value) Then
                        smStatusCaption = "Deleting Commercial Changes for " & tmVef.sName
                        plcStatus.Cls
                        plcStatus_Paint
                        tmStfSrchKey.iVefCode = tmVef.iCode
                        gPackDate slStartDate, tmStfSrchKey.iLogDate(0), tmStfSrchKey.iLogDate(1)
                        tmStfSrchKey.iLogTime(0) = 0
                        tmStfSrchKey.iLogTime(1) = 0
                        ilRet = btrGetGreaterOrEqual(hmStf, tmStf, imStfRecLen, tmStfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmStf.iVefCode = tmVef.iCode)
                            gUnpackDateLong tmStf.iLogDate(0), tmStf.iLogDate(1), llTestDate
                            If (llTestDate >= gDateValue(slStartDate)) And (llTestDate <= gDateValue(slEndDate)) Then
                                If tmStf.sPrint = "R" Then
                                    ilRet = btrGetPosition(hmStf, lmStfRecPos)
                                    Do
                                        'tmRec = tmStf
                                        'ilRet = gGetByKeyForUpdate("STF", hmStf, tmRec)
                                        'tmStf = tmRec
                                        'On Error GoTo mGenLogErr
                                        'gBtrvErrorMsg ilRet, "mGenLog (Get by Key)", Logs
                                        'On Error GoTo 0
                                        tmStf.sPrint = "D"
                                        ilRet = btrUpdate(hmStf, tmStf, imStfRecLen)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            ilCRet = btrGetDirect(hmStf, tmStf, imStfRecLen, lmStfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    On Error GoTo mGenLogErr
                                    gBtrvErrorMsg ilRet, "mGenLog (btrUpdate)", Logs
                                    On Error GoTo 0
                                End If
                            Else
                                Exit Do
                            End If
                            ilRet = btrGetNext(hmStf, tmStf, imStfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                    If (ckcAssignCopy.Value = vbChecked) And (Not rbcLogType(4).Value) Then
                        smStatusCaption = "Assigning Copy for " & Trim$(tmVef.sName)
                        plcStatus.Cls
                        plcStatus_Paint
                        If rbcLogType(0).Value Then
                            ilRet = gAssignCopyToSpots(ilType, tmVef.iCode, 0, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                        Else
                            ilRet = gAssignCopyToSpots(ilType, tmVef.iCode, 1, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                        End If
                        Do
                            tmVpfSrchKey.iVefKCode = tmVef.iCode
                            ilRet = btrGetEqual(hmVpf, tgVpf(tgSel(ilLoop).iVpfIndex), imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                            On Error GoTo mGenLogErr
                            gBtrvErrorMsg ilRet, "mGenLog (btrGetEqual)", Logs
                            On Error GoTo 0
                            gUnpackDateLong tgVpf(tgSel(ilLoop).iVpfIndex).iLLastDateCpyAsgn(0), tgVpf(tgSel(ilLoop).iVpfIndex).iLLastDateCpyAsgn(1), llCDate
                            If gDateValue(slEndDate) > llCDate Then
                                gPackDate slEndDate, tgVpf(tgSel(ilLoop).iVpfIndex).iLLastDateCpyAsgn(0), tgVpf(tgSel(ilLoop).iVpfIndex).iLLastDateCpyAsgn(1)
                                ilRet = btrUpdate(hmVpf, tgVpf(tgSel(ilLoop).iVpfIndex), imVpfRecLen)
                            Else
                                ilRet = BTRV_ERR_NONE
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilZoneExist Then
                            'Assign copy
                            If rbcLogType(0).Value Then
                                ilRet = gAssignCopyToSpots(ilType, tmVef.iCode, 0, slTZStartDate, slTZEndDate, slTZStartTime, slTZEndTime)
                            Else
                                ilRet = gAssignCopyToSpots(ilType, tmVef.iCode, 1, slTZStartDate, slTZEndDate, slTZStartTime, slTZEndTime)
                            End If
                            If Not ilRet Then
                                gUserActivityLog "E", slGenVehName & ": Log Generation"
                                imTerminate = True
                                mGenLog = False
                                Exit Function
                            End If
                        End If

                    End If
                    DoEvents
                    If imTerminate Then
                        gUserActivityLog "E", slGenVehName & ": Log Generation"
                        Exit Function
                    End If
                    smStatusCaption = "Generating Log for " & Trim$(tmVef.sName)
                    plcStatus.Cls
                    plcStatus_Paint
                    If (tmVef.sType = "G") And (tgVpf(ilVpfIndex).sGenLog <> "M") Then
                        'imATTExist = False
                        If (Asc(tgVpf(ilVpfIndex).sUsingFeatures1) And EXPORTLOG) = EXPORTLOG Then
                            mOpenExportFile slStartDate
                        End If
                        ilRet = True
                        tmGhfSrchKey1.iVefCode = tmVef.iCode
                        ilCRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        'If ilCRet = BTRV_ERR_NONE Then
                        Do While (ilCRet = BTRV_ERR_NONE) And (tmGhf.iVefCode = tmVef.iCode)
                            gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStart
                            gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEnd
                            If (llEndDate >= llSeasonStart) And (llStartDate <= llSeasonEnd) Then
                                tmGsfSrchKey1.lghfcode = tmGhf.lCode
                                tmGsfSrchKey1.iGameNo = 0
                                ilCRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                                Do While (ilCRet = BTRV_ERR_NONE) And (tmGhf.lCode = tmGsf.lghfcode)
                                    gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
                                    If (llDate >= llStartDate) And (llDate <= llEndDate) And (tmGsf.sGameStatus <> "C") Then
                                        If (tgVpf(ilVpfIndex).sGenLog <> "A") Or ((tgVpf(ilVpfIndex).sGenLog = "A") And (tmGsf.sLiveLogMerge <> "M")) Then
                                            ilType = tmGsf.iGameNo
                                            ilODFVefCode = 0
                                            ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, Format(llDate, "ddddd"), Format(llDate, "ddddd"), slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", tmGsf.iGameNo, tmGsf.lCode, 0, True)
                                            'Check for combine
                                            If tmVef.iCombineVefCode > 0 Then
                                                For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                                    If (tgMVef(ilVef).iCode = tmVef.iCombineVefCode) Then
                                                        If (tgMVef(ilVef).sType = "G") Then
                                                            ilFound = False
                                                            ilODFVefCode = tmVef.iCode  'Retain veCode that ODF should be created within
                                                            tmGhfSrchKey1.iVefCode = tgMVef(ilVef).iCode
                                                            ilCRet = btrGetEqual(hmGhf, tmCombineGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                                            'If ilCRet = BTRV_ERR_NONE Then
                                                            Do While (ilCRet = BTRV_ERR_NONE) And (tmCombineGhf.iVefCode = tgMVef(ilVef).iCode)
                                                                gUnpackDateLong tmCombineGhf.iSeasonStartDate(0), tmCombineGhf.iSeasonStartDate(1), llSeasonStart
                                                                gUnpackDateLong tmCombineGhf.iSeasonEndDate(0), tmCombineGhf.iSeasonEndDate(1), llSeasonEnd
                                                                If (llEndDate >= llSeasonStart) And (llStartDate <= llSeasonEnd) Then
                                                                    tmGsfSrchKey1.lghfcode = tmCombineGhf.lCode
                                                                    tmGsfSrchKey1.iGameNo = 0
                                                                    ilCRet = btrGetGreaterOrEqual(hmGsf, tmCombineGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                                                                    Do While (ilCRet = BTRV_ERR_NONE) And (tmCombineGhf.lCode = tmCombineGsf.lghfcode)
                                                                        If tmCombineGsf.sGameStatus <> "C" Then
                                                                            If bgLogFirstCallToVpfFind Then
                                                                                ilCombineVpfIndex = gVpfFind(Logs, tgMVef(ilVef).iCode)
                                                                                bgLogFirstCallToVpfFind = False
                                                                            Else
                                                                                ilCombineVpfIndex = gVpfFindIndex(tgMVef(ilVef).iCode)
                                                                            End If
                                                                            If ilCombineVpfIndex <> -1 Then
                                                                                If (tgVpf(ilCombineVpfIndex).sGenLog <> "A") Or ((tgVpf(ilCombineVpfIndex).sGenLog = "A") And (tmCombineGsf.sLiveLogMerge <> "M")) Then
                                                                                    If (tmGsf.iAirDate(0) = tmCombineGsf.iAirDate(0)) And (tmGsf.iAirDate(1) = tmCombineGsf.iAirDate(1)) And (tmGsf.iAirTime(0) = tmCombineGsf.iAirTime(0)) And (tmGsf.iAirTime(1) = tmCombineGsf.iAirTime(1)) And (tmGsf.iGameNo = tmCombineGsf.iGameNo) Then
                                                                                        ilFound = True
                                                                                        ilType = tmCombineGsf.iGameNo
                                                                                        ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, Format(llDate, "ddddd"), Format(llDate, "ddddd"), slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", tmGsf.iGameNo, tmGsf.lCode, 0)
                                                                                        Exit For
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        ilCRet = btrGetNext(hmGsf, tmCombineGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                                                    Loop
                                                                End If
                                                                ilCRet = btrGetNext(hmGhf, tmCombineGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                                            Loop
                                                            If Not ilFound Then
                                                                ilODFVefCode = tmVef.iCode  'Retain veCode that ODF should be created within
                                                                tmGhfSrchKey1.iVefCode = tgMVef(ilVef).iCode
                                                                ilCRet = btrGetEqual(hmGhf, tmCombineGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                                                'If ilCRet = BTRV_ERR_NONE Then
                                                                Do While (ilCRet = BTRV_ERR_NONE) And (tmCombineGhf.iVefCode = tgMVef(ilVef).iCode)
                                                                    gUnpackDateLong tmCombineGhf.iSeasonStartDate(0), tmCombineGhf.iSeasonStartDate(1), llSeasonStart
                                                                    gUnpackDateLong tmCombineGhf.iSeasonEndDate(0), tmCombineGhf.iSeasonEndDate(1), llSeasonEnd
                                                                    If (llEndDate >= llSeasonStart) And (llStartDate <= llSeasonEnd) Then
                                                                        tmGsfSrchKey1.lghfcode = tmCombineGhf.lCode
                                                                        tmGsfSrchKey1.iGameNo = 0
                                                                        ilCRet = btrGetGreaterOrEqual(hmGsf, tmCombineGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                                                                        Do While (ilCRet = BTRV_ERR_NONE) And (tmCombineGhf.lCode = tmCombineGsf.lghfcode)
                                                                            If tmCombineGsf.sGameStatus <> "C" Then
                                                                                If bgLogFirstCallToVpfFind Then
                                                                                    ilCombineVpfIndex = gVpfFind(Logs, tgMVef(ilVef).iCode)
                                                                                    bgLogFirstCallToVpfFind = False
                                                                                Else
                                                                                    ilCombineVpfIndex = gVpfFindIndex(tgMVef(ilVef).iCode)
                                                                                End If
                                                                                If ilCombineVpfIndex <> -1 Then
                                                                                    If (tgVpf(ilCombineVpfIndex).sGenLog <> "A") Or ((tgVpf(ilCombineVpfIndex).sGenLog = "A") And (tmCombineGsf.sLiveLogMerge <> "M")) Then
                                                                                        If (tmGsf.iAirDate(0) = tmCombineGsf.iAirDate(0)) And (tmGsf.iAirDate(1) = tmCombineGsf.iAirDate(1)) And (tmGsf.iAirTime(0) = tmCombineGsf.iAirTime(0)) And (tmGsf.iAirTime(1) = tmCombineGsf.iAirTime(1)) And (tmGsf.iHomeMnfCode = tmCombineGsf.iHomeMnfCode) And (tmGsf.iVisitMnfCode = tmCombineGsf.iVisitMnfCode) Then
                                                                                            If ((Asc(tgSpf.sSportInfo) And USINGLANG) <> USINGLANG) Or (((Asc(tgSpf.sSportInfo) And USINGLANG) = USINGLANG) And (tmGsf.iLangMnfCode = tmCombineGsf.iLangMnfCode)) Then
                                                                                                ilType = tmCombineGsf.iGameNo
                                                                                                ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, Format(llDate, "ddddd"), Format(llDate, "ddddd"), slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", tmGsf.iGameNo, tmGsf.lCode, 0)
                                                                                                Exit For
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                            ilCRet = btrGetNext(hmGsf, tmCombineGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                                                        Loop
                                                                    End If
                                                                    ilCRet = btrGetNext(hmGhf, tmCombineGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                                                Loop
                                                            End If
                                                        End If
                                                        Exit For
                                                    End If
                                                Next ilVef
                                                'Reset key so that getNext works
                                                tmGsfSrchKey1.lghfcode = tmGhf.lCode
                                                tmGsfSrchKey1.iGameNo = tmGsf.iGameNo
                                                ilCRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                            End If
                                            ilType = tmGsf.iGameNo
                                            'Check if game should be exported
                                            If (Asc(tgVpf(ilVpfIndex).sUsingFeatures1) And EXPORTLOG) = EXPORTLOG Then
                                                If tmGsf.sGameStatus <> "C" Then
                                                    mExportLog slStartDate, slEndDate, tmGsf.iGameNo, tmGsf.lCode
                                                End If
                                            End If
                                            gUnpackDate tmGsf.iAirDate(0), tmGsf.iAirDate(1), slGameDate
                                            mPrintLog ilVpfIndex, tgSel(ilLoop), 0, tmGsf.iGameNo, slGameDate
                                            gDeleteOdf "G", ilType, sLCP, tmVef.iCode
                                            sgGenDate = Format$(gNow(), "m/d/yy")
                                            sgGenTime = Format$(gNow(), "h:mm:ssAM/PM")
                                            gPackDate sgGenDate, igGenDate(0), igGenDate(1)
                                            gPackTime sgGenTime, igGenTime(0), igGenTime(1)
                                            '10-9-01
                                            gUnpackTimeLong igGenTime(0), igGenTime(1), False, lgGenTime
                                        End If
                                    End If
                                    ilCRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                Loop
                                '4/10/15: Game header found for the season, don't need to check any other
                                Exit Do
                            End If
                            ilCRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        Loop
                        ilType = 0
                        ilODFVefCode = 0
                        If (Asc(tgVpf(ilVpfIndex).sUsingFeatures1) And EXPORTLOG) = EXPORTLOG Then
                            Print #hmTo, "Spot: End"
                            Close #hmTo
                        End If
                    Else
                        ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0, True)
                        ilValue = Asc(tgSpf.sSportInfo)
                        If (ilValue And USINGSPORTS) = USINGSPORTS Then
                            tmGsfSrchKey4.iAirVefCode = tmVef.iCode
                            gPackDate slStartDate, tmGsfSrchKey4.iAirDate(0), tmGsfSrchKey4.iAirDate(1)
                            ilCRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
                            Do While (ilCRet = BTRV_ERR_NONE) And (tmGsf.iAirVefCode = tmVef.iCode)
                                gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
                                If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                                    If tmGsf.sGameStatus <> "C" Then
                                        If bgLogFirstCallToVpfFind Then
                                            ilCombineVpfIndex = gVpfFind(Logs, tmGsf.iVefCode)
                                            bgLogFirstCallToVpfFind = False
                                        Else
                                            ilCombineVpfIndex = gVpfFindIndex(tmGsf.iVefCode)
                                        End If
                                        If ilCombineVpfIndex <> -1 Then
                                            If (tgVpf(ilCombineVpfIndex).sGenLog = "M") Or ((tgVpf(ilCombineVpfIndex).sGenLog = "A") And (tmGsf.sLiveLogMerge = "M")) Then
                                                ilType = tmGsf.iGameNo
                                                ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmGsf.iVefCode, Format(llDate, "ddddd"), Format(llDate, "ddddd"), slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, tmVef.iCode, "L", 0, 0, 0)
                                            End If
                                        End If
                                    End If
                                End If
                                If llDate > llEndDate Then
                                    Exit Do
                                End If
                                ilCRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                            Loop
                            ilType = 0
                        End If
                    End If
                    If Not ilRet Then
                        gUserActivityLog "E", slGenVehName & ": Log Generation"
                        imTerminate = True
                        mGenLog = False
                        Exit Function
                    End If
                    'Simulcast Vehicle Log Generation array creation
                    'Removed 1/6/99: Shadow request- to reinstate remove comments on lines below
                    '                Code has previously been tested and it works to generate Logs
                    '                for all simulcast vehicles
                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If (tgMVef(ilVef).sType = "T") And (tgMVef(ilVef).iVefCode = tmVef.iCode) And (tgMVef(ilVef).sState = "A") Then
                    '        tmLogGen(UBound(tmLogGen)).iGenVefCode = tmVef.iCode
                    '        tmLogGen(UBound(tmLogGen)).iSimVefCode = tgMVef(ilVef).iCode
                    '        ReDim Preserve tmLogGen(0 To UBound(tmLogGen) + 1) As LOGGEN
                    '    End If
                    'Next ilVef
                End If
                If rbcLogType(0).Value Or rbcLogType(1).Value Then
                    Do
                        tmVpfSrchKey.iVefKCode = tgSel(ilLoop).iVefCode
                        ilRet = btrGetEqual(hmVpf, tgVpf(tgSel(ilLoop).iVpfIndex), imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        On Error GoTo mGenLogErr
                        gBtrvErrorMsg ilRet, "mGenLog (btrGetEqual)", Logs
                        On Error GoTo 0
                        If rbcLogType(1).Value Then
                            gUnpackDate tgVpf(tgSel(ilLoop).iVpfIndex).iLLD(0), tgVpf(tgSel(ilLoop).iVpfIndex).iLLD(1), slDate
                            If gDateValue(slEndDate) > gDateValue(slDate) Then
                                gPackDate slEndDate, tgVpf(tgSel(ilLoop).iVpfIndex).iLLD(0), tgVpf(tgSel(ilLoop).iVpfIndex).iLLD(1)
                                'gPackDate slSyncDate, tgVpf(tgSel(ilLoop).iVpfIndex).iSyncDate(0), tgVpf(tgSel(ilLoop).iVpfIndex).iSyncDate(1)
                                'gPackTime slSyncTime, tgVpf(tgSel(ilLoop).iVpfIndex).iSyncTime(0), tgVpf(tgSel(ilLoop).iVpfIndex).iSyncTime(1)
                                ''tgVpf(tgSel(ilLoop).iVpfIndex).iSourceID = tgUrf(0).iRemoteUserID
                            End If
                        End If
                        gPackDate slEndDate, tgVpf(tgSel(ilLoop).iVpfIndex).iLPD(0), tgVpf(tgSel(ilLoop).iVpfIndex).iLPD(1)
                        ilRet = btrUpdate(hmVpf, tgVpf(tgSel(ilLoop).iVpfIndex), imVpfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                End If
                If ilPass = 1 Then
                    tgSel(ilLoop) = tlSvSel
                    tmVefSrchKey.iCode = tgSel(ilLoop).iVefCode
                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If bgLogFirstCallToVpfFind Then
                        ilVpfIndex = gVpfFind(Logs, tmVef.iCode)
                        bgLogFirstCallToVpfFind = False
                    Else
                        ilVpfIndex = gVpfFindIndex(tmVef.iCode)
                    End If
                End If
            Next ilPass
            Erase tmCombineLstInfo
            ilRet = mBuildBlackouts()

            On Error GoTo mGenLogErr
            For ilLogGen = LBound(tmLogGen) To UBound(tmLogGen) - 1 Step 1
                On Error GoTo 0
                If ilLogGen <= LBound(tmLogGen) Then
                    ilCycle = tgSel(ilLoop).lEndDate - tgSel(ilLoop).lStartDate + 1
                    llCycleDate = gDateValue(slMonStartDate)
                    Do
                        slCycleDate = Format$(llCycleDate, "m/d/yy")
                        'Clear and save Alerts
                        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            ilProcVefCode = -1
                            If tmVef.sType = "L" Then
                                '7/27/12: Include Sports within Log vehicles
                                'If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                                If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G")) And (tgMVef(ilVef).iVefCode = tmVef.iCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                                    ilProcVefCode = tgMVef(ilVef).iCode
                                End If
                            Else
                                If tgMVef(ilVef).iCode = tmVef.iCode Then
                                    ilProcVefCode = tgMVef(ilVef).iCode
                                End If
                            End If
                            If ilProcVefCode > 0 Then
                                If ilLogGen <= LBound(tmLogGen) Then
                                    ilLogAlertExisted = gAlertClear("A", "L", "S", 0, ilProcVefCode, slCycleDate)
                                    ilCopyAlertExisted = gAlertClear("A", "L", "C", 0, ilProcVefCode, slCycleDate)
                                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                    '    If (tgMVef(ilVef).sType = "T") And (tgMVef(ilVef).iVefCode = ilProcVefCode) Then
                                            If ilLogAlertExisted Then
                                                ilFound = False
                                                For ilTest = 0 To UBound(tgLSTUpdateInfo) - 1 Step 1
                                                    If (tgMVef(ilVef).iCode = tgLSTUpdateInfo(ilTest).iVefCode) And (tgLSTUpdateInfo(ilTest).iType = 0) Then
                                                        If (llCycleDate >= tgLSTUpdateInfo(ilTest).lSDate) And (llCycleDate <= tgLSTUpdateInfo(ilTest).lEDate) Then
                                                            ilFound = True
                                                            Exit For
                                                        End If
                                                    End If
                                                Next ilTest
                                                If Not ilFound Then
                                                    tgLSTUpdateInfo(UBound(tgLSTUpdateInfo)).iType = 0
                                                    tgLSTUpdateInfo(UBound(tgLSTUpdateInfo)).iVefCode = ilProcVefCode
                                                    tgLSTUpdateInfo(UBound(tgLSTUpdateInfo)).lSDate = llCycleDate
                                                    tgLSTUpdateInfo(UBound(tgLSTUpdateInfo)).lEDate = llCycleDate + 6
                                                    ReDim Preserve tgLSTUpdateInfo(0 To UBound(tgLSTUpdateInfo) + 1) As LSTUPDATEINFO
                                                End If
                                            End If
                                            If ilCopyAlertExisted Then
                                                ilFound = False
                                                For ilTest = 0 To UBound(tgLSTUpdateInfo) - 1 Step 1
                                                    If (tgMVef(ilVef).iCode = tgLSTUpdateInfo(ilTest).iVefCode) And (tgLSTUpdateInfo(ilTest).iType = 1) Then
                                                        If (llCycleDate >= tgLSTUpdateInfo(ilTest).lSDate) And (llCycleDate <= tgLSTUpdateInfo(ilTest).lEDate) Then
                                                            ilFound = True
                                                            Exit For
                                                        End If
                                                    End If
                                                Next ilTest
                                                If Not ilFound Then
                                                    tgLSTUpdateInfo(UBound(tgLSTUpdateInfo)).iType = 1
                                                    tgLSTUpdateInfo(UBound(tgLSTUpdateInfo)).iVefCode = ilProcVefCode
                                                    tgLSTUpdateInfo(UBound(tgLSTUpdateInfo)).lSDate = llCycleDate
                                                    tgLSTUpdateInfo(UBound(tgLSTUpdateInfo)).lEDate = llCycleDate + 6
                                                    ReDim Preserve tgLSTUpdateInfo(0 To UBound(tgLSTUpdateInfo) + 1) As LSTUPDATEINFO
                                                End If
                                            End If
                                    '    End If
                                    'Next ilVef
                                End If
                                If tmVef.sType <> "L" Then
                                    Exit For
                                End If
                            End If
                        Next ilVef
                        ilCycle = ilCycle - 7
                        llCycleDate = llCycleDate + 7
                    Loop While ilCycle > 0
                End If
                'Generate Cert of Perf records for Affiliate system
                If (imATTExist) And (rbcLogType(1).Value Or rbcLogType(2).Value Or rbcLogType(3).Value) Then
                    'Find agreement file
                    If ilLogGen <= LBound(tmLogGen) Then
                        'mATTPop tmVef.iCode    'Done above to determine if LST required
                    Else
                        mATTPop tmLogGen(ilLogGen).iSimVefCode, slLSTOnly  'tmVef.iCode
                    End If
                    'Test if CPTT exist- if so delete, otherwise create it
                    ReDim tmCPTTInfo(0 To 0) As CPTTINFO
                    ilCycle = tgSel(ilLoop).lEndDate - tgSel(ilLoop).lStartDate + 1
                    llCycleDate = gDateValue(slMonStartDate)
                    Do
                        'Added when changed to generate Affiliate with each Conventional Vehicle instead of
                        'the Log Vehicle.  11/20/03
                        '11/4/09:  Re-add the generation of LST by Log Vehicle
                        'If tmVef.sType = "L" Then
                        If (tmVef.sType = "L") And (ilLSTForLogVeh = 0) Then
                            For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                '7/27/12: Include Sports within Log vehicles
                                'If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                                If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "A")) And (tgMVef(ilVef).iVefCode = tmVef.iCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                                    Do
                                        tmCPTTSrchKey1.iVefCode = tgMVef(ilVef).iCode
                                        gPackDateLong llCycleDate, tmCPTTSrchKey1.iStartDate(0), tmCPTTSrchKey1.iStartDate(1)
                                        ilRet = btrGetEqual(hmCPTT, tmCPTT, imCPTTRecLen, tmCPTTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                        If ilRet = BTRV_ERR_NONE Then
                                            tmCPTTSrchKey.lCode = tmCPTT.lCode
                                            ilRet = btrGetEqual(hmCPTT, tmCPTT, imCPTTRecLen, tmCPTTSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                            If ilRet = BTRV_ERR_NONE Then
                                                ilRet = btrDelete(hmCPTT)
                                            End If
                                            If ilRet = BTRV_ERR_NONE Then
                                                tmCPTTInfo(UBound(tmCPTTInfo)).lAtfCode = tmCPTT.lAtfCode
                                                tmCPTTInfo(UBound(tmCPTTInfo)).iShfCode = tmCPTT.iShfCode
                                                tmCPTTInfo(UBound(tmCPTTInfo)).iVefCode = tmCPTT.iVefCode
                                                tmCPTTInfo(UBound(tmCPTTInfo)).iReturnDate(0) = tmCPTT.iReturnDate(0)
                                                tmCPTTInfo(UBound(tmCPTTInfo)).iReturnDate(1) = tmCPTT.iReturnDate(1)
                                                tmCPTTInfo(UBound(tmCPTTInfo)).iStatus = tmCPTT.iStatus
                                                tmCPTTInfo(UBound(tmCPTTInfo)).iPostingStatus = tmCPTT.iPostingStatus
                                                'tmCPTTInfo(UBound(tmCPTTInfo)).iPrintStatus = tmCPTT.iPrintStatus
                                                tmCPTTInfo(UBound(tmCPTTInfo)).sAstStatus = tmCPTT.sAstStatus
                                                tmCPTTInfo(UBound(tmCPTTInfo)).lCycleDate = llCycleDate
                                                ReDim Preserve tmCPTTInfo(0 To UBound(tmCPTTInfo) + 1) As CPTTINFO
                                            End If
                                        Else
                                            Exit Do
                                        End If
                                    Loop
                                End If
                            Next ilVef
                            '1/18/13: Find CPTT for Log vehicle
                            Do
                                tmCPTTSrchKey1.iVefCode = tmVef.iCode
                                gPackDateLong llCycleDate, tmCPTTSrchKey1.iStartDate(0), tmCPTTSrchKey1.iStartDate(1)
                                ilRet = btrGetEqual(hmCPTT, tmCPTT, imCPTTRecLen, tmCPTTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                If ilRet = BTRV_ERR_NONE Then
                                    tmCPTTSrchKey.lCode = tmCPTT.lCode
                                    ilRet = btrGetEqual(hmCPTT, tmCPTT, imCPTTRecLen, tmCPTTSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                    If ilRet = BTRV_ERR_NONE Then
                                        ilRet = btrDelete(hmCPTT)
                                    End If
                                    If ilRet = BTRV_ERR_NONE Then
                                        tmCPTTInfo(UBound(tmCPTTInfo)).lAtfCode = tmCPTT.lAtfCode
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iShfCode = tmCPTT.iShfCode
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iVefCode = tmCPTT.iVefCode
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iReturnDate(0) = tmCPTT.iReturnDate(0)
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iReturnDate(1) = tmCPTT.iReturnDate(1)
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iStatus = tmCPTT.iStatus
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iPostingStatus = tmCPTT.iPostingStatus
                                        'tmCPTTInfo(UBound(tmCPTTInfo)).iPrintStatus = tmCPTT.iPrintStatus
                                        tmCPTTInfo(UBound(tmCPTTInfo)).sAstStatus = tmCPTT.sAstStatus
                                        tmCPTTInfo(UBound(tmCPTTInfo)).lCycleDate = llCycleDate
                                        ReDim Preserve tmCPTTInfo(0 To UBound(tmCPTTInfo) + 1) As CPTTINFO
                                    End If
                                Else
                                    Exit Do
                                End If
                            Loop
                            'End if code added plus Else and EndIf
                        Else
                            Do
                                If ilLogGen <= LBound(tmLogGen) Then
                                    tmCPTTSrchKey1.iVefCode = tmVef.iCode
                                Else
                                    tmCPTTSrchKey1.iVefCode = tmLogGen(ilLogGen).iSimVefCode  'tmVef.iCode
                                End If
                                gPackDateLong llCycleDate, tmCPTTSrchKey1.iStartDate(0), tmCPTTSrchKey1.iStartDate(1)
                                ilRet = btrGetEqual(hmCPTT, tmCPTT, imCPTTRecLen, tmCPTTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                If ilRet = BTRV_ERR_NONE Then
                                    tmCPTTSrchKey.lCode = tmCPTT.lCode
                                    ilRet = btrGetEqual(hmCPTT, tmCPTT, imCPTTRecLen, tmCPTTSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                    If ilRet = BTRV_ERR_NONE Then
                                        ilRet = btrDelete(hmCPTT)
                                    End If
                                    If ilRet = BTRV_ERR_NONE Then
                                        tmCPTTInfo(UBound(tmCPTTInfo)).lAtfCode = tmCPTT.lAtfCode
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iShfCode = tmCPTT.iShfCode
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iVefCode = tmCPTT.iVefCode
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iReturnDate(0) = tmCPTT.iReturnDate(0)
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iReturnDate(1) = tmCPTT.iReturnDate(1)
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iStatus = tmCPTT.iStatus
                                        tmCPTTInfo(UBound(tmCPTTInfo)).iPostingStatus = tmCPTT.iPostingStatus
                                        'tmCPTTInfo(UBound(tmCPTTInfo)).iPrintStatus = tmCPTT.iPrintStatus
                                        tmCPTTInfo(UBound(tmCPTTInfo)).sAstStatus = tmCPTT.sAstStatus
                                        tmCPTTInfo(UBound(tmCPTTInfo)).lCycleDate = llCycleDate
                                        ReDim Preserve tmCPTTInfo(0 To UBound(tmCPTTInfo) + 1) As CPTTINFO
                                    End If
                                Else
                                    Exit Do
                                End If
                            Loop
                        End If
                        ilCycle = ilCycle - 7
                        llCycleDate = llCycleDate + 7
                    Loop While ilCycle > 0
                    'tmVAtt is by Conventional vehicle if Not processing Log vehicle or the Conventional vehicles that make up a Log Vehicle
                    'or it is the Simulcast vehicle if processing the simulcast
                    For llLoop1 = LBound(tmVATT) To UBound(tmVATT) - 1 Step 1
                        ilCycle = tgSel(ilLoop).lEndDate - tgSel(ilLoop).lStartDate + 1
                        llCycleDate = gDateValue(slMonStartDate)
                        Do
                            tmAtt = tmVATT(llLoop1)
                            slCycleDate = Format$(llCycleDate, "m/d/yy")
                            'If ilLogGen <= LBound(tmLogGen) Then
                                'Get Alerts from saved value
                                ilLogAlertExisted = False
                                ilCopyAlertExisted = False
                                For ilTest = 0 To UBound(tgLSTUpdateInfo) - 1 Step 1
                                    If tmAtt.iVefCode = tgLSTUpdateInfo(ilTest).iVefCode Then
                                        If (llCycleDate >= tgLSTUpdateInfo(ilTest).lSDate) And (llCycleDate <= tgLSTUpdateInfo(ilTest).lEDate) Then
                                            If tgLSTUpdateInfo(ilTest).iType = 0 Then
                                                ilLogAlertExisted = True
                                                Exit For
                                            ElseIf tgLSTUpdateInfo(ilTest).iType = 1 Then
                                                ilCopyAlertExisted = True
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next ilTest
                            'Else
                            '    ilLogAlertExisted = False
                            '    ilCopyAlertExisted = False
                            '    For ilTest = 0 To UBound(tgLSTUpdateInfo) - 1 Step 1
                            '        If tmLogGen(ilLogGen).iSimVefCode = tgLSTUpdateInfo(ilTest).iVefCode Then
                            '            If (llCycleDate >= tgLSTUpdateInfo(ilTest).lSDate) And (llCycleDate <= tgLSTUpdateInfo(ilTest).lEDate) Then
                            '                If tgLSTUpdateInfo(ilTest).iType = 0 Then
                            '                    ilLogAlertExisted = True
                            '                    Exit For
                            '                ElseIf tgLSTUpdateInfo(ilTest).iType = 1 Then
                            '                    ilCopyAlertExisted = True
                            '                    Exit For
                            '                End If
                            '            End If
                            '        End If
                            '    Next ilTest
                            'End If
                            gUnpackDateLong tmAtt.iOnAir(0), tmAtt.iOnAir(1), llDate1
                            Do While gWeekDayLong(llDate1) <> 0
                                llDate1 = llDate1 - 1
                            Loop
                            gUnpackDateLong tmAtt.iOffAir(0), tmAtt.iOffAir(1), llDate2
                            If llDate2 <> 0 Then
                                Do While gWeekDayLong(llDate2) <> 6
                                    llDate2 = llDate2 + 1
                                Loop
                            End If
                            If (llCycleDate >= llDate1) And ((llCycleDate <= llDate2) Or (llDate2 = 0)) Then
                                gUnpackDateLong tmAtt.iDropDate(0), tmAtt.iDropDate(1), llDate1
                                If llDate1 <> 0 Then
                                    Do While gWeekDayLong(llDate1) <> 6
                                        llDate1 = llDate1 + 1
                                    Loop
                                End If
                                If ((llCycleDate <= llDate1) Or (llDate1 = 0)) Then
                                    tmCPTT.iStatus = 0  'Not returned
                                    tmCPTT.iPostingStatus = 0
                                    'tmCPTT.iPrintStatus = 0
                                    gPackDate "", tmCPTT.iReturnDate(0), tmCPTT.iReturnDate(1)
                                    tmCPTT.sAstStatus = "N"
                                    For llLoop2 = 0 To UBound(tmCPTTInfo) - 1 Step 1
                                        If (tmCPTTInfo(llLoop2).lAtfCode = tmAtt.lCode) And (tmCPTTInfo(llLoop2).iShfCode = tmAtt.iShfCode) And (tmCPTTInfo(llLoop2).iVefCode = tmAtt.iVefCode) And (tmCPTTInfo(llLoop2).lCycleDate = llCycleDate) Then
                                            tmCPTT.iStatus = tmCPTTInfo(llLoop2).iStatus
                                            tmCPTT.iPostingStatus = tmCPTTInfo(llLoop2).iPostingStatus
                                            'tmCPTT.iPrintStatus = tmCPTTInfo(llLoop2).iPrintStatus
                                            tmCPTT.iReturnDate(0) = tmCPTTInfo(llLoop2).iReturnDate(0)
                                            tmCPTT.iReturnDate(1) = tmCPTTInfo(llLoop2).iReturnDate(1)
                                            'If rbcLogType(1).Value Then
                                            '    'tmCPTT.iStatus = 0  'Not returned
                                            '    tmCPTT.iPrintStatus = 0
                                            'Else
                                            '    'tmCPTT.iStatus = 0  'Not returned
                                            '    If tmCPTT.iPrintStatus > 0 Then
                                            '        tmCPTT.iPrintStatus = 2
                                            '    End If
                                            'End If
                                            If (rbcLogType(2).Value) Or (rbcLogType(3).Value) Then
                                                If tgSel(ilLoop).lEndDate < lmNowDate Then
                                                    tmCPTT.sAstStatus = tmCPTTInfo(llLoop2).sAstStatus
                                                Else
                                                    If tmCPTTInfo(llLoop2).sAstStatus = "C" Then
                                                        'Only set to R if coppy changed or Spots moved
                                                        'Code later
                                                        If ilLogAlertExisted Or ilCopyAlertExisted Then
                                                            tmCPTT.sAstStatus = "R"
                                                        Else
                                                            tmCPTT.sAstStatus = "C"
                                                        End If
                                                    Else
                                                        tmCPTT.sAstStatus = tmCPTTInfo(llLoop2).sAstStatus
                                                    End If
                                                End If
                                            End If
                                            Exit For
                                        End If
                                    Next llLoop2
                                    '10/7/11: Test that LST created, if not don't create cptt
                                    imLstRecLen = Len(tmLst)
                                    tmLstSrchKey2.iLogVefCode = tmAtt.iVefCode
                                    gPackDateLong llCycleDate, tmLstSrchKey2.iLogDate(0), tmLstSrchKey2.iLogDate(1)
                                    ilRet = btrGetGreaterOrEqual(hmLst, tmLst, imLstRecLen, tmLstSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                    If (ilRet = BTRV_ERR_NONE) And (tmLst.iLogVefCode = tmAtt.iVefCode) Then
                                        gUnpackDateLong tmLst.iLogDate(0), tmLst.iLogDate(1), llLstDate
                                    Else
                                        llLstDate = 0
                                    End If
                                    If (llLstDate >= llCycleDate) And (llLstDate <= llCycleDate + 6) Then
                                        tmCPTT.lCode = 0
                                        tmCPTT.lAtfCode = tmAtt.lCode
                                        tmCPTT.iShfCode = tmAtt.iShfCode
                                        tmCPTT.iVefCode = tmAtt.iVefCode
                                        gPackDate smNowDate, tmCPTT.iCreateDate(0), tmCPTT.iCreateDate(1)
                                        gPackDateLong llCycleDate, tmCPTT.iStartDate(0), tmCPTT.iStartDate(1)
                                        'tmCPTT.iCycle = tgSel(ilLoop).lEndDate - tgSel(ilLoop).lStartDate + 1
                                        'Set air time as start of daypart of avail
                                        'gPackTime "12M", tmCPTT.iAirTime(0), tmCPTT.iAirTime(1)
                                        tmCPTT.iNoSpotsGen = 0
                                        tmCPTT.iNoSpotsAired = 0
                                        tmCPTT.iNoCompliant = 0
                                        tmCPTT.iUsfCode = 0
                                        '1/18/13: Add removal of Duplicate CPTT's as a safety feature.  It should not be required
                                        'ilRet = btrInsert(hmCPTT, tmCPTT, imCPTTRecLen, INDEXKEY0)
                                        ilRet = mInsertCPTT(tmCPTT, tmAtt)
                                        slCycleDate = Format$(llCycleDate, "m/d/yy")
                                        If ilLogAlertExisted Or rbcLogType(1).Value Or rbcLogType(2).Value Or rbcLogType(3).Value Then
                                            'Only create alert if export via Web or Marketron
                                            If (tmAtt.iExportType = 1) Or (tmAtt.iExportType = 2) Then
                                                'ilRet = gAlertAdd(smLogType, "S", 0, tmAtt.iVefCode, slCycleDate)
                                                gMakeExportAlert tmAtt.iVefCode, llCycleDate, smLogType, "S"
                                            End If
                                        End If
                                        If ilCopyAlertExisted Or rbcLogType(1).Value Or rbcLogType(2).Value Or rbcLogType(3).Value Then
                                            'Add both types of alerts to affiliate when copy changed as copy is exported and shows on logs
                                            If (tmAtt.iExportType = 1) Or (tmAtt.iExportType = 2) Then
                                                'ilRet = gAlertAdd(smLogType, "S", 0, tmAtt.iVefCode, slCycleDate)
                                                gMakeExportAlert tmAtt.iVefCode, llCycleDate, smLogType, "S"
                                            End If
                                            'Only update ISCI Alert if Provider defined and not embedded or Provider and Produce defined and embedded
                                            'Same logic in affiliate when showing vehicles to generate ISCI for.
                                            ilRet = gBinarySearchVpf(tmAtt.iVefCode)
                                            If ilRet <> -1 Then
                                                If tgVpf(ilRet).iCommProvArfCode > 0 Then
                                                    ''1/21/10:  Removed embedded from producer definition (Vehicle-Option).
                                                    ''If (tgVpf(ilRet).sEmbeddedComm <> "Y") Or ((tgVpf(ilRet).sEmbeddedComm = "Y") And (tgVpf(ilRet).iProducerArfCode > 0)) Then
                                                        'ilRet = gAlertAdd(smLogType, "I", 0, tmAtt.iVefCode, slCycleDate)
                                                        gMakeExportAlert tmAtt.iVefCode, llCycleDate, smLogType, "I"
                                                    ''End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            ilCycle = ilCycle - 7
                            llCycleDate = llCycleDate + 7
                        Loop While ilCycle > 0
                    Next llLoop1
                End If
                If (tmVef.sType = "G") And (tgVpf(ilVpfIndex).sGenLog <> "M") Then
                Else
                    If (Asc(tgVpf(ilVpfIndex).sUsingFeatures1) And EXPORTLOG) = EXPORTLOG Then
                        mOpenExportFile slStartDate
                        mExportLog slStartDate, slEndDate, 0, 0
                    End If
                    mPrintLog ilVpfIndex, tgSel(ilLoop), ilLogGen, 0, ""
                    If (Asc(tgVpf(ilVpfIndex).sUsingFeatures1) And EXPORTLOG) = EXPORTLOG Then
                        Print #hmTo, "Spot: End"
                        Close #hmTo
                    End If
                    If ilLogGen < UBound(tmLogGen) - 1 Then
                        'Simulcast vehicles
                        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            If (tgMVef(ilVef).sType = "T") And (tgMVef(ilVef).iCode = tmLogGen(ilLogGen + 1).iSimVefCode) Then
                                smStatusCaption = "Generating Log for " & Trim$(tgMVef(ilVef).sName)
                                plcStatus.Cls
                                plcStatus_Paint
                                'If using blackouts, then odf and lst must be corrected to match first output OR
                                'Loop thru ODF and Change vehicle and insert LST (remove old lst).
                                'Returning to use btrDelete instead of copy odf_blk
                                'Jim-3/26/01
                                'If tgSpf.sGUseAffSys = "Y" Then
                                    gDeleteOdf "G", ilType, sLCP, tgSel(ilLoop).iVefCode
                                'End If
                                sgGenDate = Format$(gNow(), "m/d/yy")
                                sgGenTime = Format$(gNow(), "h:mm:ssAM/PM")
                                gPackDate sgGenDate, igGenDate(0), igGenDate(1)
                                gPackTime sgGenTime, igGenTime(0), igGenTime(1)
                                ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmLogGen(ilLogGen + 1).iGenVefCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), tmLogGen(ilLogGen + 1).iSimVefCode, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)
                                Exit For
                            End If
                        Next ilVef
                    End If
                End If
            Next ilLogGen
            'Returning to use btrDelete instead of copy odf_blk
            'Jim-3/26/01
            'If tgSpf.sGUseAffSys = "Y" Then
                gDeleteOdf "G", ilType, sLCP, tgSel(ilLoop).iVefCode
            'End If
            gUserActivityLog "E", slGenVehName & ": Log Generation"
            ilRet = mSaveAbf(False)
        End If
    Next ilLoop
    
    ilRet = mSaveAbf(True)
    mUpdateLLDForBypaasedVeh slSyncDate, slSyncTime
    
    Erase tgNTRInfo             '1-23-14 for podcasting vehicles of NTR built into memory
    Erase tgNtrSortInfo
    smStatusCaption = ""
    plcStatus.Cls
    mGenLog = True
    Exit Function

    ilRet = 1
    Resume Next
mGenLogErr:
    imTerminate = True
    mGenLog = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Initialize module              *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Screen.MousePointer = vbHourglass
    ReDim tgLogSel(0 To 0) As LOGSEL
    ReDim tgSel(0 To 0) As LOGSEL
    igJobShowing(LOGSJOB) = True
    imLBCDCtrls = 1
    imLBCtrls = 1
    imFirstActivate = True
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.height = 165
    imCurSort = 0
    'Initialize variables
    imTerminate = False         'terminate if true
    bgLogFirstCallToVpfFind = True
    'Logs.Height = cmcGenerate.Top + 5 * cmcGenerate.Height / 3
    'gCenterForm Logs
    'Initialize positioning and show form
    mInitBox
    gCenterForm Logs
    'imcHelp.Picture = Traffic!imcHelp.Picture
    Logs.Show
    'Moved to Basic10 area
    'If tgSpf.sMktBase = "Y" Then
    '    imGettingMkt = True
    '    LogMkt.Show vbModal
    'End If
    Screen.MousePointer = vbHourglass
    imGettingMkt = False
    imFirstTime = True
    imGeneratingLog = False
    imBypassFocus = False       'don't bypass focus on any control
    imChgMode = False           'no change made
    imBSMode = False            'back space key
    imCalType = 0               'Standard type
    imBoxNo = -1                'Initialize current Box to N/A
    imRowNo = -1
    imTabDirection = 0          'Left to right movement
    imSettingValue = False
    imShiftKey = 0
    imLbcArrowSetting = False   'List box invisible
    imIgnoreRightMove = False
    imButton = 0
    smDate = Format$(gNow(), "m/d/yy") 'Correctly format current date
    smNowDate = smDate
    lmNowDate = gDateValue(smNowDate)
    smDate = gIncOneDay(smDate) 'Default date
    imLcfFound = False          'Valid Lcf date found=false
    smDefaultTime = "12M"       'set default time to 12 midnight
    imVehCode = -1              'Invalidate vehicle code
    imPFAllowed = True      'False         'Final only
    imFeedCode = -1
    imAssCopy = True    'False       'Change to True- Record locking required before allowing assign copy
    smDefaultDate = smDate      'temporarily set default date to now +1
    imCpyFlag = False     'set flag to allow assign copy field
    imRPGen = False
    imAlertGen = False
    imLogType = 1   'Final Log
    imPbcIndex = 1
    'Open btrieve files
    imVefRecLen = Len(tmVef)    'Save VEF record length
    hmVef = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: VEF.BTR)", Logs
    On Error GoTo 0
    imVpfRecLen = Len(tmVpf)    'Save VEF record length
    hmVpf = CBtrvTable(TWOHANDLES)          'Save VEF handle
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: VPF.BTR)", Logs
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)    'Save VEF record length
    hmVLF = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: VLF.BTR)", Logs
    On Error GoTo 0
    imStfRecLen = Len(tmStf)    'Save STF record length
    hmStf = CBtrvTable(TWOHANDLES)          'Save STF handle
    ilRet = btrOpen(hmStf, "", sgDBPath & "Stf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: STF.BTR)", Logs
    On Error GoTo 0
    imRnfRecLen = Len(tmRnf)    'Save RNF record length
    hmRnf = CBtrvTable(ONEHANDLE)          'Save RNF handle
    ilRet = btrOpen(hmRnf, "", sgDBPath & "Rnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: RNF.BTR)", Logs
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)    'Save RNF record length
    hmGhf = CBtrvTable(ONEHANDLE)          'Save RNF handle
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: GHF.BTR)", Logs
    On Error GoTo 0
    imGsfRecLen = Len(tmGsf)    'Save RNF record length
    hmGsf = CBtrvTable(ONEHANDLE)          'Save RNF handle
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: GSF.BTR)", Logs
    On Error GoTo 0

    imSsfRecLen = Len(tmSsf)    'Save VEF record length
    hmSsf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: SSF.BTR)", Logs
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)    'Save VEF record length
    hmSdf = CBtrvTable(TWOHANDLES)          'Save VEF handle
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: SDF.BTR)", Logs
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)  'Get and save ADF record length
    hmCpf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", Logs
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)  'Get and save ADF record length
    hmMcf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mcf.Btr)", Logs
    On Error GoTo 0
    imCifRecLen = Len(tmCif)  'Get and save ADF record length
    hmCif = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", Logs
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)  'Get and save ADF record length
    hmAnf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Anf.Btr)", Logs
    On Error GoTo 0
    If tgSpf.sGUseAffSys = "Y" Then
        imATTExist = True
        imAttRecLen = Len(tmAtt)  'Get and save ADF record length
        hmAtt = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmAtt, "", sgDBPath & "ATT.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: ATT.MKD)", Logs
        On Error GoTo 0
        imCPTTRecLen = Len(tmCPTT)  'Get and save ADF record length
        hmCPTT = CBtrvTable(TWOHANDLES)        'Create ADF object handle
        ilRet = btrOpen(hmCPTT, "", sgDBPath & "CPTT.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: CPTT.MKD)", Logs
        On Error GoTo 0
        imSHTTRecLen = Len(tmSHTT)  'Get and save ADF record length
        hmSHTT = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmSHTT, "", sgDBPath & "SHTT.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: SHTT.MKD)", Logs
        On Error GoTo 0
        imLstExist = True
        imLstRecLen = Len(tmLst)  'Get and save ADF record length
        hmLst = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmLst, "", sgDBPath & "Lst.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: LST.MKD)", Logs
        On Error GoTo 0
        imAbfRecLen = Len(tmAbf)  'Get and save ADF record length
        hmAbf = CBtrvTable(TWOHANDLES)        'Create ADF object handle
        ilRet = btrOpen(hmAbf, "", sgDBPath & "Abf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: ABF.BTR)", Logs
        On Error GoTo 0
        imSefRecLen = Len(tmSef)    'Save RNF record length
        hmSef = CBtrvTable(ONEHANDLE)          'Save RNF handle
        ilRet = btrOpen(hmSef, "", sgDBPath & "Sef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: SEF.BTR)", Logs
        On Error GoTo 0
        imAxfRecLen = Len(tmAxf)    'Save RNF record length
        hmAxf = CBtrvTable(ONEHANDLE)          'Save RNF handle
        ilRet = btrOpen(hmAxf, "", sgDBPath & "Axf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: AXF.BTR)", Logs
        On Error GoTo 0
    Else
        imATTExist = False
    End If
    If (tgSpf.sCBlackoutLog = "Y") Or (((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS)) Or ((Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY) Then
        'If tgSpf.sGUseAffSys <> "Y" Then
        '    imMcfRecLen = Len(tmMcf)  'Get and save ADF record length
        '    hmMcf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        '    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        '    On Error GoTo mInitErr
        '    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mcf.Btr)", Logs
        '    On Error GoTo 0
        'End If
        imCHFRecLen = Len(tmChf)  'Get and save ADF record length
        hmCHF = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", Logs
        On Error GoTo 0
        imClfRecLen = Len(tmClf)  'Get and save ADF record length
        hmClf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", Logs
        On Error GoTo 0
        imBofRecLen = Len(tmBof)  'Get and save ADF record length
        hmBof = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmBof, "", sgDBPath & "Bof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Bof.Btr)", Logs
        On Error GoTo 0
        imRsfRecLen = Len(tmRsf)    'Save RNF record length
        hmRsf = CBtrvTable(TWOHANDLES)          'Save RNF handle
        ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: RSF.BTR)", Logs
        On Error GoTo 0
        imPrfRecLen = Len(tmPrf)  'Get and save ADF record length
        hmPrf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Prf.Btr)", Logs
        On Error GoTo 0
        imSifRecLen = Len(tmSif)  'Get and save ADF record length
        hmSif = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Sif.Btr)", Logs
        On Error GoTo 0
        imCrfRecLen = Len(tmCrf)  'Get and save ADF record length
        hmCrf = CBtrvTable(TWOHANDLES)        'Create ADF object handle
        ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Crf.Btr)", Logs
        On Error GoTo 0
        imCnfRecLen = Len(tmCnf)  'Get and save ADF record length
        hmCnf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Cnf.Btr)", Logs
        On Error GoTo 0
    End If
    If tgSpf.sCDefLogCopy = "Y" Then
        ckcAssignCopy.Value = vbChecked
    Else
        ckcAssignCopy.Value = vbUnchecked
    End If
    On Error GoTo 0


    'Logs.Height = cmcGenerate.Top + 5 * cmcGenerate.Height / 3
    'gCenterForm Logs
    ilRet = gPopAdvtBox(Logs, lbcAdvt, tmAdvertiser(), smAdvertiserTag)
    
    '4/26/11: Add test of avail attribute
    ilRet = gObtainAvail()
    
    '10/21/12: Required because of Log Merge test
    ilRet = gVffRead()

    
    lbcVehicle.Clear    'Initialize List boxes
    'lbcVehCode.Clear
    ReDim tgLogVehicle(0 To 0) As SORTCODE
    sgLogVehicleTag = ""
    ' dan M 6-17-08 moved below to mLogPop to allow refresh
'    lbcTimeZ.Clear
'    lbcTimeZ.AddItem "[All]"
'    lbcTimeZ.AddItem "EST"
'    lbcTimeZ.AddItem "CST"
'    lbcTimeZ.AddItem "MST"
'    lbcTimeZ.AddItem "PST"
    mRnfPop
    mLogPop           'Populate vehicle list boxes
    If imTerminate Then
        Exit Sub
    End If

    ilRet = gObtainMnfForType("Z", smTeamTag, tmTeam())
    ilRet = gObtainMnfForType("L", smLangTag, tmLang())

    edcTime(0).Text = "12M"
    edcTime(1).Text = "12M"
    'If (tgSpf.sSSellNet = "Y") Or (tgSpf.sSDelNet = "Y") Then
    '    imDelivery = gRecExistForFile("Dlf.Btr")
    'Else
        imDelivery = False
    'End If
    If ckcOutput(0).Value = vbUnchecked Then   'Select printing
        If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            If tgSpf.sGUseAffSys = "Y" Then
                ''rbcOutput(0).Enabled = True
                ''rbcOutput(0).Caption = "None"
                ''rbcOutput(0).Value = True
                '5/10/12:  Allow display of final log if site set to Y
                'ckcOutput(0).Enabled = False
                'ckcOutput(0).Value = vbUnchecked
                If tgSaf(0).sFinalLogDisplay <> "Y" Then
                    ckcOutput(0).Enabled = False
                    ckcOutput(0).Value = vbUnchecked
                Else
                    ckcOutput(0).Enabled = True
                End If
                'ckcOutput(1).Value = vbUnchecked
                ckcOutput(2).Value = vbUnchecked
            Else
                ''rbcOutput(0).Enabled = False
                '5/10/12:  Allow display of final log if site set to Y
                'ckcOutput(0).Enabled = False
                'ckcOutput(0).Value = vbUnchecked
                If tgSaf(0).sFinalLogDisplay <> "Y" Then
                    ckcOutput(0).Enabled = False
                    ckcOutput(0).Value = vbUnchecked
                Else
                    ckcOutput(0).Enabled = True
                End If
            End If
        Else
            'rbcOutput(0).Enabled = True
            ckcOutput(0).Enabled = True
        End If
    End If
    If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        If tgSpf.sAllowPrelLog <> "Y" Then
            rbcLogType(0).Enabled = False
        Else
            rbcLogType(0).Enabled = True
        End If
    Else
        rbcLogType(0).Enabled = True
    End If
    If tgSpf.sGUseAffSys = "Y" Then
        rbcLogType(4).Visible = True
    Else
        rbcLogType(4).Visible = False
    End If


    'cbcFile.AddItem "Report"
    'cbcFile.AddItem "Fixed Column Width"
    'cbcFile.AddItem "Comma-Separated with Quotes"
    'cbcFile.AddItem "Tab-Separated with Quotes"
    'cbcFile.AddItem "Tab-Separated w/o Quotes"
    'cbcFile.AddItem "DIF"
    'cbcFile.AddItem "Rich Text"

    'If (Trim$(tgUrf(0).sPDFDrvChar) <> "") And (tgUrf(0).iPDFDnArrowCnt >= 0) And (Trim$(tgUrf(0).sPrtDrvChar) <> "") And (tgUrf(0).iPrtDnArrowCnt >= 0) Then
    '    cbcFile.AddItem "Acrobat PDF"
    'End If
    ilRet = gPopExportTypes(cbcFile)    '10-19-01

    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        ilRet = mObtainSplitReplacments()
    Else
        ReDim tmRBofRec(0 To 0) As SPLITBOFREC
        ReDim tmSplitNetLastFill(0 To 0) As SPLITNETLASTFILL
    End If
    mSetCommands        'Set Commands
    If UBound(tgLogSel) <= LBound(tgLogSel) Then
        cmcCancel.SetFocus
    End If
    mPopListKey
    imcKey.Picture = IconTraf!imcKey.Picture
    pbcTab.Left = -pbcTab.Width - 100
    pbcSTab.Left = -pbcSTab.Width - 100
    pbcClickFocus.Left = -pbcClickFocus.Width - 100
    '12/18/14: Colors not shown on log screen replaced bt task monitor
    'If tgSpf.sGUseAffSys = "Y" Then
    '    imcKey.Left = 0
    '    imcKey.Top = plcScreen.Top + plcScreen.Height + 120
    '    lbcKey.Left = 0
    'Else
    '    imcKey.Visible = False
    'End If
    imcKey.Visible = False
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub

    ilRet = 1
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long

    On Error GoTo mInitBoxErr

    lacInfo(1).Visible = False
    lacInfo(2).Visible = False
    plcInfo.Move 120, 255, 9165, 285    '705
    If tgSpf.sCBlackoutLog = "Y" Then
        cmcBlackout.Visible = True
        cmcGenerate.Move 2415, 5655
        cmcCancel.Move cmcGenerate.Left + cmcGenerate.Width + 120, cmcGenerate.Top
        cmcLogChk.Move cmcCancel.Left + cmcCancel.Width + 120, cmcGenerate.Top
        cmcBlackout.Move cmcLogChk.Left + cmcLogChk.Width + 120, cmcGenerate.Top
        If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
            cmcSplitFill.Visible = True
            cmcSplitFill.Move cmcBlackout.Left + cmcBlackout.Width + 120, cmcGenerate.Top
        Else
            cmcSplitFill.Visible = False
        End If
    Else
        cmcBlackout.Visible = False
        If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
            cmcSplitFill.Visible = True
            cmcGenerate.Move 2415, 5655
            cmcCancel.Move cmcGenerate.Left + cmcGenerate.Width + 120, cmcGenerate.Top
            cmcLogChk.Move cmcCancel.Left + cmcCancel.Width + 120, cmcGenerate.Top
            cmcSplitFill.Move cmcLogChk.Left + cmcLogChk.Width + 120, cmcGenerate.Top
        Else
            cmcGenerate.Move 2940, 5655
            cmcCancel.Move cmcGenerate.Left + cmcGenerate.Width + 120, cmcGenerate.Top
            cmcLogChk.Move cmcCancel.Left + cmcCancel.Width + 120, cmcGenerate.Top
            cmcSplitFill.Visible = False
        End If
    End If
    flTextHeight = pbcLogs(imPbcIndex).TextHeight("1") - 35
    plcLogInfo.Move 585, 45
    'Position panel and picture areas with panel
    'plcLogs.Move 135, 1035, pbcLogs(0).Width + vbcLogs.Width + fgPanelAdj, pbcLogs(0).Height + fgPanelAdj
    'pbcLogs(0).Move plcLogs.Left + fgBevelX, plcLogs.Top + fgBevelY
    'pbcLogs(1).Move plcLogs.Left + fgBevelX, plcLogs.Top + fgBevelY
    'vbcLogs.Move pbcLogs(0).Left + pbcLogs(0).Width, pbcLogs(0).Top
    plcLogs.Move 135, 1035
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop

    'Control Fields
    'Check
    gSetCtrl tmCtrls(CHKINDEX), 30, 375, 375, fgBoxGridH
    'Working Date
    gSetCtrl tmCtrls(WRKDATEINDEX), 420, tmCtrls(CHKINDEX).fBoxY, 705, fgBoxGridH
    'Vehicle
    gSetCtrl tmCtrls(VEHINDEX), 1140, tmCtrls(CHKINDEX).fBoxY, 1830, fgBoxGridH
    'Last Log Date
    gSetCtrl tmCtrls(LLDINDEX), 2985, tmCtrls(CHKINDEX).fBoxY, 705, fgBoxGridH
    'Lead Time
    gSetCtrl tmCtrls(LEADTIMEINDEX), 3705, tmCtrls(CHKINDEX).fBoxY, 435, fgBoxGridH
    'Cycle
    gSetCtrl tmCtrls(CYCLEINDEX), 4155, tmCtrls(CHKINDEX).fBoxY, 435, fgBoxGridH
    'Start Date
    gSetCtrl tmCtrls(SDATEINDEX), 4605, tmCtrls(CHKINDEX).fBoxY, 705, fgBoxGridH
    'End Date
    gSetCtrl tmCtrls(EDATEINDEX), 5325, tmCtrls(CHKINDEX).fBoxY, 705, fgBoxGridH
    'Log Name
    gSetCtrl tmCtrls(LOGINDEX), 6045, tmCtrls(CHKINDEX).fBoxY, 540, fgBoxGridH
    'CP Name
    gSetCtrl tmCtrls(CPINDEX), 6600, tmCtrls(CHKINDEX).fBoxY, 540, fgBoxGridH
    'CP Logo
    gSetCtrl tmCtrls(LOGOINDEX), 7155, tmCtrls(CHKINDEX).fBoxY, 540, fgBoxGridH
    'Play List
    gSetCtrl tmCtrls(OTHERINDEX), 7710, tmCtrls(CHKINDEX).fBoxY, 540, fgBoxGridH
    'Zone
    gSetCtrl tmCtrls(ZONEINDEX), 8265, tmCtrls(CHKINDEX).fBoxY, 465, fgBoxGridH

    'edcInfo.Move plcLogs.Left, plcLogs.Top + plcLogs.Height + 30
    edcInfo.Move 90, 15
    pbcRptSample(0).Move edcInfo.Left + edcInfo.Width, edcInfo.Top, 6630    'plcLogs.Left + plcLogs.Width - (edcInfo.Left + edcInfo.Width + 15)

    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxW)
        Do While (tmCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmCtrls(ilLoop).fBoxW = tmCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxX)
            Do While (tmCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmCtrls(ilLoop).fBoxX = tmCtrls(ilLoop).fBoxX + 1
            Loop
            If tmCtrls(ilLoop).fBoxX > 90 Then
                Do
                    If tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 < tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 > tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    pbcLogs(0).Picture = LoadPicture("")
    pbcLogs(1).Picture = LoadPicture("")
    pbcLogs(0).Width = llMax
    pbcLogs(1).Width = llMax
    plcLogs.Width = llMax + vbcLogs.Width + 2 * fgBevelX + 15
    lacFrame(0).Width = llMax - 15
    lacFrame(1).Width = lacFrame(0).Width
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcGenerate.Left + cmcGenerate.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    cmcGenerate.Left = (Logs.Width - cmcGenerate.Width - cmcCancel.Width - cmcLogChk.Width - cmcBlackout.Width - cmcSplitFill.Width - 4 * ilSpaceBetweenButtons) / 2
    cmcCancel.Left = cmcGenerate.Left + cmcGenerate.Width + ilSpaceBetweenButtons
    cmcLogChk.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    cmcBlackout.Left = cmcLogChk.Left + cmcLogChk.Width + ilSpaceBetweenButtons
    cmcSplitFill.Left = cmcBlackout.Left + cmcBlackout.Width + ilSpaceBetweenButtons
    cmcGenerate.Top = Logs.height - (3 * cmcGenerate.height) / 2
    cmcCancel.Top = cmcGenerate.Top
    cmcLogChk.Top = cmcGenerate.Top
    cmcBlackout.Top = cmcGenerate.Top
    cmcSplitFill.Top = cmcGenerate.Top
    ckcCheckOn.Top = cmcGenerate.Top
    plcStatus.Top = cmcGenerate.Top - plcStatus.height - 30
    llAdjTop = plcStatus.Top - plcLogs.Top - 120 - tmCtrls(1).fBoxH
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    Do While plcLogs.Top + llAdjTop + 2 * fgBevelY + 240 < plcStatus.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    plcLogs.height = llAdjTop + 2 * fgBevelY
    pbcLogs(0).Left = plcLogs.Left + fgBevelX
    pbcLogs(0).Top = plcLogs.Top + fgBevelY
    pbcLogs(0).height = plcLogs.height - 2 * fgBevelY
    pbcLogs(1).Left = pbcLogs(0).Left
    pbcLogs(1).Top = pbcLogs(0).Top
    pbcLogs(1).height = pbcLogs(0).height
    vbcLogs.Left = pbcLogs(0).Left + pbcLogs(0).Width + 15
    vbcLogs.Top = pbcLogs(0).Top
    vbcLogs.height = pbcLogs(0).height

    Exit Sub
mInitBoxErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLogPop                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan     *
'*                                                     *
'*            Comments: Populate Log Vehicles, Times,. *
'*                                                     *
'*******************************************************
Private Sub mLogPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer            'return status
    Dim llFilter As Long    'btrieve filter
    Dim slStr As String             'g-user vehicle name string
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llWrkDate As Long
    Dim llLLD As Long
    Dim llDate As Long
    Dim slWrkSort As String
    Dim ilRnf As Integer
    Dim ilList As Integer
    Dim ilPos As Integer
    Dim ilVef As Integer
    Dim llTempSDate As Long
    Dim ilLink As Integer
    Dim ilFound As Integer
    Dim slStartDate As String
    Dim slField(0 To 4) As String
    ReDim tgLogSel(0 To 0) As LOGSEL
    ReDim tgLogChkMsg(0 To 0) As LOGSEL

    'dan M  6-17-08 repop zone.
    lbcTimeZ.Clear
    lbcTimeZ.AddItem "[All]"
    lbcTimeZ.AddItem "EST"
    lbcTimeZ.AddItem "CST"
    lbcTimeZ.AddItem "MST"
    lbcTimeZ.AddItem "PST"

    'Populate vehicle list box
    sgVpfStamp = "~"    'Force read
    ilRet = gVpfRead()
    'Note: VEHSPORT will be added in mTestVehType
    If tgSpf.sMktBase = "Y" Then
        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH + VEHBYMKT + VEHBYPASSNOLOG + VEHTESTLOGMERGE + VEHEXCLUDEPODNOPRGM ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 Exclude CPM Vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
        ilRet = gPopUserVehicleByMkt(Logs, llFilter, igLogMktCode(), lbcVehicle, tgLogVehicle(), sgLogVehicleTag)
    Else
        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH + VEHBYPASSNOLOG + VEHTESTLOGMERGE + VEHEXCLUDEPODNOPRGM  ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 Exclude CPM Vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
        ilRet = gPopUserVehicleBox(Logs, llFilter, lbcVehicle, tgLogVehicle(), sgLogVehicleTag)
    End If
    ''ilFilter = 10   ' Airing, selling and all conventional vehicles (except with Log) and Log
    ''ilRet = gPopUserVehicleBox(Logs, ilFilter, lbcVehicle, lbcVehCode)
    'Moved within if above
    'ilRet = gPopUserVehicleBox(Logs, ilFilter, lbcVehicle, tgVehicle(), sgVehicleTag)
    ''If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLogPopErr
        gCPErrorMsg ilRet, "mLogPop (gPopUserVehicleBox)", Logs
        On Error GoTo 0
        'Add last log date
        If lbcVehicle.ListCount > 0 Then
            ReDim tgLogSel(0 To lbcVehicle.ListCount) As LOGSEL
        End If
        For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
            slNameCode = tgLogVehicle(ilLoop).sKey 'lbcVehCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgLogSel(ilLoop).iStatus = 0
            tgLogSel(ilLoop).iVefCode = Val(slCode)
            If bgLogFirstCallToVpfFind Then
                tgLogSel(ilLoop).iVpfIndex = gVpfFind(Logs, tgLogSel(ilLoop).iVefCode)
                bgLogFirstCallToVpfFind = False
            Else
                tgLogSel(ilLoop).iVpfIndex = gVpfFindIndex(tgLogSel(ilLoop).iVefCode)
            End If
            gUnpackDate tgVpf(tgLogSel(ilLoop).iVpfIndex).iLLD(0), tgVpf(tgLogSel(ilLoop).iVpfIndex).iLLD(1), tgLogSel(ilLoop).sLLD
            If Trim$(tgLogSel(ilLoop).sLLD) <> "" Then
                llLLD = gDateValue(Trim$(tgLogSel(ilLoop).sLLD))
                tgLogSel(ilLoop).lStartDate = llLLD + 1
                tgLogSel(ilLoop).iLLDChgAllowed = False
            Else
                tgLogSel(ilLoop).lStartDate = gDateValue(Format$(gNow(), "m/d/yy")) + 1
                tgLogSel(ilLoop).iLLDChgAllowed = True
            End If
            If tgVpf(tgLogSel(ilLoop).iVpfIndex).iLNoDaysCycle > 0 Then
                tgLogSel(ilLoop).iCycle = tgVpf(tgLogSel(ilLoop).iVpfIndex).iLNoDaysCycle
            Else
                tgLogSel(ilLoop).iCycle = 1
            End If
            tgLogSel(ilLoop).lEndDate = tgLogSel(ilLoop).lStartDate + tgLogSel(ilLoop).iCycle - 1
            tmVefSrchKey.iCode = tgLogSel(ilLoop).iVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            tgLogSel(ilLoop).sVehicle = Trim$(tmVef.sName)
            If tmVef.sType = "L" Then
                ilFound = False
                llTempSDate = 0
                tgLogSel(ilLoop).iStatus = 3
                For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '7/27/12: Include Sports within Log vehicles
                    'If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                    If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "A")) And (tgMVef(ilVef).iVefCode = tmVef.iCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                        tgLogSel(ilLoop).iStatus = 0
                        If tgMVef(ilVef).sType = "C" Then
                            llDate = mObtainClosingDate(tgMVef(ilVef).iCode, tgLogSel(ilLoop).lStartDate, tgLogSel(ilLoop).lEndDate, True)
                        ElseIf tgMVef(ilVef).sType = "G" Then
                            llDate = mObtainGameClosingDate(tgMVef(ilVef).iCode, tgLogSel(ilLoop).lStartDate, tgLogSel(ilLoop).lEndDate)
                        ElseIf tgMVef(ilVef).sType = "A" Then
                            llDate = mObtainAiringClosingDate(tgMVef(ilVef), ilLoop, ilFound)
                        End If
                        If llDate > 0 Then
                            If llTempSDate = 0 Then
                                llTempSDate = llDate
                            Else
                                If llDate < llTempSDate Then
                                    llTempSDate = llDate
                                End If
                            End If
                        ElseIf llDate = 0 Then
                            ilFound = True
                            Exit For
                        End If
                    End If
                Next ilVef
                If Not ilFound Then
                    If llTempSDate = 0 Then
                        tgLogSel(ilLoop).lStartDate = 0
                        tgLogSel(ilLoop).lEndDate = 0
                        tgLogSel(ilLoop).iStatus = 4
                    Else
                        tgLogSel(ilLoop).lStartDate = llTempSDate
                        tgLogSel(ilLoop).lEndDate = tgLogSel(ilLoop).lStartDate + tgLogSel(ilLoop).iCycle - 1
                    End If
                End If
            ElseIf tmVef.sType = "A" Then
                If (tmVef.iVefCode <= 0) Or (Not mMergeWithLog(tmVef.iCode)) Then
                    'If Trim$(tgLogSel(ilLoop).sLLD) = "" Then
                    '    slStartdate = Format$(gNow(), "m/d/yy")
                    'Else
                    '    slStartdate = Trim$(tgLogSel(ilLoop).sLLD)
                    'End If
                    ''slStartDate = gIncOneWeek(slStartDate)
                    'slStartdate = gIncOneDay(slStartdate)
                    'gBuildLinkArray hmVLF, tmVef, slStartdate, igSVefCode()
                    'If (UBound(igSVefCode) <= LBound(igSVefCode)) And (Trim$(tgLogSel(ilLoop).sLLD) = "") Then
                    '    tmVlfSrchKey1.iAirCode = tmVef.iCode
                    '    tmVlfSrchKey1.iAirDay = 0
                    '    gPackDate slStartdate, tmVlfSrchKey1.iEffDate(0), tmVlfSrchKey1.iEffDate(1)
                    '    tmVlfSrchKey1.iAirTime(0) = 0
                    '    tmVlfSrchKey1.iAirTime(1) = 0
                    '    tmVlfSrchKey1.iAirPosNo = 0
                    '    ilRet = btrGetGreaterOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    '    If (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = tmVef.iCode) Then
                    '        gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slStartdate
                    '        gBuildLinkArray hmVLF, tmVef, slStartdate, igSVefCode()
                    '    End If
                    'End If
                    'If (UBound(igSVefCode) <= LBound(igSVefCode)) Then
                    '    tgLogSel(ilLoop).iStatus = 1
                    'End If
                    'ilFound = False
                    'llTempSDate = 0
                    ''For ilLink = LBound(tgVpf(tgLogSel(ilLoop).iVpfIndex).iGLink) To UBound(tgVpf(tgLogSel(ilLoop).iVpfIndex).iGLink) Step 1
                    ''    If tgVpf(tgLogSel(ilLoop).iVpfIndex).iGLink(ilLink) > 0 Then
                    'For ilLink = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                    '        'llDate = mObtainClosingDate(tgVpf(tgLogSel(ilLoop).iVpfIndex).iGLink(ilLink), tgLogSel(ilLoop).lStartDate, tgLogSel(ilLoop).lEndDate, True)
                    '        llDate = mObtainClosingDate(igSVefCode(ilLink), tgLogSel(ilLoop).lStartDate, tgLogSel(ilLoop).lEndDate, True)
                    '        If llDate > 0 Then
                    '            If llTempSDate = 0 Then
                    '                llTempSDate = llDate
                    '            Else
                    '                If llDate < llTempSDate Then
                    '                    llTempSDate = llDate
                    '                End If
                    '            End If
                    '        ElseIf llDate = 0 Then
                    '            ilFound = True
                    '            Exit For
                    '        End If
                    ''    End If
                    'Next ilLink
                    llTempSDate = mObtainAiringClosingDate(tmVef, ilLoop, ilFound)

                    If Not ilFound Then
                        If llTempSDate = 0 Then
                            tgLogSel(ilLoop).lStartDate = 0
                            tgLogSel(ilLoop).lEndDate = 0
                            tgLogSel(ilLoop).iStatus = 2
                        Else
                            tgLogSel(ilLoop).lStartDate = llTempSDate
                            tgLogSel(ilLoop).lEndDate = tgLogSel(ilLoop).lStartDate + tgLogSel(ilLoop).iCycle - 1
                        End If
                    End If
                End If
            ElseIf tmVef.sType = "G" Then
                llDate = mObtainGameClosingDate(tmVef.iCode, tgLogSel(ilLoop).lStartDate, tgLogSel(ilLoop).lEndDate)
                If llDate > 0 Then
                    tgLogSel(ilLoop).lStartDate = llDate
                    tgLogSel(ilLoop).lEndDate = tgLogSel(ilLoop).lStartDate + tgLogSel(ilLoop).iCycle - 1
                ElseIf llDate < 0 Then
                    tgLogSel(ilLoop).lStartDate = 0
                    tgLogSel(ilLoop).lEndDate = 0
                    tgLogSel(ilLoop).iStatus = 5
                End If
            Else
                'Test start date
                llDate = mObtainClosingDate(tmVef.iCode, tgLogSel(ilLoop).lStartDate, tgLogSel(ilLoop).lEndDate, True)
                If llDate > 0 Then
                    tgLogSel(ilLoop).lStartDate = llDate
                    tgLogSel(ilLoop).lEndDate = tgLogSel(ilLoop).lStartDate + tgLogSel(ilLoop).iCycle - 1
                ElseIf llDate < 0 Then
                    tgLogSel(ilLoop).lStartDate = 0
                    tgLogSel(ilLoop).lEndDate = 0
                    tgLogSel(ilLoop).iStatus = 5
                End If
            End If
            If tgVpf(tgLogSel(ilLoop).iVpfIndex).iLLeadTime <= 0 Then
                tgVpf(tgLogSel(ilLoop).iVpfIndex).iLLeadTime = 1
            End If
            llWrkDate = tgLogSel(ilLoop).lStartDate - tgVpf(tgLogSel(ilLoop).iVpfIndex).iLLeadTime
            If (llWrkDate > 0) And (tgLogSel(ilLoop).iStatus = 0) Then
                tgLogSel(ilLoop).sWrkDate = Format$(llWrkDate, "m/d/yy")
                slWrkSort = Trim$(str$(llWrkDate))
                Do While Len(slWrkSort) < 6
                    slWrkSort = "0" & slWrkSort
                Loop
            Else
                tgLogSel(ilLoop).sWrkDate = ""
                slWrkSort = "999999"
            End If
            tgLogSel(ilLoop).iLeadTime = tgVpf(tgLogSel(ilLoop).iVpfIndex).iLLeadTime
            tgLogSel(ilLoop).iLog = 0
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If tmRnfList(ilRnf).tRnf.iCode = tgVpf(tgLogSel(ilLoop).iVpfIndex).iRnfLogCode Then
                    For ilList = 0 To lbcLog.ListCount - 1 Step 1
                        If lbcLog.List(ilList) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            tgLogSel(ilLoop).iLog = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
            tgLogSel(ilLoop).iCP = 0
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If tmRnfList(ilRnf).tRnf.iCode = tgVpf(tgLogSel(ilLoop).iVpfIndex).iRnfCertCode Then
                    For ilList = 0 To lbcCP.ListCount - 1 Step 1
                        If lbcCP.List(ilList) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            tgLogSel(ilLoop).iCP = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
            tgLogSel(ilLoop).iLogo = 0
            For ilList = 0 To lbcLogo.ListCount - 1 Step 1
                slStr = lbcLogo.List(ilList)
                ilPos = InStr(slStr, ".")
                If ilPos > 0 Then
                    slStr = Left$(slStr, ilPos)
                End If
                slStr = UCase$(slStr)
                If slStr = "G" & UCase$(Trim$(tgVpf(tgLogSel(ilLoop).iVpfIndex).sCPLogo)) Then
                    tgLogSel(ilLoop).iLogo = ilList
                    Exit For
                End If
            Next ilList
            tgLogSel(ilLoop).iOther = 0
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If tmRnfList(ilRnf).tRnf.iCode = tgVpf(tgLogSel(ilLoop).iVpfIndex).iRnfPlayCode Then
                    For ilList = 0 To lbcOther.ListCount - 1 Step 1
                        If lbcOther.List(ilList) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            tgLogSel(ilLoop).iOther = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
            tgLogSel(ilLoop).iZone = 0
            For ilList = 0 To lbcTimeZ.ListCount - 1 Step 1
                slStr = lbcTimeZ.List(ilList)
                slStr = Left$(slStr, 1)
                slStr = UCase$(slStr)
                If slStr = UCase$(Trim$(tgVpf(tgLogSel(ilLoop).iVpfIndex).slZone)) Then
                    tgLogSel(ilLoop).iZone = ilList
                    Exit For
                End If
            Next ilList
            If tgLogSel(ilLoop).iStatus = 0 Then
                If llWrkDate <= lmNowDate Then
                    tgLogSel(ilLoop).iChk = 1
                Else
                    tgLogSel(ilLoop).iChk = 0
                End If
            Else
                tgLogSel(ilLoop).iChk = 0
            End If
            tgLogSel(ilLoop).iInitChk = tgLogSel(ilLoop).iChk
            tgLogSel(ilLoop).iChg = False
            tgLogSel(ilLoop).sKey = slWrkSort & "|" & slNameCode
        Next ilLoop
        If UBound(tgLogSel) - 1 > 0 Then
            ArraySortTyp fnAV(tgLogSel(), 0), UBound(tgLogSel), 0, LenB(tgLogSel(0)), 0, LenB(tgLogSel(0).sKey), 0
        End If
        For ilLoop = 0 To UBound(tgLogSel) - 1 Step 1
            If tgLogSel(ilLoop).lStartDate = 0 Then
                tgLogChkMsg(UBound(tgLogChkMsg)) = tgLogSel(ilLoop)
                ReDim Preserve tgLogChkMsg(0 To UBound(tgLogChkMsg) + 1) As LOGSEL
                'Exit For
            End If
        Next ilLoop
        For ilLoop = 0 To UBound(tgLogSel) - 1 Step 1
            If tgLogSel(ilLoop).lStartDate = 0 Then
                ReDim Preserve tgLogSel(0 To ilLoop) As LOGSEL
                Exit For
            End If
        Next ilLoop
    'End If
    If imCurSort = 1 Then
        For ilLoop = LBound(tgLogSel) To UBound(tgLogSel) - 1 Step 1
            ilRet = gParseItemNoTrim(tgLogSel(ilLoop).sKey, 1, "|", slField(0))
            ilRet = gParseItemNoTrim(tgLogSel(ilLoop).sKey, 2, "|", slField(1))
            ilRet = gParseItemNoTrim(tgLogSel(ilLoop).sKey, 3, "|", slField(2))
            ilRet = gParseItemNoTrim(tgLogSel(ilLoop).sKey, 4, "|", slStr)
            ilRet = gParseItemNoTrim(slStr, 1, "\", slField(3))
            ilRet = gParseItemNoTrim(slStr, 2, "\", slField(4))
            tgLogSel(ilLoop).sKey = slField(3) & "|" & slField(1) & "|" & slField(2) & "|" & slField(0) & "\" & slField(4)
        Next ilLoop
        If UBound(tgLogSel) - 1 > 0 Then
            ArraySortTyp fnAV(tgLogSel(), 0), UBound(tgLogSel), 0, LenB(tgLogSel(0)), 0, LenB(tgLogSel(0).sKey), 0
        End If
    End If
    ReDim tgSel(0 To UBound(tgLogSel)) As LOGSEL
    For ilLoop = 0 To UBound(tgLogSel) Step 1
        tgSel(ilLoop) = tgLogSel(ilLoop)
    Next ilLoop
    vbcLogs.Value = vbcLogs.Min
    If UBound(tgLogSel) <= vbcLogs.LargeChange + 1 Then
        vbcLogs.Max = vbcLogs.Min
    Else
        vbcLogs.Max = UBound(tgLogSel) - vbcLogs.LargeChange + 1    'Show one extra line
    End If
    Exit Sub
mLogPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainClosingDate              *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Determine next starting closing*
'*                      date                           *
'*                                                     *
'*******************************************************
Private Function mObtainClosingDate(ilVefCode As Integer, llSTestDate As Long, llETestDate As Long, ilTestNow As Integer) As Long
    Dim ilType As Integer
    Dim llDate As Long
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer
    ilType = 0
    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
    tmSsfSrchKey.iType = ilType
    tmSsfSrchKey.iVefCode = ilVefCode
    gPackDateLong llSTestDate, ilDate0, ilDate1
    tmSsfSrchKey.iDate(0) = ilDate0
    tmSsfSrchKey.iDate(1) = ilDate1
    tmSsfSrchKey.iStartTime(0) = 0
    tmSsfSrchKey.iStartTime(1) = 0
    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode) Then
        gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
        If (llDate >= llSTestDate) And (llDate <= llETestDate) Then
            If (llDate <= lmNowDate) And (ilTestNow) Then
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                tmSsfSrchKey.iType = ilType
                tmSsfSrchKey.iVefCode = ilVefCode
                gPackDateLong lmNowDate + 1, ilDate0, ilDate1
                tmSsfSrchKey.iDate(0) = ilDate0
                tmSsfSrchKey.iDate(1) = ilDate1
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode) Then
                    mObtainClosingDate = 0
                Else
                    mObtainClosingDate = -1
                End If
            Else
                mObtainClosingDate = 0
            End If
        Else
            If (llDate <= lmNowDate) And (ilTestNow) Then
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                tmSsfSrchKey.iType = ilType
                tmSsfSrchKey.iVefCode = ilVefCode
                gPackDateLong lmNowDate + 1, ilDate0, ilDate1
                tmSsfSrchKey.iDate(0) = ilDate0
                tmSsfSrchKey.iDate(1) = ilDate1
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode) Then
                    mObtainClosingDate = llDate
                Else
                    mObtainClosingDate = -1
                End If
            Else
                mObtainClosingDate = llDate
            End If
        End If
    Else
        mObtainClosingDate = -1
    End If
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
Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    'On Error GoTo mOpenMsgFileErr:
    ''slToFile = sgExportPath & "Logs.Txt"
    ''slToFile = "c:\csi\" & "Logs.Txt"
    slToFile = sgDBPath & "Messages\" & "Logs" & CStr(tgUrf(0).iCode) & ".Txt"
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
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
            ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
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
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "Log Generation: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    Print #hmMsg, ""
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mLogPop                         *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Log, CP and Other     *
'*                      list boxes                     *
'*                                                     *
'*******************************************************
Private Sub mRnfPop()
    Dim ilLoop As Integer
    Dim ilLen As Integer
    Dim slChar As String
    Dim ilOk As Integer
    Dim ilValue As Integer
    Dim slName As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    Dim slDateTime1 As String
    Dim slDateTime2 As String
    gObtainRNF hmRnf
    lbcLog.Clear
    lbcCP.Clear
    lbcOther.Clear
    'Move tgRnfList to tmRnfList because if Report button selected, tgRnfList could be changed
    ReDim tmRnfList(0 To UBound(tgRnfList)) As RNFLIST
    For ilLoop = 0 To UBound(tgRnfList) - 1 Step 1
        tmRnfList(ilLoop) = tgRnfList(ilLoop)
    Next ilLoop
    For ilLoop = 0 To UBound(tmRnfList) - 1 Step 1
        If tmRnfList(ilLoop).tRnf.sType = "R" Then
            ilLen = Len(Trim$(tmRnfList(ilLoop).tRnf.sName))
            slChar = UCase$(Left$(tmRnfList(ilLoop).tRnf.sName, 1))
            If (ilLen > 1) And (ilLen < 4) Then
                ilOk = False
                If slChar = "L" Then
                    ilOk = True
                ElseIf slChar = "C" Then
                    ilOk = True
                'ElseIf slChar = "O" Then
                '    ilOk = True
                End If
                If ilOk Then
                    ilValue = Asc(Mid$(tmRnfList(ilLoop).tRnf.sName, 2, 1))
                    If (ilValue < Asc("0")) Or (ilValue > Asc("9")) Then
                        ilOk = False
                    End If
                End If
                If ilOk Then
                    slName = UCase$(Trim$(tmRnfList(ilLoop).tRnf.sName))
                    If slChar = "L" Then
                        lbcLog.AddItem slName
                        lbcOther.AddItem slName
                    ElseIf slChar = "C" Then
                        If StrComp(Left$(slName, 3), "C17", 1) = 0 Then
                            If tgSpf.sGUseAffSys = "Y" Then
                                lbcCP.AddItem slName
                                lbcOther.AddItem slName     '2-16-01
                            End If
                        Else
                            lbcCP.AddItem slName
                            lbcOther.AddItem slName     '2-16-01
                        End If
                    'ElseIf slChar = "O" Then
                    '    lbcOther.AddItem slName
                    End If
                End If
            End If
        End If
    Next ilLoop
    lbcLog.AddItem "[None]", 0
    lbcCP.AddItem "[None]", 0
    lbcOther.AddItem "[None]", 0
    lbcLogo.Clear
    On Error GoTo mRnfErr
    lbcFile.Path = Left$(sgLogoPath, Len(sgLogoPath) - 1)
'8-18-14 no longer copy all the logos to root path
'    If lbcFile.ListCount - 1 >= 0 Then
'        'Move CP logo to local C drice (c:\csi\G???.bmp)
'        ilRet = 0
'        On Error GoTo mLogoErr:
'        '5676 remove hardcoded c:
'        'slDateTime1 = FileDateTime("C:\CSI\RptLogo.Bmp")
'        slDateTime1 = FileDateTime(sgRootDrive & "CSI\RptLogo.Bmp")
'        If ilRet <> 0 Then
'            ilRet = 0
'            'MkDir "C:\CSI"
'            MkDir sgRootDrive & "CSI"
'        End If
'        If ilRet = 0 Then
'            For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
'                ilRet = 0
'                slStr = lbcFile.List(ilLoop)
'                'slDateTime1 = FileDateTime("C:\CSI\" & slStr)
'                slDateTime1 = FileDateTime(sgRootDrive & "CSI\" & slStr)
'                If ilRet = 0 Then
'                    slDateTime2 = FileDateTime(sgLogoPath & slStr)
'                    If ilRet = 0 Then
'                        If StrComp(slDateTime1, slDateTime2, 1) <> 0 Then
'                            'FileCopy sgLogoPath & slStr, "C:\CSI\" & slStr
'                            FileCopy sgLogoPath & slStr, sgRootDrive & "CSI\" & slStr
'                        End If
'                    End If
'                Else
'                    'FileCopy sgLogoPath & slStr, "C:\CSI\" & slStr
'                    FileCopy sgLogoPath & slStr, sgRootDrive & "CSI\" & slStr
'                End If
'            Next ilLoop
'        End If
'    End If

    On Error GoTo mRnfErr
    For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
        slStr = lbcFile.List(ilLoop)
        ilPos = InStr(slStr, ".")
        If ilPos > 0 Then
            slStr = Left$(slStr, ilPos - 1)
        End If
        lbcLogo.AddItem slStr
    Next ilLoop
    lbcLogo.AddItem "[None]", 0
    On Error GoTo 0
    Exit Sub
mRnfErr:
    Resume Next
mLogoErr:
    ilRet = err.Number
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRPPop                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan     *
'*                                                     *
'*            Comments: Populate Vehicles, Times,.. for*
'*                      Reprint of Logs                *
'*                                                     *
'*******************************************************
Private Sub mRPPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer            'return status
    Dim llFilter As Long    'btrieve filter
    Dim slStr As String             'g-user vehicle name string
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llWrkDate As Long
    Dim llLLD As Long
    Dim slWrkSort As String
    Dim ilRnf As Integer
    Dim ilList As Integer
    Dim ilPos As Integer
    Dim slField(0 To 4) As String

    If imRPGen Then
        Exit Sub
    End If
    imRPGen = True
    ReDim tgRPSel(0 To 0) As LOGSEL
    'Populate vehicle list box
    If tgSpf.sMktBase = "Y" Then
        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH + VEHBYMKT + VEHBYPASSNOLOG + VEHTESTLOGMERGE ' Airing and all conventional vehicles (except with Log) and Log
        '12/29/05-  Added byMky call instead of the gPopUserVehicleBox which was showing all vehicles
        ilRet = gPopUserVehicleByMkt(Logs, llFilter, igLogMktCode(), lbcVehicle, tgLogVehicle(), sgLogVehicleTag)
    Else
        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH + VEHBYPASSNOLOG + VEHTESTLOGMERGE ' Airing and all conventional vehicles (except with Log) and Log
        ilRet = gPopUserVehicleBox(Logs, llFilter, lbcVehicle, tgLogVehicle(), sgLogVehicleTag)
    End If
    ''ilFilter = 10   ' Airing, selling and all conventional vehicles (except with Log) and Log
    ''ilRet = gPopUserVehicleBox(Logs, ilFilter, lbcVehicle, lbcVehCode)
    'ilRet = gPopUserVehicleBox(Logs, llFilter, lbcVehicle, tgLogVehicle(), sgLogVehicleTag)
    ''If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mRPPopErr
        gCPErrorMsg ilRet, "mRPPop (gPopUserVehicleBox)", Logs
        On Error GoTo 0
        'Add last log date
        If lbcVehicle.ListCount > 0 Then
            ReDim tgRPSel(0 To lbcVehicle.ListCount) As LOGSEL
        End If
        For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
            slNameCode = tgLogVehicle(ilLoop).sKey 'lbcVehCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgRPSel(ilLoop).iStatus = 0
            tgRPSel(ilLoop).iVefCode = Val(slCode)
            If bgLogFirstCallToVpfFind Then
                tgRPSel(ilLoop).iVpfIndex = gVpfFind(Logs, tgRPSel(ilLoop).iVefCode)
                bgLogFirstCallToVpfFind = False
            Else
                tgRPSel(ilLoop).iVpfIndex = gVpfFindIndex(tgRPSel(ilLoop).iVefCode)
            End If
            If tgVpf(tgRPSel(ilLoop).iVpfIndex).iLNoDaysCycle > 0 Then
                tgRPSel(ilLoop).iCycle = tgVpf(tgRPSel(ilLoop).iVpfIndex).iLNoDaysCycle
            Else
                tgRPSel(ilLoop).iCycle = 1
            End If
            tgRPSel(ilLoop).iLLDChgAllowed = False
            gUnpackDate tgVpf(tgRPSel(ilLoop).iVpfIndex).iLLD(0), tgVpf(tgRPSel(ilLoop).iVpfIndex).iLLD(1), tgRPSel(ilLoop).sLLD
            If Trim$(tgRPSel(ilLoop).sLLD) <> "" Then
                llLLD = gDateValue(Trim$(tgRPSel(ilLoop).sLLD))
                tgRPSel(ilLoop).lStartDate = llLLD - tgRPSel(ilLoop).iCycle + 1
                tgRPSel(ilLoop).lEndDate = tgRPSel(ilLoop).lStartDate + tgRPSel(ilLoop).iCycle - 1
            Else
                tgRPSel(ilLoop).lStartDate = 0
                tgRPSel(ilLoop).lEndDate = 0
                tgRPSel(ilLoop).iStatus = 6
            End If
            tmVefSrchKey.iCode = tgRPSel(ilLoop).iVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            tgRPSel(ilLoop).sVehicle = Trim$(tmVef.sName)
            If tgVpf(tgRPSel(ilLoop).iVpfIndex).iLLeadTime <= 0 Then
                tgVpf(tgRPSel(ilLoop).iVpfIndex).iLLeadTime = 1
            End If
            llWrkDate = tgRPSel(ilLoop).lStartDate - tgVpf(tgRPSel(ilLoop).iVpfIndex).iLLeadTime
            If (llWrkDate > 0) And (tgRPSel(ilLoop).iStatus = 0) Then
                tgRPSel(ilLoop).sWrkDate = Format$(llWrkDate, "m/d/yy")
                slWrkSort = Trim$(str$(999999 - llWrkDate))
                Do While Len(slWrkSort) < 6
                    slWrkSort = "0" & slWrkSort
                Loop
            Else
                tgRPSel(ilLoop).sWrkDate = ""
                slWrkSort = "999999"
            End If
            tgRPSel(ilLoop).iLeadTime = tgVpf(tgRPSel(ilLoop).iVpfIndex).iLLeadTime
            tgRPSel(ilLoop).iLog = 0
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If tmRnfList(ilRnf).tRnf.iCode = tgVpf(tgRPSel(ilLoop).iVpfIndex).iRnfLogCode Then
                    For ilList = 0 To lbcLog.ListCount - 1 Step 1
                        If lbcLog.List(ilList) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            tgRPSel(ilLoop).iLog = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
            tgRPSel(ilLoop).iCP = 0
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If tmRnfList(ilRnf).tRnf.iCode = tgVpf(tgRPSel(ilLoop).iVpfIndex).iRnfCertCode Then
                    For ilList = 0 To lbcCP.ListCount - 1 Step 1
                        If lbcCP.List(ilList) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            tgRPSel(ilLoop).iCP = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
            tgRPSel(ilLoop).iLogo = 0
            For ilList = 0 To lbcLogo.ListCount - 1 Step 1
                slStr = lbcLogo.List(ilList)
                ilPos = InStr(slStr, ".")
                If ilPos > 0 Then
                    slStr = Left$(slStr, ilPos)
                End If
                slStr = UCase$(slStr)
                If slStr = "G" & UCase$(Trim$(tgVpf(tgRPSel(ilLoop).iVpfIndex).sCPLogo)) Then
                    tgRPSel(ilLoop).iLogo = ilList
                    Exit For
                End If
            Next ilList
            tgRPSel(ilLoop).iOther = 0
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If tmRnfList(ilRnf).tRnf.iCode = tgVpf(tgRPSel(ilLoop).iVpfIndex).iRnfPlayCode Then
                    For ilList = 0 To lbcOther.ListCount - 1 Step 1
                        If lbcOther.List(ilList) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            tgRPSel(ilLoop).iOther = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
            tgRPSel(ilLoop).iZone = 0
            For ilList = 0 To lbcTimeZ.ListCount - 1 Step 1
                slStr = lbcTimeZ.List(ilList)
                slStr = Left$(slStr, 1)
                slStr = UCase$(slStr)
                If slStr = UCase$(Trim$(tgVpf(tgRPSel(ilLoop).iVpfIndex).slZone)) Then
                    tgRPSel(ilLoop).iZone = ilList
                    Exit For
                End If
            Next ilList
            tgRPSel(ilLoop).iChk = 0
            tgRPSel(ilLoop).iInitChk = tgRPSel(ilLoop).iChk
            tgRPSel(ilLoop).sKey = slWrkSort & "|" & slNameCode
        Next ilLoop
        If UBound(tgRPSel) - 1 > 0 Then
            ArraySortTyp fnAV(tgRPSel(), 0), UBound(tgRPSel), 0, LenB(tgRPSel(0)), 0, LenB(tgRPSel(0).sKey), 0
        End If
        For ilLoop = 0 To UBound(tgRPSel) - 1 Step 1
            If tgRPSel(ilLoop).lStartDate = 0 Then
                ReDim Preserve tgRPSel(0 To ilLoop) As LOGSEL
                Exit For
            End If
        Next ilLoop
        If imCurSort = 1 Then
            For ilLoop = LBound(tgRPSel) To UBound(tgRPSel) - 1 Step 1
                ilRet = gParseItemNoTrim(tgRPSel(ilLoop).sKey, 1, "|", slField(0))
                ilRet = gParseItemNoTrim(tgRPSel(ilLoop).sKey, 2, "|", slField(1))
                ilRet = gParseItemNoTrim(tgRPSel(ilLoop).sKey, 3, "|", slField(2))
                ilRet = gParseItemNoTrim(tgRPSel(ilLoop).sKey, 4, "|", slStr)
                ilRet = gParseItemNoTrim(slStr, 1, "\", slField(3))
                ilRet = gParseItemNoTrim(slStr, 2, "\", slField(4))
                tgRPSel(ilLoop).sKey = slField(3) & "|" & slField(1) & "|" & slField(2) & "|" & slField(0) & "\" & slField(4)
            Next ilLoop
            If UBound(tgRPSel) - 1 > 0 Then
                ArraySortTyp fnAV(tgRPSel(), 0), UBound(tgRPSel), 0, LenB(tgRPSel(0)), 0, LenB(tgRPSel(0).sKey), 0
            End If
        End If
    'End If
    'If UBound(tgRPSel) <= vbcLogs.LargeChange + 1 Then
    '    vbcLogs.Max = vbcLogs.Min
    'Else
    '    vbcLogs.Max = UBound(tgRPSel) - vbcLogs.LargeChange
    'End If

    Exit Sub
mRPPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan     *
'*                                                     *
'*            Comments: Save User defined values       *
'*                                                     *
'*******************************************************
Private Sub mSaveRec()
    Dim ilRet As Integer
    Dim ilRnf As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    '9/16/15: Allow last row to be saved
    'If (imRowNo <= 0) Or (imRowNo >= UBound(tgSel)) Then
    If (imRowNo <= 0) Or (imRowNo > UBound(tgSel)) Then
        Exit Sub
    End If
    If rbcLogType(2).Value Then
        Exit Sub
    End If
    If rbcLogType(3).Value Then
        Exit Sub
    End If
    If rbcLogType(4).Value Then
        Exit Sub
    End If
    If Not tgSel(imRowNo - 1).iChg Then
        Exit Sub
    End If
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    Do
        tmVpfSrchKey.iVefKCode = tgSel(imRowNo - 1).iVefCode
        ilRet = btrGetEqual(hmVpf, tgVpf(tgSel(imRowNo - 1).iVpfIndex), imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            If tgSel(imRowNo - 1).iLLDChgAllowed Then
                gPackDate tgSel(imRowNo - 1).sLLD, tgVpf(tgSel(imRowNo - 1).iVpfIndex).iLLD(0), tgVpf(tgSel(imRowNo - 1).iVpfIndex).iLLD(1)
                'gPackDate slSyncDate, tgVpf(tgSel(imRowNo - 1).iVpfIndex).iSyncDate(0), tgVpf(tgSel(imRowNo - 1).iVpfIndex).iSyncDate(1)
                'gPackTime slSyncTime, tgVpf(tgSel(imRowNo - 1).iVpfIndex).iSyncTime(0), tgVpf(tgSel(imRowNo - 1).iVpfIndex).iSyncTime(1)
                ''tgVpf(tgSel(imRowNo - 1).iVpfIndex).iSourceID = tgUrf(0).iRemoteUserID
            End If
            tgVpf(tgSel(imRowNo - 1).iVpfIndex).iLNoDaysCycle = tgSel(imRowNo - 1).iCycle
            tgVpf(tgSel(imRowNo - 1).iVpfIndex).iLLeadTime = tgSel(imRowNo - 1).iLeadTime
            tgVpf(tgSel(imRowNo - 1).iVpfIndex).iRnfLogCode = 0
            If tgSel(imRowNo - 1).iLog > 0 Then
                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                    If lbcLog.List(tgSel(imRowNo - 1).iLog) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                        tgVpf(tgSel(imRowNo - 1).iVpfIndex).iRnfLogCode = tmRnfList(ilRnf).tRnf.iCode
                        Exit For
                    End If
                Next ilRnf
            End If
            tgVpf(tgSel(imRowNo - 1).iVpfIndex).iRnfCertCode = 0
            If tgSel(imRowNo - 1).iCP > 0 Then
                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                    If lbcCP.List(tgSel(imRowNo - 1).iCP) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                        tgVpf(tgSel(imRowNo - 1).iVpfIndex).iRnfCertCode = tmRnfList(ilRnf).tRnf.iCode
                        Exit For
                    End If
                Next ilRnf
            End If
            tgVpf(tgSel(imRowNo - 1).iVpfIndex).sCPLogo = ""
            If tgSel(imRowNo - 1).iLogo > 0 Then
                slStr = lbcLogo.List(tgSel(imRowNo - 1).iLogo)
                ilPos = InStr(slStr, ".")
                If ilPos > 0 Then
                    slStr = Left$(slStr, ilPos)
                End If
                tgVpf(tgSel(imRowNo - 1).iVpfIndex).sCPLogo = Mid$(slStr, 2)    'Remove "G"
            End If
            tgVpf(tgSel(imRowNo - 1).iVpfIndex).iRnfPlayCode = 0
            If tgSel(imRowNo - 1).iOther > 0 Then
                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                    If lbcOther.List(tgSel(imRowNo - 1).iOther) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                        tgVpf(tgSel(imRowNo - 1).iVpfIndex).iRnfPlayCode = tmRnfList(ilRnf).tRnf.iCode
                        Exit For
                    End If
                Next ilRnf
            End If
            If tgSel(imRowNo - 1).iZone <= 0 Then
                tgVpf(tgSel(imRowNo - 1).iVpfIndex).slZone = ""
            Else
                Select Case tgSel(imRowNo - 1).iZone
                    Case 1
                        tgVpf(tgSel(imRowNo - 1).iVpfIndex).slZone = "E"
                    Case 2
                        tgVpf(tgSel(imRowNo - 1).iVpfIndex).slZone = "C"
                    Case 3
                        tgVpf(tgSel(imRowNo - 1).iVpfIndex).slZone = "M"
                    Case 4
                        tgVpf(tgSel(imRowNo - 1).iVpfIndex).slZone = "P"
                    Case Else
                        tgVpf(tgSel(imRowNo - 1).iVpfIndex).slZone = ""
                End Select
            End If
        Else
            Exit Do
        End If
        ilRet = btrUpdate(hmVpf, tgVpf(tgSel(imRowNo - 1).iVpfIndex), imVpfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    tgSel(imRowNo - 1).iChg = False
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, "mSaveRec (btrUpdate Or btrGetEqual)", Logs
    '11/26/17
    gFileChgdUpdate "vef.btr", True
    gFileChgdUpdate "vpf.btr", True
    
    On Error GoTo 0
    Exit Sub
mSaveRecErr:
    imTerminate = True
    On Error GoTo 0
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:4/27/94       By:D. Hannifan    *
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
    'ilOneSelected = False
    'For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
    '    If lbcVehicle.Selected(ilLoop) Then
    '        ilOneSelected = True
    '        Exit For
    '    End If
    'Next ilLoop
    'If Not ilOneSelected Then
    '    cmcGenerate.Enabled = False
    '    cmcGenOnly.Enabled = False
    '    cmcUnsold.Enabled = False
    '    Exit Sub
    'End If
    ''Check all mandatory control fields
    'If mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = NO Then
    '    cmcGenerate.Enabled = False
    '    cmcGenOnly.Enabled = False
    '    If (smSave(2) = "") Or (smSave(3) = "") Then
    '        cmcUnsold.Enabled = False
    '    Else
    '        cmcUnsold.Enabled = True
    '    End If
    '    Exit Sub
    'End If
    'cmcGenerate.Enabled = True
    'cmcGenOnly.Enabled = True
    'cmcUnsold.Enabled = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Set focus specified control    *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If


    Select Case ilBoxNo 'Branch on box type (control)
        Case CHKINDEX
            pbcSelections.Visible = True
            pbcSelections.SetFocus
        Case LLDINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            plcCalendar.Visible = True
            edcDropDown.SetFocus
        Case LEADTIMEINDEX
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case CYCLEINDEX
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case SDATEINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            plcCalendar.Visible = True
            edcDropDown.SetFocus
        Case LOGINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case CPINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case LOGOINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case OTHERINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ZONEINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String     'show string
    Dim llWrkDate As Long
    Dim ilIndex As Integer
    pbcArrow.Visible = False
    lacFrame(0).Visible = False
    lacFrame(1).Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case CHKINDEX
            pbcSelections.Visible = False
        Case LLDINDEX
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            plcCalendar.Visible = False
            slStr = edcDropDown.Text
            If gValidDate(slStr) Then
                If (gDateValue(tgSel(imRowNo - 1).sLLD) <> gDateValue(slStr)) And (tgSel(imRowNo - 1).iLLDChgAllowed) Then
                    tgSel(imRowNo - 1).iChg = True
                End If
                tgSel(imRowNo - 1).sLLD = slStr
            End If
            tgSel(imRowNo - 1).lStartDate = gDateValue(tgSel(imRowNo - 1).sLLD) + 1
            tgSel(imRowNo - 1).lEndDate = tgSel(imRowNo - 1).lStartDate + tgLogSel(imRowNo - 1).iCycle - 1
            llWrkDate = tgSel(imRowNo - 1).lStartDate - tgSel(imRowNo - 1).iLeadTime
            tgSel(imRowNo - 1).sWrkDate = Format$(llWrkDate, "m/d/yy")
        Case LEADTIMEINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If tgSel(imRowNo - 1).iLeadTime <> Val(slStr) Then
                tgSel(imRowNo - 1).iChg = True
            End If
            tgSel(imRowNo - 1).iLeadTime = Val(slStr)
            llWrkDate = tgSel(imRowNo - 1).lStartDate - tgSel(imRowNo - 1).iLeadTime
            tgSel(imRowNo - 1).sWrkDate = Format$(llWrkDate, "m/d/yy")
        Case CYCLEINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If tgSel(imRowNo - 1).iCycle <> Val(slStr) Then
                tgSel(imRowNo - 1).iChg = True
            End If
            tgSel(imRowNo - 1).iCycle = Val(slStr)
            tgSel(imRowNo - 1).lEndDate = tgSel(imRowNo - 1).lStartDate + tgSel(imRowNo - 1).iCycle - 1
        Case SDATEINDEX
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            plcCalendar.Visible = False
            slStr = edcDropDown.Text
            If gValidDate(slStr) Then
                tgSel(imRowNo - 1).lStartDate = gDateValue(slStr)
            End If
            tgSel(imRowNo - 1).lEndDate = tgSel(imRowNo - 1).lStartDate + tgSel(imRowNo - 1).iCycle - 1
            llWrkDate = tgSel(imRowNo - 1).lStartDate - tgSel(imRowNo - 1).iLeadTime
            tgSel(imRowNo - 1).sWrkDate = Format$(llWrkDate, "m/d/yy")
        Case LOGINDEX
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            lbcLog.Visible = False
            If tgSel(imRowNo - 1).iLog <> lbcLog.ListIndex Then
                tgSel(imRowNo - 1).iChg = True
            End If
            tgSel(imRowNo - 1).iLog = lbcLog.ListIndex
        Case CPINDEX
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            lbcCP.Visible = False
            If tgSel(imRowNo - 1).iCP <> lbcCP.ListIndex Then
                tgSel(imRowNo - 1).iChg = True
            End If
            tgSel(imRowNo - 1).iCP = lbcCP.ListIndex
        Case LOGOINDEX
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            lbcLogo.Visible = False
            If tgSel(imRowNo - 1).iLogo <> lbcLogo.ListIndex Then
                tgSel(imRowNo - 1).iChg = True
            End If
            tgSel(imRowNo - 1).iLogo = lbcLogo.ListIndex
        Case OTHERINDEX
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            lbcOther.Visible = False
            If tgSel(imRowNo - 1).iOther <> lbcOther.ListIndex Then
                tgSel(imRowNo - 1).iChg = True
            End If
            tgSel(imRowNo - 1).iOther = lbcOther.ListIndex
        Case ZONEINDEX
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            lbcTimeZ.Visible = False
            ilIndex = 0
            slStr = Trim$(lbcTimeZ.List(lbcTimeZ.ListIndex))
            Select Case slStr
                Case "EST"
                    ilIndex = 1
                Case "CST"
                    ilIndex = 2
                Case "MST"
                    ilIndex = 3
                Case "PST"
                    ilIndex = 4
            End Select
            If tgSel(imRowNo - 1).iZone <> ilIndex Then
                tgSel(imRowNo - 1).iChg = True
            End If
            tgSel(imRowNo - 1).iZone = ilIndex
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mShowInfo                       *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show information for           *
'*                      right mouse                    *
'*                                                     *
'*******************************************************
Private Sub mShowInfo()
    Dim ilButtonIndex As Integer
    Dim slStr As String
    ilButtonIndex = imButtonRow '- 1
    If (imButtonRow < LBound(tgSel)) Or (imButtonRow > UBound(tgSel) - 1) Then
        plcInfo.Visible = False
        Exit Sub
    End If
    slStr = "Vehicle Name " & Trim$(tgSel(ilButtonIndex).sVehicle)
    lacInfo(0).Caption = slStr
    'lacInfo(1).Caption = ""
    'lacInfo(2).Caption = ""
    If (imButtonRow < LBound(tgSel)) Or (imButtonRow > UBound(tgSel) - 1) Then
        plcInfo.Visible = False
        Exit Sub
    End If
    plcInfo.ZOrder vbBringToFront
    plcInfo.Visible = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
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
    On Error Resume Next
    


    imTerminate = False
    Screen.MousePointer = vbDefault
    'Unload form
    igManUnload = YES
    Unload Logs
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mUpdateLLDForBypaasedVeh        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: set Last Log date for vehicle  *
'*                      which no logs are generated    *
'*                                                     *
'*******************************************************
Private Sub mUpdateLLDForBypaasedVeh(slSyncDate As String, slSyncTime As String)
    Dim ilVef As Integer
    Dim llDate As Long
    Dim ilRet As Integer
    If rbcLogType(0).Value Or rbcLogType(1).Value Then
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            Do
                tmVpfSrchKey.iVefKCode = tgMVef(ilVef).iCode
                ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If (ilRet = BTRV_ERR_NONE) And (tmVpf.sGenLog = "N") Then
                    gUnpackDateLong tmVpf.iLLD(0), tmVpf.iLLD(1), llDate
                    If lmNowDate > llDate And rbcLogType(1).Value Then
                        gPackDate smNowDate, tmVpf.iLLD(0), tmVpf.iLLD(1)
                        'gPackDate slSyncDate, tmVpf.iSyncDate(0), tmVpf.iSyncDate(1)
                        'gPackTime slSyncTime, tmVpf.iSyncTime(0), tmVpf.iSyncTime(1)
                    End If
                    gPackDate smNowDate, tmVpf.iLPD(0), tmVpf.iLPD(1)
                    ilRet = btrUpdate(hmVpf, tmVpf, imVpfRecLen)
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            ilRet = gBinarySearchVpf(tmVpf.iVefKCode)
            If ilRet <> -1 Then
                tgVpf(ilRet) = tmVpf
            End If
        Next ilVef
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
                edcDropDown.Text = Format$(llDate, "m/d/yy")
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                imBypassFocus = True
                edcDropDown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate

    edcDropDown.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    plcTme.Visible = False
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcLogs_GotFocus(Index As Integer)
    plcTme.Visible = False
End Sub
Private Sub pbcLogs_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    imButton = Button
    If Button = 2 Then  'Right Mouse
        ilCompRow = vbcLogs.LargeChange + 1
        'If UBound(smSave, 2) - 1 > ilCompRow Then
        If UBound(tgSel) - 1 > ilCompRow Then
            'If UBound(smSave, 2) = vbcLogs.Value + ilCompRow - 1 Then
            If UBound(tgSel) = vbcLogs.Value + ilCompRow - 1 Then
                ilMaxRow = ilCompRow - 1
            Else
                ilMaxRow = ilCompRow
            End If
        Else
            ilMaxRow = UBound(tgSel) '- 1   'UBound(smSave, 2) - 1
        End If
        ' Look through all rows
        For ilRow = 1 To ilMaxRow Step 1
            If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(1).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(1).fBoxY + tmCtrls(1).fBoxH)) Then
                imButtonRow = ilRow + vbcLogs.Value - 2
                imIgnoreRightMove = True
                mShowInfo
                imIgnoreRightMove = False
                Exit For
            End If
        Next ilRow
        Exit Sub
    End If
End Sub
Private Sub pbcLogs_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        ilCompRow = vbcLogs.LargeChange + 1
        If UBound(tgSel) - 1 > ilCompRow Then
            'If UBound(smSave, 2) = vbcLogs.Value + ilCompRow - 1 Then
            If UBound(tgSel) = vbcLogs.Value + ilCompRow - 1 Then
                ilMaxRow = ilCompRow - 1
            Else
                ilMaxRow = ilCompRow
            End If
        Else
            ilMaxRow = UBound(tgSel) '- 1   'UBound(smSave, 2) - 1
        End If
        For ilRow = 1 To ilMaxRow Step 1
            If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(VEHINDEX).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(VEHINDEX).fBoxY + tmCtrls(VEHINDEX).fBoxH)) Then
                If (imButtonRow = ilRow + vbcLogs.Value - 2) And (plcInfo.Visible) Then
                    Exit Sub
                End If
                imButtonRow = ilRow + vbcLogs.Value - 2
                imIgnoreRightMove = True
                mShowInfo
                imIgnoreRightMove = False
                Exit Sub
            End If
        Next ilRow
        plcInfo.Visible = False
        Exit Sub
    End If
End Sub
Private Sub pbcLogs_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If Button = 2 Then
        imButtonIndex = -1
        plcInfo.Visible = False
        Exit Sub
    End If
    'Check if sort field selected
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If ((X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW))) Then
            If (Y < tmCtrls(ilBox).fBoxY - 15) Then
                If ((ilBox = WRKDATEINDEX) And (imCurSort = 1)) Or ((ilBox = VEHINDEX) And (imCurSort = 0)) Then
                    plcTme.Visible = False
                    mSetShow imBoxNo
                    mSaveRec
                    imBoxNo = -1
                    imRowNo = -1
                    mResortLog
                    If imCurSort = 1 Then
                        imCurSort = 0
                    Else
                        imCurSort = 1
                    End If
                End If
                Exit Sub
            End If
        End If
    Next ilBox
    ilCompRow = vbcLogs.LargeChange + 1
    If UBound(tgSel) >= ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tgSel)
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcLogs.Value - 1
                    If ilRowNo > UBound(tgSel) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If tgSel(ilRowNo - 1).iStatus <> 0 Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If (ilBox = WRKDATEINDEX) Or (ilBox = VEHINDEX) Or (ilBox = EDATEINDEX) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If (ilBox = LLDINDEX) And (Not tgSel(ilRowNo - 1).iLLDChgAllowed) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    'If (ilBox = LOGOINDEX) And (tgSel(ilRowNo - 1).iCP <= 0) Then     '8-24-01 allow any log to have a customized log
                    '    Beep
                    '    mSetFocus imBoxNo
                    '    Exit Sub
                    'End If
                    If (ilBox = SDATEINDEX) And (imPbcIndex = 1) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    mSetShow imBoxNo
                    If ilRowNo <> imRowNo Then
                        mSaveRec
                    End If
                    imRowNo = ilRow + vbcLogs.Value - 1
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSetFocus imBoxNo
End Sub
Private Sub pbcLogs_Paint(Index As Integer)
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim slFont As String
    Dim ilRet As Integer

    If imGettingMkt Then
        Exit Sub
    End If
    mPaintLogTitle Index
    ilStartRow = vbcLogs.Value '+ 1  'Top location
    ilEndRow = vbcLogs.Value + vbcLogs.LargeChange ' + 1
    ilRet = 0
    On Error GoTo pbcLogsErr
    ilRow = LBound(tgSel)
    On Error GoTo 0
    If ilRet = 1 Then
        Exit Sub
    End If
    If ilEndRow - 1 >= UBound(tgSel) Then
        ilEndRow = UBound(tgSel)
    End If
    slFont = pbcLogs(Index).FontName
    For ilRow = ilStartRow To ilEndRow Step 1
        If tgSel(ilRow - 1).iStatus = 0 Then
            For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
                pbcLogs(Index).CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcLogs(Index).CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                Select Case ilBox
                    Case CHKINDEX
                        pbcLogs(Index).CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) + 15 '+ fgBoxInsetY
                        pbcLogs(Index).FontName = "Monotype Sorts"
                        pbcLogs(Index).FontBold = False
                        If tgSel(ilRow - 1).iChk = 0 Then
                            slStr = "  "
                        Else
                            slStr = "4"
                        End If
                    Case WRKDATEINDEX
                        slStr = tgSel(ilRow - 1).sWrkDate
                    Case VEHINDEX
                        slStr = tgSel(ilRow - 1).sVehicle
                    Case LLDINDEX
                        slStr = tgSel(ilRow - 1).sLLD
                    Case LEADTIMEINDEX
                        slStr = Trim$(str$(tgSel(ilRow - 1).iLeadTime))
                    Case CYCLEINDEX
                        slStr = Trim$(str$(tgSel(ilRow - 1).iCycle))
                    Case SDATEINDEX
                        slStr = Format$(tgSel(ilRow - 1).lStartDate, "m/d/yy")
                    Case EDATEINDEX
                        slStr = Format$(tgSel(ilRow - 1).lEndDate, "m/d/yy")
                    Case LOGINDEX
                        slStr = lbcLog.List(tgSel(ilRow - 1).iLog)
                    Case CPINDEX
                        slStr = lbcCP.List(tgSel(ilRow - 1).iCP)
                    Case LOGOINDEX
                        slStr = lbcLogo.List(tgSel(ilRow - 1).iLogo)
                    Case OTHERINDEX
                        slStr = lbcOther.List(tgSel(ilRow - 1).iOther)
                    Case ZONEINDEX
                        'slStr = lbcTimeZ.List(tgSel(ilRow - 1).iZone)
                        slStr = ""
                        Select Case tgSel(ilRow - 1).iZone
                            Case 0
                                slStr = "[All]"
                            Case 1
                                slStr = "EST"
                            Case 2
                                slStr = "CST"
                            Case 3
                                slStr = "MST"
                            Case 4
                                slStr = "PST"
                        End Select
                End Select
                gSetShow pbcLogs(Index), slStr, tmCtrls(ilBox)
                pbcLogs(Index).Print tmCtrls(ilBox).sShow
                If ilBox = CHKINDEX Then
                    pbcLogs(Index).FontName = slFont
                    pbcLogs(Index).FontBold = True
                End If
            Next ilBox
        Else
            pbcLogs(Index).CurrentX = tmCtrls(VEHINDEX).fBoxX + fgBoxInsetX
            pbcLogs(Index).CurrentY = tmCtrls(VEHINDEX).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            slStr = tgSel(ilRow - 1).sVehicle
            gSetShow pbcLogs(Index), slStr, tmCtrls(VEHINDEX)
            pbcLogs(Index).Print tmCtrls(VEHINDEX).sShow
            pbcLogs(Index).CurrentX = tmCtrls(LLDINDEX).fBoxX + fgBoxInsetX
            pbcLogs(Index).CurrentY = tmCtrls(LLDINDEX).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            Select Case tgSel(ilRow - 1).iStatus
                Case 1
                    slStr = "No Selling to Airing Links"
                Case 2
                    slStr = "No Dates in Future"
                Case 3
                    slStr = "(Log) No Conventional"
                Case 4
                    slStr = "(Log) Conventional No Dates in Future"
                Case 5
                    slStr = "No Dates in Future"
            End Select
            pbcLogs(Index).Print slStr
        End If
    Next ilRow
    Exit Sub
pbcLogsErr:
    ilRet = 1
    Resume Next
End Sub
Private Sub pbcSelections_GotFocus()
    If imFirstTime Then
        imFirstTime = False
        pbcSelections_Paint
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcSelections_KeyPress(KeyAscii As Integer)
    If imBoxNo = CHKINDEX Then   'Log or commercial format
        If KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            tgSel(imRowNo - 1).iChk = 0
            pbcSelections_Paint
        ElseIf KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            tgSel(imRowNo - 1).iChk = 1
            pbcSelections_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If (tgSel(imRowNo - 1).iChk = 0) Then
                tgSel(imRowNo - 1).iChk = 1
                pbcSelections_Paint
            ElseIf (tgSel(imRowNo - 1).iChk = 1) Then
                tgSel(imRowNo - 1).iChk = 0
                pbcSelections_Paint
            End If
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcSelections_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = CHKINDEX Then  'Log or commercial format
        If (tgSel(imRowNo - 1).iChk = 0) Then
            tgSel(imRowNo - 1).iChk = 1
            pbcSelections_Paint
        ElseIf (tgSel(imRowNo - 1).iChk = 1) Then
            tgSel(imRowNo - 1).iChk = 0
            pbcSelections_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcSelections_Paint()
    pbcSelections.Cls
    pbcSelections.CurrentX = fgBoxInsetX
    pbcSelections.CurrentY = 30 'fgBoxInsetY
    If imBoxNo = CHKINDEX Then     'Log or commercial format
        If tgSel(imRowNo - 1).iChk = 0 Then
            pbcSelections.Print "  "   '"Log Format"
        ElseIf tgSel(imRowNo - 1).iChk = 1 Then
            pbcSelections.Print "4" '"Delivery Format"
        Else
            pbcSelections.Print "   "
        End If
    End If
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer    'control index
    Dim ilFound As Integer  'loop exit flag
    Dim slDate As String    'date string
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = -1  'Set-right to left
    plcTme.Visible = False
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Initial
                If tgSel(0).iStatus <> 0 Then
                    cmcCancel.SetFocus
                    Exit Sub
                End If
                imTabDirection = 0  'Set-Left to right
                imSettingValue = True
                vbcLogs.Value = 1
                imSettingValue = False
                imRowNo = 1
                ilBox = CHKINDEX
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            Case CHKINDEX
                mSetShow imBoxNo
                mSaveRec
                If imRowNo <= 1 Then
                    If cmcGenerate.Enabled Then
                        imBoxNo = -1
                        imRowNo = -1
                        cmcGenerate.SetFocus
                        Exit Sub
                    End If
                    imBoxNo = -1
                    imRowNo = -1
                    cmcGenerate.SetFocus
                Else
                    ilBox = ZONEINDEX
                    imRowNo = imRowNo - 1
                    If imRowNo < vbcLogs.Value Then
                        imSettingValue = True
                        vbcLogs.Value = vbcLogs.Value - 1
                        imSettingValue = False
                    End If
                End If
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            Case LEADTIMEINDEX
                If tgSel(imRowNo - 1).iLLDChgAllowed Then
                    ilBox = LLDINDEX
                Else
                    ilBox = CHKINDEX
                End If
            Case LOGINDEX
                If imPbcIndex = 1 Then
                    ilBox = CYCLEINDEX
                Else
                    ilBox = SDATEINDEX
                End If
            'Case OTHERINDEX            '8-24-01 allow any log/cp to have a customized log
            '    If tgSel(imRowNo - 1).iCP > 0 Then
            '        ilBox = LOGOINDEX
            '    Else
            '        ilBox = CPINDEX
            '    End If
            Case SDATEINDEX       'Start date
                slDate = edcDropDown.Text
                If gValidDate(slDate) Then
                    ilBox = ilBox - 1
                Else                      'Invalid date
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Case Else
                ilBox = ilBox - 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer    'local Control box counter
    Dim ilFound As Integer  'redirect focus flag
    Dim slDate As String    'date string
    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    plcTme.Visible = False
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1          'Initial
                If tgSel(0).iStatus <> 0 Then
                    cmcCancel.SetFocus
                    Exit Sub
                End If
                imTabDirection = -1  'Set-Right to left
                'imRowNo = UBound(tgSel)
                'imSettingValue = True
                'If imRowNo <= vbcLogs.LargeChange + 1 Then
                '    vbcLogs.Value = 1
                'Else
                '    vbcLogs.Value = imRowNo - vbcLogs.LargeChange - 1
                'End If
                'imSettingValue = False
                'ilBox = CHKINDEX
                imSettingValue = True
                vbcLogs.Value = 1
                imSettingValue = False
                imRowNo = 1
                ilBox = CHKINDEX
            Case CHKINDEX
                If tgSel(imRowNo).iStatus <> 0 Then
                    mSetShow imBoxNo
                    If cmcGenerate.Enabled Then
                        cmcGenerate.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                If tgSel(imRowNo).iChk = 0 Then
                    mSetShow imBoxNo
                    If imRowNo >= UBound(tgSel) Then     'UBound(tgBvfRec) Then
                        If cmcGenerate.Enabled Then
                            cmcGenerate.SetFocus
                        Else
                            cmcCancel.SetFocus
                        End If
                        Exit Sub
                    End If
                    imRowNo = imRowNo + 1
                    If imRowNo > vbcLogs.Value + vbcLogs.LargeChange - 1 Then
                        imSettingValue = True
                        vbcLogs.Value = vbcLogs.Value + 1
                        imSettingValue = False
                    End If
                    ilBox = CHKINDEX
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                Else
                    If tgSel(imRowNo - 1).iLLDChgAllowed Then
                        ilBox = LLDINDEX
                    Else
                        ilBox = LEADTIMEINDEX
                    End If
                End If
            Case CYCLEINDEX
                If imPbcIndex = 1 Then
                    ilBox = LOGINDEX
                Else
                    ilBox = SDATEINDEX
                End If
            Case SDATEINDEX      'Start Date
                slDate = edcDropDown.Text
                If gValidDate(slDate) Then
                    ilBox = LOGINDEX
                Else                      'Invalid date
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            'Case CPINDEX         '8-24-01 allow any log/cp to have a customized log
            '    mSetShow imBoxNo
            '    If tgSel(imRowNo - 1).iCP > 0 Then
            '        ilBox = LOGOINDEX
            '    Else
            '        ilBox = OTHERINDEX
            '    End If
            Case ZONEINDEX
                mSetShow imBoxNo
                mSaveRec
                If tgSel(imRowNo).iStatus <> 0 Then
                    If cmcGenerate.Enabled Then
                        cmcGenerate.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                If imRowNo >= UBound(tgSel) Then     'UBound(tgBvfRec) Then
                    If cmcGenerate.Enabled Then
                        cmcGenerate.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                imRowNo = imRowNo + 1
                If imRowNo > vbcLogs.Value + vbcLogs.LargeChange - 1 Then
                    imSettingValue = True
                    vbcLogs.Value = vbcLogs.Value + 1
                    imSettingValue = False
                End If
                ilBox = CHKINDEX
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            Case Else
                ilBox = ilBox + 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
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
                    Select Case imTmeIndex
                        Case 0
                            imBypassFocus = True    'Don't change select text
                            edcTime(0).SetFocus
                            'SendKeys slKey
                            gSendKeys edcTime(0), slKey
                        Case 1
                            imBypassFocus = True    'Don't change select text
                            edcTime(1).SetFocus
                            'SendKeys slKey
                            gSendKeys edcTime(1), slKey
                    End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub plcLogInfo_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcLogs_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcStatus_Paint()
    plcStatus.CurrentX = 0
    plcStatus.CurrentY = 0
    plcStatus.Print smStatusCaption
End Sub

Private Sub rbcLogType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcLogType(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    If Value Then
        Screen.MousePointer = vbHourglass
        '0=Prel; 1=final and 2=reprint
        If (Index = 0) Or (Index = 2) Or (Index = 4) Then  'Preliminary or Reprint or Internal
            imPbcIndex = 0
        Else    'Final or Alert (Disallow start date from being altered)
            imPbcIndex = 1
        End If
        If (Index = 0) Or (Index = 1) Then  'Preliminary or Final
            If Index = 0 Then
                'rbcOutput(0).Enabled = True
                ckcOutput(0).Enabled = True
                pbcLogs(0).Visible = True
                pbcLogs(1).Visible = False
            Else
                'If rbcOutput(0).Caption <> "None" Then
                If tgSpf.sGUseAffSys <> "Y" Then
                    'rbcOutput(1).Value = True   'Select printing
                    ckcOutput(1).Value = vbChecked
                End If
                If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                    'If rbcOutput(0).Caption <> "None" Then
                    '    rbcOutput(0).Enabled = False
                    'Else
                    '    rbcOutput(0).Enabled = True
                    'End If
                    '5/10/12:  Allow display of final log if site set to Y
                    If tgSaf(0).sFinalLogDisplay <> "Y" Then
                        ckcOutput(0).Enabled = False
                        ckcOutput(0).Value = vbUnchecked
                    Else
                        ckcOutput(0).Enabled = True
                    End If
                Else
                    'rbcOutput(0).Enabled = True
                    ckcOutput(0).Enabled = True
                End If
                pbcLogs(1).Visible = True
                pbcLogs(0).Visible = False
            End If
            ckcCheckOn.Visible = True
            If (imLogType = 2) Or (imLogType = 4) Then
                pbcLogs(Index).Cls
                For ilLoop = 0 To UBound(tgSel) Step 1
                    tgRPSel(ilLoop) = tgSel(ilLoop)
                Next ilLoop
            ElseIf imLogType = 3 Then
                pbcLogs(Index).Cls
                For ilLoop = 0 To UBound(tgSel) Step 1
                    tgAlertSel(ilLoop) = tgSel(ilLoop)
                Next ilLoop
            Else
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            ReDim tgSel(0 To UBound(tgLogSel)) As LOGSEL
            For ilLoop = 0 To UBound(tgLogSel) Step 1
                tgSel(ilLoop) = tgLogSel(ilLoop)
            Next ilLoop
        ElseIf (Index = 2) Or (Index = 4) Then
            pbcLogs(0).Visible = True
            pbcLogs(1).Visible = False
            pbcLogs(0).Cls
            If imLogType = 0 Or imLogType = 1 Then
                For ilLoop = 0 To UBound(tgSel) Step 1
                    tgLogSel(ilLoop) = tgSel(ilLoop)
                Next ilLoop
            ElseIf imLogType = 3 Then
                For ilLoop = 0 To UBound(tgSel) Step 1
                    tgAlertSel(ilLoop) = tgSel(ilLoop)
                Next ilLoop
            End If
            If Not imRPGen Then
                Screen.MousePointer = vbHourglass
                mRPPop
                imRPGen = True
                Screen.MousePointer = vbDefault
            End If
            ReDim tgSel(0 To UBound(tgRPSel)) As LOGSEL
            For ilLoop = 0 To UBound(tgRPSel) Step 1
                tgSel(ilLoop) = tgRPSel(ilLoop)
            Next ilLoop
            'If rbcOutput(0).Caption <> "None" Then
            If tgSpf.sGUseAffSys <> "Y" Then
                'rbcOutput(1).Value = True   'Select printing
                ckcOutput(1).Value = vbChecked
            End If
            '1/29/98  Allow reprint to display so Save To is allowed.
            'If (Trim$(tgUrf(0).sName) <> sgCPName) Then
            '    rbcOutput(0).Enabled = False
            'Else
                'rbcOutput(0).Enabled = True
                ckcOutput(0).Enabled = True
            'End If
            ckcCheckOn.Visible = False
        ElseIf Index = 3 Then
            pbcLogs(0).Visible = False
            pbcLogs(1).Visible = True
            pbcLogs(1).Cls
            If imLogType = 0 Or imLogType = 1 Then
                For ilLoop = 0 To UBound(tgSel) Step 1
                    tgLogSel(ilLoop) = tgSel(ilLoop)
                Next ilLoop
            ElseIf (imLogType = 2) Or (imLogType = 4) Then
                For ilLoop = 0 To UBound(tgSel) Step 1
                    tgRPSel(ilLoop) = tgSel(ilLoop)
                Next ilLoop
            End If
            If Not imAlertGen Then
                Screen.MousePointer = vbHourglass
                mAlertPop
                imAlertGen = True
                Screen.MousePointer = vbDefault
            End If
            ReDim tgSel(0 To UBound(tgAlertSel)) As LOGSEL
            For ilLoop = 0 To UBound(tgAlertSel) Step 1
                tgSel(ilLoop) = tgAlertSel(ilLoop)
            Next ilLoop
            'If rbcOutput(0).Caption <> "None" Then
            If tgSpf.sGUseAffSys <> "Y" Then
                'rbcOutput(1).Value = True   'Select printing
                ckcOutput(1).Value = vbChecked
            End If
            '1/29/98  Allow reprint to display so Save To is allowed.
            'If (Trim$(tgUrf(0).sName) <> sgCPName) Then
            '    rbcOutput(0).Enabled = False
            'Else
                'rbcOutput(0).Enabled = True
                ckcOutput(0).Enabled = True
            'End If
            ckcCheckOn.Visible = True
        End If
        If Index = 4 Then
            ckcAssignCopy.Visible = False
        Else
            ckcAssignCopy.Visible = True
        End If
        imLogType = Index
        imSettingValue = True
        vbcLogs.Value = vbcLogs.Min
        imSettingValue = True
        If UBound(tgSel) <= vbcLogs.LargeChange + 1 Then
            vbcLogs.Max = vbcLogs.Min
        Else
            vbcLogs.Max = UBound(tgSel) - vbcLogs.LargeChange + 1   'Show one extra line
        End If
        pbcLogs_Paint imPbcIndex
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub rbcLogType_GotFocus(Index As Integer)
    plcTme.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub



Private Sub vbcLogs_Change()
    If imSettingValue Then
        pbcLogs(imPbcIndex).Cls
        pbcLogs_Paint imPbcIndex
        imSettingValue = False
    Else
        mSetShow imBoxNo
        pbcLogs(imPbcIndex).Cls
        pbcLogs_Paint imPbcIndex
        mEnableBox imBoxNo
    End If
    '2/7/09:  Added to avoid possible error
    On Error Resume Next
    pbcClickFocus.SetFocus
    On Error GoTo 0
End Sub
Private Sub vbcLogs_GotFocus()
    plcTme.Visible = False
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame(0).Visible = False
    lacFrame(1).Visible = False
End Sub
Private Sub plcLogMsg_Paint()
    plcLogMsg.CurrentX = 0
    plcLogMsg.CurrentY = 0
    plcLogMsg.Print "Log Generation Errors"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Logs"
End Sub

Private Sub mResortLog()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim slField(0 To 4) As String


    pbcLogs(0).Cls
    pbcLogs(1).Cls
    ilRet = 0
    On Error GoTo mResortLogErr
    If imRPGen Then
        If ilRet = 0 Then
            If (imLogType = 2) Or (imLogType = 4) Then
                For ilLoop = 0 To UBound(tgSel) Step 1
                    tgRPSel(ilLoop) = tgSel(ilLoop)
                Next ilLoop
            End If
            For ilLoop = LBound(tgRPSel) To UBound(tgRPSel) - 1 Step 1
                ilRet = gParseItemNoTrim(tgRPSel(ilLoop).sKey, 1, "|", slField(0))
                ilRet = gParseItemNoTrim(tgRPSel(ilLoop).sKey, 2, "|", slField(1))
                ilRet = gParseItemNoTrim(tgRPSel(ilLoop).sKey, 3, "|", slField(2))
                ilRet = gParseItemNoTrim(tgRPSel(ilLoop).sKey, 4, "|", slStr)
                ilRet = gParseItemNoTrim(slStr, 1, "\", slField(3))
                ilRet = gParseItemNoTrim(slStr, 2, "\", slField(4))
                tgRPSel(ilLoop).sKey = slField(3) & "|" & slField(1) & "|" & slField(2) & "|" & slField(0) & "\" & slField(4)
            Next ilLoop
            If UBound(tgRPSel) - 1 > 0 Then
                ArraySortTyp fnAV(tgRPSel(), 0), UBound(tgRPSel), 0, LenB(tgRPSel(0)), 0, LenB(tgRPSel(0).sKey), 0
            End If
        End If
    End If
    If imAlertGen Then
        If ilRet = 0 Then
            If imLogType = 3 Then
                For ilLoop = 0 To UBound(tgSel) Step 1
                    tgAlertSel(ilLoop) = tgSel(ilLoop)
                Next ilLoop
            End If
            For ilLoop = LBound(tgAlertSel) To UBound(tgAlertSel) - 1 Step 1
                ilRet = gParseItemNoTrim(tgAlertSel(ilLoop).sKey, 1, "|", slField(0))
                ilRet = gParseItemNoTrim(tgAlertSel(ilLoop).sKey, 2, "|", slField(1))
                ilRet = gParseItemNoTrim(tgAlertSel(ilLoop).sKey, 3, "|", slField(2))
                ilRet = gParseItemNoTrim(tgAlertSel(ilLoop).sKey, 4, "|", slStr)
                ilRet = gParseItemNoTrim(slStr, 1, "\", slField(3))
                ilRet = gParseItemNoTrim(slStr, 2, "\", slField(4))
                tgAlertSel(ilLoop).sKey = slField(3) & "|" & slField(1) & "|" & slField(2) & "|" & slField(0) & "\" & slField(4)
            Next ilLoop
            If UBound(tgAlertSel) - 1 > 0 Then
                ArraySortTyp fnAV(tgAlertSel(), 0), UBound(tgAlertSel), 0, LenB(tgAlertSel(0)), 0, LenB(tgAlertSel(0).sKey), 0
            End If
        End If
    End If
    If (imLogType <> 2) And (imLogType <> 3) And (imLogType <> 4) Then
        For ilLoop = 0 To UBound(tgSel) Step 1
            tgLogSel(ilLoop) = tgSel(ilLoop)
        Next ilLoop
    End If
    For ilLoop = LBound(tgLogSel) To UBound(tgLogSel) - 1 Step 1
        ilRet = gParseItemNoTrim(tgLogSel(ilLoop).sKey, 1, "|", slField(0))
        ilRet = gParseItemNoTrim(tgLogSel(ilLoop).sKey, 2, "|", slField(1))
        ilRet = gParseItemNoTrim(tgLogSel(ilLoop).sKey, 3, "|", slField(2))
        ilRet = gParseItemNoTrim(tgLogSel(ilLoop).sKey, 4, "|", slStr)
        ilRet = gParseItemNoTrim(slStr, 1, "\", slField(3))
        ilRet = gParseItemNoTrim(slStr, 2, "\", slField(4))
        tgLogSel(ilLoop).sKey = slField(3) & "|" & slField(1) & "|" & slField(2) & "|" & slField(0) & "\" & slField(4)
    Next ilLoop
    If UBound(tgLogSel) - 1 > 0 Then
        ArraySortTyp fnAV(tgLogSel(), 0), UBound(tgLogSel), 0, LenB(tgLogSel(0)), 0, LenB(tgLogSel(0).sKey), 0
    End If
    If (imLogType = 2) Or (imLogType = 4) Then
        For ilLoop = 0 To UBound(tgRPSel) Step 1
            tgSel(ilLoop) = tgRPSel(ilLoop)
        Next ilLoop
    ElseIf imLogType = 3 Then
        For ilLoop = 0 To UBound(tgAlertSel) Step 1
            tgSel(ilLoop) = tgAlertSel(ilLoop)
        Next ilLoop
    Else
        For ilLoop = 0 To UBound(tgLogSel) Step 1
            tgSel(ilLoop) = tgLogSel(ilLoop)
        Next ilLoop
    End If
    pbcLogs_Paint imPbcIndex
    Exit Sub
mResortLogErr:
    ilRet = 1
    Resume Next
End Sub

Private Function mSetGenLST(ilSelIndex As Integer, slMonStartDate As String, ilExportType As Integer) As Integer
    Dim llLoop1 As Long
    Dim ilCycle As Integer
    Dim llCycleDate As Long
    Dim llDate1 As Long
    Dim llDate2 As Long
    Dim ilGenLST As Integer

    ilGenLST = False
    ilExportType = 0
    For llLoop1 = LBound(tmVATT) To UBound(tmVATT) - 1 Step 1
        ilCycle = tgSel(ilSelIndex).lEndDate - tgSel(ilSelIndex).lStartDate + 1
        llCycleDate = gDateValue(slMonStartDate)
        Do
            tmAtt = tmVATT(llLoop1)
            gUnpackDateLong tmAtt.iOnAir(0), tmAtt.iOnAir(1), llDate1
            Do While gWeekDayLong(llDate1) <> 0
                llDate1 = llDate1 - 1
            Loop
            gUnpackDateLong tmAtt.iOffAir(0), tmAtt.iOffAir(1), llDate2
            If llDate2 <> 0 Then
                Do While gWeekDayLong(llDate2) <> 6
                    llDate2 = llDate2 + 1
                Loop
            End If
            If (llCycleDate >= llDate1) And ((llCycleDate <= llDate2) Or (llDate2 = 0)) Then
                gUnpackDateLong tmAtt.iDropDate(0), tmAtt.iDropDate(1), llDate1
                If llDate1 <> 0 Then
                    Do While gWeekDayLong(llDate1) <> 6
                        llDate1 = llDate1 + 1
                    Loop
                End If
                If ((llCycleDate <= llDate1) Or (llDate1 = 0)) Then
                    'If tmAtt.iPostingType <> 0 Then
                        ilGenLST = True
'                        Exit For
                    'End If
                    If tmAtt.iExportType > 0 Then
                        ilExportType = tmAtt.iExportType
                        Exit For
                    End If
                    Exit Do
                End If
            End If
            ilCycle = ilCycle - 7
            llCycleDate = llCycleDate + 7
        'End If
        Loop While ilCycle > 0
    Next llLoop1
    mSetGenLST = ilGenLST
End Function

Private Sub mAlertPop()
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilVef As Integer
    Dim llWrkDate As Long
    Dim slWrkSort As String
    Dim ilRnf As Integer
    Dim ilList As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    Dim slType As String
    Dim slNameCode As String
    Dim slField(0 To 4) As String

    If imAlertGen Then
        Exit Sub
    End If
    imAlertGen = True
    gAlertVehicleReplace hmVLF
    ReDim tgAlertSel(0 To 0) As LOGSEL
    ilUpper = UBound(tgAlertSel)
    imAufRecLen = Len(tmAuf)
    slType = "L"
    tmAufSrchKey1.sType = slType
    tmAufSrchKey1.sStatus = "R"
    ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (slType = tmAuf.sType) And (tmAuf.sStatus = "R")

        ilVef = gBinarySearchVef(tmAuf.iVefCode)
        If ilVef <> -1 Then
            ilUpper = UBound(tgAlertSel)
            tgAlertSel(ilUpper).iStatus = 0     'Disallow any change
            tgAlertSel(ilUpper).iVefCode = tmAuf.iVefCode
            If bgLogFirstCallToVpfFind Then
                tgAlertSel(ilUpper).iVpfIndex = gVpfFind(Logs, tgAlertSel(ilUpper).iVefCode)
                bgLogFirstCallToVpfFind = False
            Else
                tgAlertSel(ilUpper).iVpfIndex = gVpfFindIndex(tgAlertSel(ilUpper).iVefCode)
            End If
            If tgVpf(tgAlertSel(ilUpper).iVpfIndex).iLNoDaysCycle > 0 Then
                tgAlertSel(ilUpper).iCycle = tgVpf(tgAlertSel(ilUpper).iVpfIndex).iLNoDaysCycle
            Else
                tgAlertSel(ilUpper).iCycle = 1
            End If
            tgAlertSel(ilUpper).iLLDChgAllowed = False
            gUnpackDate tgVpf(tgAlertSel(ilUpper).iVpfIndex).iLLD(0), tgVpf(tgAlertSel(ilUpper).iVpfIndex).iLLD(1), tgAlertSel(ilUpper).sLLD
'            If Trim$(tgAlertSel(ilUpper).sLLD) <> "" Then
'                llLLD = gDateValue(Trim$(tgAlertSel(ilUpper).sLLD))
'                tgAlertSel(ilUpper).lStartDate = llLLD - tgAlertSel(ilUpper).iCycle + 1
'                tgAlertSel(ilUpper).lEndDate = tgAlertSel(ilUpper).lStartDate + tgAlertSel(ilUpper).iCycle - 1
'            Else
'                tgAlertSel(ilUpper).lStartDate = 0
'                tgAlertSel(ilUpper).lEndDate = 0
'                tgAlertSel(ilUpper).iStatus = 6
'            End If
            gUnpackDateLong tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), tgAlertSel(ilUpper).lStartDate
            tgAlertSel(ilUpper).lEndDate = tgAlertSel(ilUpper).lStartDate + tgAlertSel(ilUpper).iCycle - 1
            'tmVefSrchKey.iCode = tgAlertSel(ilUpper).iVefCode
            'ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            tgAlertSel(ilUpper).sVehicle = Trim$(tgMVef(ilVef).sName)
            If tgVpf(tgAlertSel(ilUpper).iVpfIndex).iLLeadTime <= 0 Then
                tgVpf(tgAlertSel(ilUpper).iVpfIndex).iLLeadTime = 1
            End If
            llWrkDate = tgAlertSel(ilUpper).lStartDate - tgVpf(tgAlertSel(ilUpper).iVpfIndex).iLLeadTime
            If (llWrkDate > 0) And (tgAlertSel(ilUpper).iStatus = 0) Then
                tgAlertSel(ilUpper).sWrkDate = Format$(llWrkDate, "m/d/yy")
                slWrkSort = Trim$(str$(999999 - llWrkDate))
                Do While Len(slWrkSort) < 6
                    slWrkSort = "0" & slWrkSort
                Loop
            Else
                tgAlertSel(ilUpper).sWrkDate = ""
                slWrkSort = "999999"
            End If
            tgAlertSel(ilUpper).iLeadTime = tgVpf(tgAlertSel(ilUpper).iVpfIndex).iLLeadTime
            tgAlertSel(ilUpper).iLog = 0
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If tmRnfList(ilRnf).tRnf.iCode = tgVpf(tgAlertSel(ilUpper).iVpfIndex).iRnfLogCode Then
                    For ilList = 0 To lbcLog.ListCount - 1 Step 1
                        If lbcLog.List(ilList) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            tgAlertSel(ilUpper).iLog = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
            tgAlertSel(ilUpper).iCP = 0
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If tmRnfList(ilRnf).tRnf.iCode = tgVpf(tgAlertSel(ilUpper).iVpfIndex).iRnfCertCode Then
                    For ilList = 0 To lbcCP.ListCount - 1 Step 1
                        If lbcCP.List(ilList) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            tgAlertSel(ilUpper).iCP = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
            tgAlertSel(ilUpper).iLogo = 0
            For ilList = 0 To lbcLogo.ListCount - 1 Step 1
                slStr = lbcLogo.List(ilList)
                ilPos = InStr(slStr, ".")
                If ilPos > 0 Then
                    slStr = Left$(slStr, ilPos)
                End If
                slStr = UCase$(slStr)
                If slStr = "G" & UCase$(Trim$(tgVpf(tgAlertSel(ilUpper).iVpfIndex).sCPLogo)) Then
                    tgAlertSel(ilUpper).iLogo = ilList
                    Exit For
                End If
            Next ilList
            tgAlertSel(ilUpper).iOther = 0
            For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                If tmRnfList(ilRnf).tRnf.iCode = tgVpf(tgAlertSel(ilUpper).iVpfIndex).iRnfPlayCode Then
                    For ilList = 0 To lbcOther.ListCount - 1 Step 1
                        If lbcOther.List(ilList) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            tgAlertSel(ilUpper).iOther = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
            tgAlertSel(ilUpper).iZone = 0
            For ilList = 0 To lbcTimeZ.ListCount - 1 Step 1
                slStr = lbcTimeZ.List(ilList)
                slStr = Left$(slStr, 1)
                slStr = UCase$(slStr)
                If slStr = UCase$(Trim$(tgVpf(tgAlertSel(ilUpper).iVpfIndex).slZone)) Then
                    tgAlertSel(ilUpper).iZone = ilList
                    Exit For
                End If
            Next ilList
            slNameCode = "000|000|" & tgMVef(ilVef).sName & "\" & Trim$(str$(tgMVef(ilVef).iCode))
            If ckcCheckOn.Value = vbChecked Then
                tgAlertSel(ilUpper).iChk = 1
            Else
                tgAlertSel(ilUpper).iChk = 0
            End If
            tgAlertSel(ilUpper).iInitChk = 1    'tgAlertSel(ilUpper).iChk
            tgAlertSel(ilUpper).sKey = slWrkSort & "|" & slNameCode
            ilFound = False
            For ilLoop = 0 To UBound(tgAlertSel) - 1 Step 1
                If (tgAlertSel(ilLoop).iVefCode = tgAlertSel(ilUpper).iVefCode) And (tgAlertSel(ilLoop).lStartDate = tgAlertSel(ilUpper).lStartDate) Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                ReDim Preserve tgAlertSel(0 To UBound(tgAlertSel) + 1) As LOGSEL
            End If
        End If
        ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    If UBound(tgAlertSel) - 1 > 0 Then
        ArraySortTyp fnAV(tgAlertSel(), 0), UBound(tgAlertSel), 0, LenB(tgAlertSel(0)), 0, LenB(tgAlertSel(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tgAlertSel) - 1 Step 1
        If tgAlertSel(ilLoop).lStartDate = 0 Then
            ReDim Preserve tgAlertSel(0 To ilLoop) As LOGSEL
            Exit For
        End If
    Next ilLoop
    If imCurSort = 1 Then
        For ilLoop = LBound(tgAlertSel) To UBound(tgAlertSel) - 1 Step 1
            ilRet = gParseItemNoTrim(tgAlertSel(ilLoop).sKey, 1, "|", slField(0))
            ilRet = gParseItemNoTrim(tgAlertSel(ilLoop).sKey, 2, "|", slField(1))
            ilRet = gParseItemNoTrim(tgAlertSel(ilLoop).sKey, 3, "|", slField(2))
            ilRet = gParseItemNoTrim(tgAlertSel(ilLoop).sKey, 4, "|", slStr)
            ilRet = gParseItemNoTrim(slStr, 1, "\", slField(3))
            ilRet = gParseItemNoTrim(slStr, 2, "\", slField(4))
            tgAlertSel(ilLoop).sKey = slField(3) & "|" & slField(1) & "|" & slField(2) & "|" & slField(0) & "\" & slField(4)
        Next ilLoop
        If UBound(tgAlertSel) - 1 > 0 Then
            ArraySortTyp fnAV(tgAlertSel(), 0), UBound(tgAlertSel), 0, LenB(tgAlertSel(0)), 0, LenB(tgAlertSel(0).sKey), 0
        End If
    End If

End Sub

Private Function mBBSpots(ilLoop As Integer) As Integer
    'Dim ilLoop As Integer
    Dim ilRet As Integer

    mBBSpots = True
    If tgSpf.sUsingBBs <> "Y" Then
        Exit Function
    End If
    'Determine vehicles to create Billboard spots
    'For ilLoop = 0 To UBound(tgSel) - 1 Step 1
    '    If tgSel(ilLoop).iChk = 1 Then
            ilRet = gMakeBBAndAssignCopy(hmSdf, hmVLF, tgSel(ilLoop).iVefCode, tgSel(ilLoop).lStartDate, tgSel(ilLoop).lEndDate)
            If Not ilRet Then
                mBBSpots = False
            End If
    '    End If
    'Next ilLoop
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mGhfGsfReadRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mObtainGameClosingDate(ilVefCode As Integer, llSTestDate As Long, llETestDate As Long) As Long
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mObtainGameClosingDateErr                                                             *
'******************************************************************************************

'
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim llTempSDate As Long
    Dim llDate As Long

    llTempSDate = 0
    tmGhfSrchKey1.iVefCode = ilVefCode
    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    'If ilRet = BTRV_ERR_NONE Then
    Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.iVefCode = ilVefCode)
'        tmGsfSrchKey1.lGhfCode = tmGhf.lCode
'        tmGsfSrchKey1.iGameNo = 0
'        ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
'        Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.lCode = tmGsf.lGhfCode)
'            If tmGsf.sGameStatus <> "C" Then
'                gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
'                If (llDate >= llSTestDate) And (llDate <= llETestDate) Then
'                    mObtainGameClosingDate = 0
'                    Exit Function
'                Else
'                    If (llDate > lmNowDate) And (llDate >= llSTestDate) Then
'                        If llTempSDate = 0 Then
'                            llTempSDate = llDate
'                        Else
'                            If llDate < llTempSDate Then
'                                llTempSDate = llDate
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'            ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
'        Loop
        tmGsfSrchKey2.lghfcode = tmGhf.lCode
        gPackDateLong llSTestDate, tmGsfSrchKey2.iAirDate(0), tmGsfSrchKey2.iAirDate(1)
        tmGsfSrchKey2.iAirTime(0) = 0
        tmGsfSrchKey2.iAirTime(1) = 0
        ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.lCode = tmGsf.lghfcode)
            If tmGsf.sGameStatus <> "C" Then
                gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
                If (llDate >= llSTestDate) And (llDate <= llETestDate) Then
                    mObtainGameClosingDate = 0
                    Exit Function
                Else
                    If (llDate > lmNowDate) And (llDate >= llSTestDate) Then
                        'llTempSDate = llDate
                        'Exit Do
                        If llTempSDate = 0 Then
                            llTempSDate = llDate
                        ElseIf llDate < llTempSDate Then
                            llTempSDate = llDate
                        End If
                    End If
                End If
            End If
            ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    'End If
        ilRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    If llTempSDate > 0 Then
        mObtainGameClosingDate = llTempSDate
    Else
        mObtainGameClosingDate = -1
    End If
    Exit Function
mObtainGameClosingDateErr: 'VBC NR
    On Error GoTo 0
    mObtainGameClosingDate = -1
    Exit Function
End Function




Private Sub mPrintLog(ilVpfIndex As Integer, tlSel As LOGSEL, ilLogGen As Integer, ilGameNo, slGameDate As String)
    Dim ilPass As Integer
    Dim ilEPass As Integer
    Dim ilGen As Integer
    Dim ilSOut As Integer
    Dim ilEOut As Integer
    Dim slLCO As String
    Dim ilRnf As Integer
    Dim ilOut As Integer
    Dim slFileName As String
    Dim slAlteredFileName As String     'save to filename altered due to regional logs to prevent overwriting previous log
                                        'when multiple regions are printed
    Dim slFileIndex As String
    Dim slOutput As String
    Dim ilSZone As Integer
    Dim ilEZone As Integer
    Dim ilZone As Integer
    Dim slZone As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slLetter As String
    Dim slStr As String
    Dim slStartDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slGameNo As String
    Dim slRegionName As String      '5-16-08,passed to rptsellg for Log header
    Dim slRegionCode As String      'passed to rptsellg for ODF selection
    Dim llGenTime As Long
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim ilLoopOnRAF
    Dim ilRet As Integer
    Dim llRegion As Long
    Dim ilStartRegion As Integer
    Dim ilFoundSpot As Integer
    Dim slSeqNo As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim llAirTime As Long
    Dim ilOdfPri As Integer
    Dim ilOdfSec As Integer
    ReDim tlSplitOdf(0 To 0) As SPLITODFREC
    Dim ilFill As Integer
    Dim ilFillLen As Integer
    Dim ilSvFillLen As Integer
    Dim ilLen As Integer
    Dim ilGenFillLog As Integer
    Dim ilChk As Integer
    ReDim llFillRegion(0 To 0) As Long
    ReDim ilChkFillLen(0 To 0) As Integer
    Dim tlFillOdf As ODF

    igRptCallType = LOGSJOB
    If ilGameNo > 0 Then
        slStartDate = slGameDate
        slGameNo = Trim$(str$(ilGameNo))
    Else
        slStartDate = Format$(tlSel.lStartDate, "m/d/yy")
        slGameNo = ""
    End If
    If ckcOutput(2).Value = vbChecked Then
        ilEPass = 5
    Else
        ilEPass = 2
    End If
    slStartTime = Trim$(edcTime(0).Text)
    slEndTime = Trim$(edcTime(1).Text)

    'determine if there are any regions defined
    ReDim tmRegionInfo(0 To 1) As REGIONINFO        'create a full network log entry, other entries represent a split network log
    tmRegionInfo(0).lRafCode = 0
    tmRegionInfo(0).sName = ""
    ilUpper = 1
    ilStartRegion = 1                    '0 represents a full network (vs split network log)
    If UCase$(Trim$(lbcLog.List(tlSel.iLog))) = "L78" Then
        hmRaf = CBtrvTable(TEMPHANDLE)        'Create RAF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRaf)
            btrDestroy hmRaf
            Exit Sub
        End If

        hmOdf = CBtrvTable(TEMPHANDLE)        'Create ODF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmOdf)
            ilRet = btrClose(hmRaf)
            btrDestroy hmOdf
            btrDestroy hmRaf
            Exit Sub
        End If
        imOdfRecLen = Len(tmOdf)
        tmOdfSrchKey2.iGenDate(0) = igGenDate(0)
        tmOdfSrchKey2.iGenDate(1) = igGenDate(1)
        gUnpackTimeLong igGenTime(0), igGenTime(1), False, tmOdfSrchKey2.lGenTime
        llGenTime = tmOdfSrchKey2.lGenTime
        ilRet = btrGetGreaterOrEqual(hmOdf, tmOdf, Len(tmOdf), tmOdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
        ilFoundSpot = False
        Do While (ilRet = BTRV_ERR_NONE And tmOdfSrchKey2.iGenDate(0) = igGenDate(0) And tmOdfSrchKey2.iGenDate(1) And tmOdfSrchKey2.lGenTime = llGenTime)
            If tmOdf.iType = 4 Then             'spot type only
                ilFoundSpot = True              '6-16-08 need to know at least 1 spot found so that a full network log will begenerated
                                                'even if no spots exist
                If tmOdf.lRafCode > 0 Then      'region exists
                    ilFound = False
                    For ilLoopOnRAF = 1 To ilUpper      'index one is always reserved for full network
                        If tmOdf.lRafCode = tmRegionInfo(ilLoopOnRAF).lRafCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoopOnRAF
                    If Not ilFound Then         'another different region defined
                        'access the region for name
                        tmRafSrchKey0.lCode = tmOdf.lRafCode
                        ilRet = btrGetEqual(hmRaf, tmRaf, Len(tmRaf), tmRafSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)       'Get last current record to obtain date
                        If ilRet <> BTRV_ERR_NONE Then
                            tmRegionInfo(ilUpper).sName = "Missing Region Name"
                            tmRegionInfo(ilUpper).lRafCode = tmOdf.lRafCode
                        Else
                            tmRegionInfo(ilUpper).lRafCode = tmOdf.lRafCode
                            tmRegionInfo(ilUpper).sName = tmRaf.sName
                        End If
                        ilUpper = ilUpper + 1
                        ReDim Preserve tmRegionInfo(0 To ilUpper) As REGIONINFO
                    End If
                    gUnpackDateForSort tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAirDate
                    gUnpackTimeLong tmOdf.iAirTime(0), tmOdf.iAirTime(1), False, llAirTime
                    slAirTime = Trim$(str$(llAirTime))
                    Do While Len(slAirTime) < 6
                        slAirTime = "0" & slAirTime
                    Loop
                    slSeqNo = Trim$(str$(tmOdf.iSeqNo))
                    Do While Len(slSeqNo) < 4
                        slSeqNo = "0" & slSeqNo
                    Loop
                    tlSplitOdf(UBound(tlSplitOdf)).sKey = slAirDate & slAirTime & slSeqNo
                    tlSplitOdf(UBound(tlSplitOdf)).tOdf = tmOdf
                    ReDim Preserve tlSplitOdf(0 To UBound(tlSplitOdf) + 1) As SPLITODFREC
                Else                        'full network spot
                    ilStartRegion = 0               'at least 1 full net spots, print a full net log
                End If
            End If
            ilRet = btrGetNext(hmOdf, tmOdf, Len(tmOdf), BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If Not ilFoundSpot Then         'was at least one spot found?
            ilStartRegion = 0
        Else
            If UBound(tlSplitOdf) > 0 Then
                'Region spot found, create fills.  Only generate if using L78
                'L78 test moved to the top of this routine
                'If UCase$(Trim$(lbcLog.List(tlSel.iLog))) = "L78" Then
                    If UBound(tlSplitOdf) > 1 Then
                        ArraySortTyp fnAV(tlSplitOdf(), 0), UBound(tlSplitOdf), 0, LenB(tlSplitOdf(0)), 0, LenB(tlSplitOdf(0).sKey), 0
                    End If
                    ilGenFillLog = False
                    'Create Fill to be used when region is not shown
                    'Design:
                    '    All Stations |---------------------------------------|
                    '    First Break  |--R1:30s--|---R2:15s---|-R3:30s-|
                    '
                    '    Second Break |--R4:30s--|---R2:15s---|
                    '
                    '    First Break create 30 sec fill for R4 and 15 sec fill for R2 (make up short time)
                    '    Second Break create 30 sec fill for R1 and R3 and 15 sec fill for R2 (make up short time)
                    '
                    For ilOdfPri = 0 To UBound(tlSplitOdf) - 1 Step 1
                        'Find Primary, than its secondary
                        If tlSplitOdf(ilOdfPri).tOdf.sSplitNetwork = "P" Then
                            gUnpackLength tlSplitOdf(ilOdfPri).tOdf.iLen(0), tlSplitOdf(ilOdfPri).tOdf.iLen(1), "1", True, slStr
                            ilFillLen = Val(slStr)
                            ilChkFillLen(0) = ilOdfPri
                            ReDim Preserve ilChkFillLen(0 To 1) As Integer
                            ReDim llFillRegion(0 To UBound(tmRegionInfo) - 1) As Long
                            For ilFill = 1 To UBound(tmRegionInfo) - 1 Step 1
                                llFillRegion(ilFill - 1) = tmRegionInfo(ilFill).lRafCode
                            Next ilFill
                            For ilFill = 0 To UBound(llFillRegion) - 1 Step 1
                                If tlSplitOdf(ilOdfPri).tOdf.lRafCode = llFillRegion(ilFill) Then
                                    llFillRegion(ilFill) = -1
                                    Exit For
                                End If
                            Next ilFill
                            'Find Secondary and remove its stations
                            For ilOdfSec = ilOdfPri + 1 To UBound(tlSplitOdf) - 1 Step 1
                                If tlSplitOdf(ilOdfSec).tOdf.sSplitNetwork = "S" Then
                                    If (tlSplitOdf(ilOdfPri).tOdf.iVefCode = tlSplitOdf(ilOdfSec).tOdf.iVefCode) And (tlSplitOdf(ilOdfPri).tOdf.iAirDate(0) = tlSplitOdf(ilOdfSec).tOdf.iAirDate(0)) And (tlSplitOdf(ilOdfPri).tOdf.iAirDate(1) = tlSplitOdf(ilOdfSec).tOdf.iAirDate(1)) Then
                                        If (tlSplitOdf(ilOdfPri).tOdf.iLocalTime(0) = tlSplitOdf(ilOdfSec).tOdf.iLocalTime(0)) And (tlSplitOdf(ilOdfPri).tOdf.iLocalTime(1) = tlSplitOdf(ilOdfSec).tOdf.iLocalTime(1)) And (tlSplitOdf(ilOdfPri).tOdf.sZone = tlSplitOdf(ilOdfSec).tOdf.sZone) Then
                                            For ilFill = 0 To UBound(llFillRegion) - 1 Step 1
                                                If tlSplitOdf(ilOdfSec).tOdf.lRafCode = llFillRegion(ilFill) Then
                                                    gUnpackLength tlSplitOdf(ilOdfSec).tOdf.iLen(0), tlSplitOdf(ilOdfSec).tOdf.iLen(1), "1", True, slStr
                                                    If ilFillLen < Val(slStr) Then
                                                        ilFillLen = Val(slStr)
                                                    End If
                                                    ilChkFillLen(UBound(ilChkFillLen)) = ilOdfSec
                                                    ReDim Preserve ilChkFillLen(0 To UBound(ilChkFillLen) + 1) As Integer
                                                    llFillRegion(ilFill) = -1
                                                    Exit For
                                                End If
                                            Next ilFill
                                        End If
                                    End If
                                Else
                                    Exit For
                                End If
                            Next ilOdfSec
                            'Create Fill for all remaining regions
                            ilSvFillLen = ilFillLen
                            ilLen = ilFillLen
                            Do
                                ilRet = mCreateSplitFill(ilLen, tlSplitOdf(ilOdfPri).tOdf, tlFillOdf)
                                If ilRet Then
                                    ilFillLen = ilFillLen - ilLen
                                    ilLen = ilFillLen
                                    ilStartRegion = 0               'at least 1 fill spot created
                                    ilGenFillLog = True
                                    For ilFill = 0 To UBound(llFillRegion) - 1 Step 1
                                        If llFillRegion(ilFill) <> -1 Then
                                            tlFillOdf.lRafCode = llFillRegion(ilFill)
                                            tlFillOdf.lCode = 0
                                            ilRet = btrInsert(hmOdf, tlFillOdf, imOdfRecLen, INDEXKEY3)
                                        End If
                                    Next ilFill
                                    tlFillOdf.lRafCode = -1
                                    tlFillOdf.sSplitNetwork = ""
                                    tlFillOdf.lCode = 0
                                    ilRet = btrInsert(hmOdf, tlFillOdf, imOdfRecLen, INDEXKEY3)
                                Else
                                    If ilLen > 60 Then
                                        ilLen = 60
                                    ElseIf ilLen > 30 Then
                                        ilLen = 30
                                    Else
                                        ilLen = ilLen - 5
                                        If ilLen <= 0 Then
                                            Exit Do
                                        End If
                                    End If
                                End If
                            Loop While ilFillLen > 0
                            'Create Fills for remaining amount of time
                            For ilChk = 0 To UBound(ilChkFillLen) - 1 Step 1
                                ilFillLen = ilSvFillLen
                                gUnpackLength tlSplitOdf(ilChkFillLen(ilChk)).tOdf.iLen(0), tlSplitOdf(ilChkFillLen(ilChk)).tOdf.iLen(1), "1", True, slStr
                                If ilFillLen > Val(slStr) Then
                                    ilFillLen = ilFillLen - Val(slStr)
                                    ilLen = ilFillLen
                                    Do
                                        ilRet = mCreateSplitFill(ilLen, tlSplitOdf(ilChkFillLen(ilChk)).tOdf, tlFillOdf)
                                        If ilRet Then
                                            ilFillLen = ilFillLen - ilLen
                                            ilLen = ilFillLen
                                            tlFillOdf.lRafCode = tlSplitOdf(ilChk).tOdf.lRafCode
                                            tlFillOdf.lCode = 0
                                            ilRet = btrInsert(hmOdf, tlFillOdf, imOdfRecLen, INDEXKEY3)
                                        Else
                                            If ilLen > 60 Then
                                                ilLen = 60
                                            ElseIf ilLen > 30 Then
                                                ilLen = 30
                                            Else
                                                ilLen = ilLen - 5
                                                If ilLen <= 0 Then
                                                    Exit Do
                                                End If
                                            End If
                                        End If
                                    Loop While ilFillLen > 0
                                End If
                            Next ilChk
                        End If
                    Next ilOdfPri
                    If ilGenFillLog Then
                        'tmRegionInfo(ilUpper).sName = "Remainder Log"
                        'tmRegionInfo(ilUpper).lRafCode = -1
                        'ilUpper = ilUpper + 1
                        'ReDim Preserve tmRegionInfo(0 To ilUpper) As REGIONINFO
                        tmRegionInfo(0).sName = "Remainder Log"
                        tmRegionInfo(0).lRafCode = -1
                    End If
                'End If
            End If
        End If
        ilRet = btrClose(hmOdf)
        ilRet = btrClose(hmRaf)
        btrDestroy hmOdf
        btrDestroy hmRaf
    Else
        ilStartRegion = 0                    '0 represents a full network (vs split network log)
    End If
    For ilPass = 0 To ilEPass Step 1
        ilGen = False
        ilSOut = 0
        ilEOut = 1
        If ilPass = 0 Then
            slLCO = "L"
            If tlSel.iLog > 0 Then
                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                    If lbcLog.List(tlSel.iLog) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                        igRptType = tmRnfList(ilRnf).tRnf.iCode
                        ilGen = True
                        Exit For
                    End If
                Next ilRnf
            End If
        ElseIf ilPass = 1 Then
            slLCO = "C"
            If tlSel.iCP > 0 Then
                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                    If lbcCP.List(tlSel.iCP) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                        igRptType = tmRnfList(ilRnf).tRnf.iCode
                        ilGen = True
                        If "C17" = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            ilGen = False
                        End If
                        Exit For
                    End If
                Next ilRnf
            End If
        ElseIf ilPass = 2 Then
            slLCO = "O"
            If tlSel.iOther > 0 Then
                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                    If lbcOther.List(tlSel.iOther) = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                        igRptType = tmRnfList(ilRnf).tRnf.iCode
                        ilGen = True
                        Exit For
                    End If
                Next ilRnf
            End If
        ElseIf ilPass = 3 Then
            slLCO = "L"
            igRptType = tgVpf(ilVpfIndex).iRnfSvLogCode
            If igRptType > 0 Then
                ilGen = True
            Else
                ilGen = False
            End If
            ilSOut = 2
            ilEOut = 2
        ElseIf ilPass = 4 Then
            slLCO = "C"
            igRptType = tgVpf(ilVpfIndex).iRnfSvCertCode
            If igRptType > 0 Then
                ilGen = True
                For ilRnf = 0 To UBound(tmRnfList) - 1 Step 1
                    If igRptType = tmRnfList(ilRnf).tRnf.iCode Then
                        If "C17" = UCase$(Trim$(tmRnfList(ilRnf).tRnf.sName)) Then
                            ilGen = False
                        End If
                        Exit For
                    End If
                Next ilRnf
            Else
                ilGen = False
            End If
            ilSOut = 2
            ilEOut = 2
        ElseIf ilPass = 5 Then
            slLCO = "O"
            igRptType = tgVpf(ilVpfIndex).iRnfSvPlayCode
            If igRptType > 0 Then
                ilGen = True
            Else
                ilGen = False
            End If
            ilSOut = 2
            ilEOut = 2
        Else
            ilGen = False
        End If
        If ilGen Then
            If tgSpf.sGUseAffSys = "Y" Then
                'If (slOutput = "0") And (rbcOutput(0).Caption = "None") Then
                If (ckcOutput(0).Value = vbUnchecked) And (ckcOutput(1).Value = vbUnchecked) And (ckcOutput(2).Value = vbUnchecked) Then
                    ilGen = False
                End If
            End If
        End If
        If ilGen Then
            For ilOut = ilSOut To ilEOut Step 1
                If ckcOutput(ilOut).Value = vbChecked Then
                    slFileName = ""
                    slFileIndex = ""
                    Select Case ilOut
                        Case 0  'Display
                            slOutput = "0"
                            'ilSZone = 0
                            'ilEZone = 0
                            If tlSel.iZone = 0 Then
                                ilSZone = 0
                                ilEZone = 0
                            Else
                                ilSZone = tlSel.iZone
                                ilEZone = tlSel.iZone
                            End If
                        Case 1  'Print
                            slOutput = "1"
                            'ilSZone = 0
                            'ilEZone = 0
                            If tlSel.iZone = 0 Then
                                ilSZone = 0
                                ilEZone = 0
                            Else
                                ilSZone = tlSel.iZone
                                ilEZone = tlSel.iZone
                            End If
                        Case 2  'File
                            slOutput = "2"
                            If tlSel.iZone = 0 Then
                                'ilSZone = 1
                                'ilEZone = 4
                                ilSZone = 0
                                ilEZone = 0
                                For ilZone = LBound(tgVpf(tlSel.iVpfIndex).sGZone) To UBound(tgVpf(tlSel.iVpfIndex).sGZone) Step 1
                                   If Trim$(tgVpf(tlSel.iVpfIndex).sGZone(ilZone)) <> "" Then
                                        ilSZone = 1
                                        ilEZone = 4
                                        Exit For
                                   End If
                                Next ilZone
                            Else
                                ilSZone = tlSel.iZone
                                ilEZone = tlSel.iZone
                            End If
                            gObtainYearMonthDayStr slStartDate, True, slYear, slMonth, slDay
                            If Val(slMonth) <= 9 Then
                                slMonth = right$(slMonth, 1)
                            ElseIf Val(slMonth) = 10 Then
                                slMonth = "A"
                            ElseIf Val(slMonth) = 11 Then
                                slMonth = "B"
                            ElseIf Val(slMonth) = 12 Then
                                slMonth = "C"
                            End If
                            'dan 6-12-09 took out hard coding of name
                            If cbcFile.ListIndex = 0 Then
                           ' If cbcFile.List(cbcFile.ListIndex) = "Acrobat PDF" Then
                                If (ilSZone = 0) And (ilEZone = 0) Then
                                    slLetter = Trim$(Left$(tmVef.sCodeStn, 5))  '1-28-05 chg from 4 to 5 char
                                Else
                                    slLetter = Trim$(Left$(tmVef.sCodeStn, 5))  '1-28-05 chg from 3 to 5 char
                                End If
                            Else
                                slLetter = Trim$(Left$(tmVef.sCodeStn, 5))      '1-28-05 chg from 3 to 5 char
                            End If
                            If ilPass = 3 Then
                                'slFileName = slMonth & slDay & Right$(slYear, 2) & "L" & slLetter
                                slFileName = slMonth & slDay & "L" & slGameNo & slLetter
                            ElseIf ilPass = 4 Then
                                'slFileName = slMonth & slDay & Right$(slYear, 2) & "C" & slLetter
                                slFileName = slMonth & slDay & "C" & slGameNo & slLetter
                            Else
                                'slFileName = slMonth & slDay & Right$(slYear, 2) & "O" & slLetter
                                slFileName = slMonth & slDay & "O" & slGameNo & slLetter
                            End If
                            
                            'TTP 10419 - Log generation: when saving directly to file, an illegal character in the station code will cause "failed to export report, invalid directory" error message
                            slFileName = gFileNameFilter(Trim$(slFileName))
                            
                            'Dan 6-12-09 updated listindex to match new values
                            Select Case cbcFile.ListIndex
                                Case 0
                                    slFileName = slFileName & ".Pdf"
                                Case 1
                                    slFileName = slFileName & ".Xls"
                                Case 2
                                    slFileName = slFileName & ".Doc"
                                Case 4
                                    slFileName = slFileName & ".Csv"
                                Case 5
                                    slFileName = slFileName & ".Rtf"
                                Case Else
                                    slFileName = slFileName & ".Txt"
                            End Select
                            'If cbcFile.ListIndex = 6 Then
'                            If cbcFile.ListIndex = 6 Then       '10-24-06
'                                slFileName = slFileName & ".Rtf"
'                            ElseIf cbcFile.List(cbcFile.ListIndex) = "Acrobat PDF" Then
'                                slFileName = slFileName & ".Pdf"
'                            Else
'                                slFileName = slFileName & ".Txt"
'                            End If
                            If cbcFile.ListIndex >= 0 Then
                                slFileIndex = Trim$(str$(cbcFile.ListIndex))
                            Else
                                slFileIndex = "6"       'default to rtf 10-24-06
                                'slFileIndex = 5     '10-19-01
                            End If
                    End Select
                    'Dan 6-12-09 removed reference to acrobat
                    If (ilOut = 2) And (cbcFile.ListIndex = 0) Then
                    'If (ilOut = 2) And (cbcFile.List(cbcFile.ListIndex) = "Acrobat PDF") Then
                        slOutput = "2"
                        slFileIndex = "0"
                        'DoEvents
                        'gSwitchToPDF cdcSetup, 0
                        'DoEvents
                    End If
                    For ilZone = ilSZone To ilEZone Step 1
                        Select Case ilZone
                            Case 0
                                slZone = ""
                            Case 1
                                slZone = "E"
                            Case 2
                                slZone = "C"
                            Case 3
                                slZone = "M"
                            Case 4
                                slZone = "P"
                        End Select

                        For ilLoopOnRAF = ilStartRegion To UBound(tmRegionInfo) - 1      'index 0 always represents full network (vs regions)
                            slRegionName = tmRegionInfo(ilLoopOnRAF).sName                      '5-16-08 required to determine the split network region (currently in L78 only) passed in command string
                            llRegion = tmRegionInfo(ilLoopOnRAF).lRafCode                       'required to determine the split network region (currently in L78 only)
                            slRegionCode = Trim$(str(llRegion))
                            slAlteredFileName = slFileName
                            If ilOut = 2 And Trim$(slRegionName) <> "" Then           'if save to file and saving region log (L78), need to adjust the filename
                                                        'so that each region doesnt overwrite the previous one.  Use the loop counter to keep it unique.
                                'find the extension and insert the region name
                                ilRet = InStr(slFileName, ".")
                                If ilRet > 0 Then
                                    slStr = Mid(slAlteredFileName, 1, ilRet - 1)       'extract up to period for extension name
                                    slStr = slStr & "_" & Trim$(str(ilLoopOnRAF))
                                    slStr = slStr & "." & Mid(slAlteredFileName, ilRet + 1)
                                    slAlteredFileName = Trim$(slStr)
                                End If
                            End If
                            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
                            edcLinkSrceDoneMsg.Text = ""
                            If ilLogGen <= LBound(tmLogGen) Then
                                If (Not igStdAloneMode) And (imShowHelpMsg) Then
                                    If igTestSystem Then
                                        slStr = "Logs^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & Trim$(str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(str$(tlSel.lEndDate - tlSel.lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(str$(tmVef.iCode)) & "\" & Trim$(str$(ilZone)) & "\" & slOutput & "\" & slFileIndex & "\" & slZone & slAlteredFileName & "\" & sgGenDate & "\" & sgGenTime & "\" & slLCO & "\" & slRegionCode & "\" & slRegionName
                                    Else
                                        slStr = "Logs^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & Trim$(str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(str$(tlSel.lEndDate - tlSel.lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(str$(tmVef.iCode)) & "\" & Trim$(str$(ilZone)) & "\" & slOutput & "\" & slFileIndex & "\" & slZone & slAlteredFileName & "\" & sgGenDate & "\" & sgGenTime & "\" & slLCO & "\" & slRegionCode & "\" & slRegionName
                                    End If
                                Else
                                    If igTestSystem Then
                                        slStr = "Logs^Test^NOHELP\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & Trim$(str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(str$(tlSel.lEndDate - tlSel.lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(str$(tmVef.iCode)) & "\" & Trim$(str$(ilZone)) & "\" & slOutput & "\" & slFileIndex & "\" & slZone & slAlteredFileName & "\" & sgGenDate & "\" & sgGenTime & "\" & slLCO & "\" & slRegionCode & "\" & slRegionName
                                    Else
                                        slStr = "Logs^Prod^NOHELP\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & Trim$(str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(str$(tlSel.lEndDate - tlSel.lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(str$(tmVef.iCode)) & "\" & Trim$(str$(ilZone)) & "\" & slOutput & "\" & slFileIndex & "\" & slZone & slAlteredFileName & "\" & sgGenDate & "\" & sgGenTime & "\" & slLCO & "\" & slRegionCode & "\" & slRegionName
                                    End If
                                End If
                            Else
                                If (Not igStdAloneMode) And (imShowHelpMsg) Then
                                    If igTestSystem Then
                                        slStr = "Logs^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & Trim$(str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(str$(tlSel.lEndDate - tlSel.lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(str$(tmLogGen(ilLogGen).iSimVefCode)) & "\" & Trim$(str$(ilZone)) & "\" & slOutput & "\" & slFileIndex & "\" & slZone & slAlteredFileName & "\" & sgGenDate & "\" & sgGenTime & "\" & slLCO & "\" & slRegionCode & "\" & slRegionName
                                    Else
                                        slStr = "Logs^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & Trim$(str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(str$(tlSel.lEndDate - tlSel.lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(str$(tmLogGen(ilLogGen).iSimVefCode)) & "\" & Trim$(str$(ilZone)) & "\" & slOutput & "\" & slFileIndex & "\" & slZone & slAlteredFileName & "\" & sgGenDate & "\" & sgGenTime & "\" & slLCO & "\" & slRegionCode & "\" & slRegionName
                                    End If
                                Else
                                    If igTestSystem Then
                                        slStr = "Logs^Test^NOHELP\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & Trim$(str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(str$(tlSel.lEndDate - tlSel.lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(str$(tmLogGen(ilLogGen).iSimVefCode)) & "\" & Trim$(str$(ilZone)) & "\" & slOutput & "\" & slFileIndex & "\" & slZone & slAlteredFileName & "\" & sgGenDate & "\" & sgGenTime & "\" & slLCO & "\" & slRegionCode & "\" & slRegionName
                                    Else
                                        slStr = "Logs^Prod^NOHELP\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & Trim$(str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(str$(tlSel.lEndDate - tlSel.lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(str$(tmLogGen(ilLogGen).iSimVefCode)) & "\" & Trim$(str$(ilZone)) & "\" & slOutput & "\" & slFileIndex & "\" & slZone & slAlteredFileName & "\" & sgGenDate & "\" & sgGenTime & "\" & slLCO & "\" & slRegionCode & "\" & slRegionName
                                    End If
                                End If
                            End If
                            '10-10-01
                            'lgShellRet = Shell(sgExePath & "RptSelLg.Exe " & slStr, 1)
                            'While GetModuleUsage(ilShell) > 0
                            '    ilRet = DoEvents()
                            'Wend

                            'gShellAndWait Logs, sgExePath & "RptSelLg.Exe " & slStr, vbNormalFocus
                            On Error Resume Next
                            Kill Trim$(sgRptPath) & "savelogo.bmp"
                            On Error GoTo 0

                            sgCommandStr = slStr
                            RptSelLg.Show vbModal
                            'If (ilOut = 2) And (cbcFile.List(cbcFile.ListIndex) = "Acrobat PDF") Then
                            '    FileCopy "c:\csi\csirpt.pdf", sgExportPath & sLZone & slFileName
                            'End If
                        Next ilLoopOnRAF
                    Next ilZone
                    'If (ilOut = 2) And (cbcFile.List(cbcFile.ListIndex) = "Acrobat PDF") Then
                    '    slOutput = "2"
                    '    gSwitchToPDF cdcSetup, 1
                    'End If
                End If
            Next ilOut
        End If
    Next ilPass
End Sub

Private Sub mExportLog(slInStartDate As String, slInEndDate As String, ilGameNo As Integer, llGsfCode As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slFileName                    slLetter                      slFMonth                  *
'*  slFDay                        slFYear                       slToFile                  *
'*  slExt                         slStartDateTime               llEndDate                 *
'*                                                                                        *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  cmcExportErr                                                                          *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim llRecPos As Long        'Record location
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE          '6-12-02 long
    Dim slRecord As String
    Dim slVisitingTeam As String
    Dim slHomeTeam As String
    Dim slLanguage As String
    Dim slTime As String
    Dim slLength As String
    Dim ilLoop As Integer
    Dim slAdvtName  As String
    Dim slCart As String
    Dim slISCI As String
    Dim slCreativeTitle As String
    Dim slAvailName As String
    Dim ilVpfIndex As Integer
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim ilExportSpot As Integer
    '11/6/14: Add season to event
    Dim slSeason As String
    
    If bgLogFirstCallToVpfFind Then
        ilVpfIndex = gVpfFind(Logs, tmVef.iCode)
        bgLogFirstCallToVpfFind = False
    Else
        ilVpfIndex = gVpfFindIndex(tmVef.iCode)
    End If
    slStartDate = slInStartDate
    slEndDate = slInEndDate

'Moved to mOpenExportFile to handle multi-games on same date
'    gObtainYearMonthDayStr slStartDate, True, slFYear, slFMonth, slFDay
'
'    slFileName = slFYear & slFMonth & slFDay
'    slExt = ".csv"
'    slLetter = Trim$(tmVef.sCodeStn)
'    ilRet = 0
'    On Error GoTo cmcExportErr:
'
'    slToFile = sgExportPath & Trim$(slLetter) & Trim$(slFileName) & slExt   'ssssZmmdd.ext
'
'    slStartDateTime = FileDateTime(slToFile)     '1-6-05 chged from sgExportPath to new sgProphetExportPath
'
'    If ilRet = 0 Then
'        Kill slToFile   '1-6-05 chg from sgExportpath to new sgProphetExportPath
'    End If
'    On Error GoTo 0
'    ilRet = 0
'    On Error GoTo cmcExportErr:
'    hmTo = FreeFile
'    Open slToFile For Output As hmTo
'    If ilRet <> 0 Then
'        'gClearODF                   'remove all the ODFs for the logs just created
'        'Print #hmMsg, "** Terminated **"
'        'Print #hmMsg, "Open error " & slToFile & " Error #" & Str$(ilRet)
'        'Close #hmMsg
'        'MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'        'Exit Sub
'    End If

'    'Generate header information
'    slRecord = "Record Layout: Start"
'    Print #hmTo, slRecord
'    slRecord = "Date:, Generation Date, Generation Time"
'    Print #hmTo, slRecord
'    slRecord = "Vehicle:, Vehicle Name, Vehicle Station Code"
'    Print #hmTo, slRecord
'    slRecord = "Game:, GameNo, Date, Visiting Team Name, Visiting Team Abbreviation, Home Team Name, Home Team Abbreviation, Start Time, Status"
'    Print #hmTo, slRecord
'    slRecord = "Spot:, Time, Position #, Break #, Advertiser Name, Advertiser ID, Product, Spot Length, Cart ISCI, Creative Title, Avail Name, Zone"
'    Print #hmTo, slRecord
'    slRecord = "Record Layout: End"
'    Print #hmTo, slRecord
'
'
'    slRecord = "Date:, " & slStartDate & ", " & gFormatTimeLong(lgGenTime, "A", "1")
'    Print #hmTo, slRecord
'    slRecord = "Vehicle:, " & Trim$(tmVef.sName) & ", " & Trim$(tmVef.sCodeStn)
'    Print #hmTo, slRecord

    If llGsfCode > 0 Then
        'tmGsfSrchKey0.lCode = llGsfCode
        'ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        'If ilRet <> BTRV_ERR_NONE Then
        '    'Close #hmTo
        '    ilRet = btrClose(hmOdf)
        'End If
        slSeason = Trim$(tmGhf.sSeasonName)
        slVisitingTeam = ","
        For ilLoop = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1
            If tmTeam(ilLoop).iCode = tmGsf.iVisitMnfCode Then
                slVisitingTeam = Trim$(tmTeam(ilLoop).sName) & "," & Trim$(tmTeam(ilLoop).sUnitType)
                Exit For
            End If
        Next ilLoop

        slHomeTeam = ","
        For ilLoop = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1
            If tmTeam(ilLoop).iCode = tmGsf.iHomeMnfCode Then
                slHomeTeam = Trim$(tmTeam(ilLoop).sName) & "," & Trim$(tmTeam(ilLoop).sUnitType)
                Exit For
            End If
        Next ilLoop
        slLanguage = ","
        If tmGsf.iLangMnfCode > 0 Then
            For ilLoop = LBound(tmLang) To UBound(tmLang) - 1 Step 1
                If tmLang(ilLoop).iCode = tmGsf.iLangMnfCode Then
                    slLanguage = Trim$(tmLang(ilLoop).sName) & "," & Trim$(tmLang(ilLoop).sUnitType)
                    Exit For
                End If
            Next ilLoop
        End If
        gUnpackTime tmGsf.iAirTime(0), tmGsf.iAirTime(1), "A", "1", slTime
        slRecord = "Event:," & slSeason & "," & ilGameNo & "," & slStartDate & "," & slVisitingTeam & "," & slHomeTeam & "," & slTime & "," & tmGsf.sGameStatus & "," & slLanguage & "," & Trim$(tmGsf.sFeedSource)
        Print #hmTo, slRecord
    Else
        slRecord = "Event:," & "" & "," & 0 & "," & slStartDate & "," & "" & "," & "" & "," & "" & "," & "" & "," & "" & "," & ""
        Print #hmTo, slRecord
    End If
    ilRet = 0
    imOdfRecLen = Len(tmOdf)  'Get and save ADF record length
    hmOdf = CBtrvTable(TEMPHANDLE)        'Create ADF object handle
    ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)

    ilRecLen = Len(tmOdf)
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hmOdf   'Clear any previous extend operation

    tmOdfSrchKey2.iGenDate(0) = igGenDate(0)   'ilLogDate0
    tmOdfSrchKey2.iGenDate(1) = igGenDate(1)
    '10-9-01
    tmOdfSrchKey2.lGenTime = lgGenTime
    ilRet = btrGetEqual(hmOdf, tmOdf, imOdfRecLen, tmOdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date

    Call btrExtSetBounds(hmOdf, llNoRec, -1, "UC", "ODF", "") 'Set extract limits (all records)

    tlDateTypeBuff.iDate0 = igGenDate(0)
    tlDateTypeBuff.iDate1 = igGenDate(1)
    ilOffSet = gFieldOffset("ODF", "ODFGenDate")
    ilRet = btrExtAddLogicConst(hmOdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlDateTypeBuff, 4)
    tlLongTypeBuff.lCode = lgGenTime            '6-12-02
    ilOffSet = gFieldOffset("ODF", "ODFGenTime")
    ilRet = btrExtAddLogicConst(hmOdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlLongTypeBuff, 4)

    ilRet = btrExtAddField(hmOdf, 0, ilRecLen)  'Extract the whole record

    ilRet = btrExtGetNext(hmOdf, tmOdf, ilRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            Exit Sub
        End If
    End If
    Do While ilRet = BTRV_ERR_REJECT_COUNT
        ilRet = btrExtGetNext(hmOdf, tmOdf, ilRecLen, llRecPos)
    Loop
    Do While (ilRet = BTRV_ERR_NONE) And (tmOdf.iGameNo = ilGameNo)
        If tmOdf.iType = 4 Then
            ilExportSpot = False
            If Trim$(tmOdf.sZone) = "" Then
                ilExportSpot = True
            Else
                For ilLoop = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
                    If (Trim$(tgVpf(ilVpfIndex).sGFed(ilLoop)) = "*") And (Left$(tgVpf(ilVpfIndex).sGZone(ilLoop), 1) = Left$(tmOdf.sZone, 1)) Then
                        ilExportSpot = True
                    End If
                Next ilLoop
            End If
            If ilExportSpot Then
                If llGsfCode <= 0 Then
                    gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), llStartDate
                    If (llStartDate <> gDateValue(slStartDate)) And (llStartDate <= gDateValue(slEndDate)) Then
                        gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slStartDate
                        slRecord = "Event:," & "" & "," & 0 & "," & slStartDate & "," & "" & "," & "" & "," & "" & "," & ""
                        Print #hmTo, slRecord
                    End If
                End If
                gUnpackTime tmOdf.iAirTime(0), tmOdf.iAirTime(1), "A", "1", slTime
                slAdvtName = "AdvertiserName, "
                ilLoop = gBinarySearchAdf(tmOdf.iAdfCode)
                If ilLoop <> -1 Then
                    'If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
                    '    slAdvtName = Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
                    'Else
                    '    slAdvtName = Trim$(tgCommAdf(ilLoop).sName) & ", "
                    'End If
                    slAdvtName = """" & Trim$(tgCommAdf(ilLoop).sName) & """" & "," & """" & Trim$(tgCommAdf(ilLoop).sAbbr) & """"
                End If
                slCart = ""
                slCreativeTitle = ""
                slISCI = ""
                tmCifSrchKey0.lCode = tmOdf.lCifCode     'copy code
                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    tmMcfSrchKey0.iCode = tmCif.iMcfCode
                    ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmMcf.sName = "C"
                        tmMcf.sPrefix = "C"
                    End If
                    If Trim$(tmCif.sCut) = "" Then
                        slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & " "
                    Else
                        slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut) & " "
                    End If
                    tmCpfSrchKey0.lCode = tmCif.lcpfCode     'product/isci/creative title
                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slCreativeTitle = """" & Trim$(tmCpf.sCreative) & """"
                        slISCI = """" & Trim$(tmCpf.sISCI) & """"
                    End If
                End If
                slAvailName = ""
                If tgVpf(ilVpfIndex).sAvailNameOnWeb = "Y" Then
                    tmAnfSrchKey0.iCode = tmOdf.ianfCode
                    ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slAvailName = """" & Trim$(tmAnf.sName) & """"
                    End If
                End If
                gUnpackLength tmOdf.iLen(0), tmOdf.iLen(1), "1", True, slLength
                slRecord = "Spot:," & slTime & "," & tmOdf.iPositionNo & "," & tmOdf.iBreakNo & ","
                slRecord = slRecord & slAdvtName & "," & Trim$(tmOdf.sProduct) & "," & slLength & ","
                slRecord = slRecord & slCart & "," & slISCI & "," & slCreativeTitle & "," & slAvailName & "," & Trim$(tmOdf.sZone)
                Print #hmTo, slRecord
                '11/6/14: Generate Region copy records
                mExportRegionCopyRecord tmOdf.lCode
            End If
        End If
        ilRecLen = Len(tmOdf)
        ilRet = btrExtGetNext(hmOdf, tmOdf, ilRecLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmOdf, tmOdf, ilRecLen, llRecPos)
        Loop
    Loop
    'slRecord = "Spot: End"
    'Print #hmTo, slRecord
    'Close #hmTo
    ilRet = btrClose(hmOdf)
    btrDestroy hmOdf
    Exit Sub

cmcExportErr: 'VBC NR
    ilRet = err.Number
    Resume Next
End Sub



Private Function mOpenExportFile(slDate As String) As Integer
    Dim ilRet As Integer
    Dim slFileName As String
    Dim slLetter As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slFYear As String
    Dim slToFile As String
    Dim slExt As String
    Dim slDateTime As String
    Dim slRecord As String

    gObtainYearMonthDayStr slDate, True, slFYear, slFMonth, slFDay

    slFileName = slFYear & slFMonth & slFDay
    slExt = ".csv"
    slLetter = Trim$(tmVef.sCodeStn)
    ilRet = 0
    'On Error GoTo mOpenExportFileErr:

    slToFile = sgExportPath & Trim$(slLetter) & Trim$(slFileName) & slExt   'ssssZmmdd.ext

    'slDateTime = FileDateTime(slToFile)     '1-6-05 chged from sgExportPath to new sgProphetExportPath
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile   '1-6-05 chg from sgExportpath to new sgProphetExportPath
    End If
    On Error GoTo 0
    ilRet = 0
    'On Error GoTo mOpenExportFileErr:
    'hmTo = FreeFile
    'Open slToFile For Output As hmTo
    ilRet = gFileOpen(slToFile, "Output", hmTo)
    If ilRet <> 0 Then
        mOpenExportFile = False
    Else
        'Generate header information
        slRecord = "Record Layout: Start"
        Print #hmTo, slRecord
        slRecord = "Date:,Generation Date,Generation Time"
        Print #hmTo, slRecord
        slRecord = "Vehicle:,Vehicle Name,Vehicle Station Code"
        Print #hmTo, slRecord
        slRecord = "Event:,Season,GameNo,Date,Visiting Team Name,Visiting Team Abbreviation,Home Team Name,Home Team Abbreviation,Start Time,Status,Language,English,Feed Source"
        Print #hmTo, slRecord
        slRecord = "Spot:,Time,Position #,Break #,Advertiser Name,Advertiser Abbreviation,Product,Spot Length,Cart,ISCI,Creative Title,Avail Name,Zone"
        Print #hmTo, slRecord
        If (tgSpf.sGUseAffSys = "Y") Then
            slRecord = "Copy:,Type,Advertiser Name,Advertiser Abbreviation,Cart,ISCI,Product,Creative Title,Call Letters,Station ID,XDS Cue"
            Print #hmTo, slRecord
        End If
        slRecord = "Record Layout: End"
        Print #hmTo, slRecord
        slRecord = "Date:," & slDate & "," & gFormatTimeLong(lgGenTime, "A", "1")
        Print #hmTo, slRecord
        slRecord = "Vehicle:," & Trim$(tmVef.sName) & "," & Trim$(tmVef.sCodeStn)
        Print #hmTo, slRecord
        mOpenExportFile = True
    End If
    Exit Function
'mOpenExportFileErr:
'    ilRet = Err.Number
'    Resume Next


End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintLnTitle                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint Header Titles            *
'*                                                     *
'*******************************************************
Private Sub mPaintLogTitle(ilIndex As Integer)
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    llColor = pbcLogs(ilIndex).ForeColor
    slFontName = pbcLogs(ilIndex).FontName
    flFontSize = pbcLogs(ilIndex).FontSize
    ilFillStyle = pbcLogs(ilIndex).FillStyle
    llFillColor = pbcLogs(ilIndex).FillColor
    pbcLogs(ilIndex).ForeColor = BLUE
    pbcLogs(ilIndex).FontBold = False
    pbcLogs(ilIndex).FontSize = 7
    pbcLogs(ilIndex).FontName = "Arial"
    pbcLogs(ilIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    pbcLogs(ilIndex).Line (tmCtrls(CHKINDEX).fBoxX - 15, 15)-Step(tmCtrls(CHKINDEX).fBoxW + 15, tmCtrls(CHKINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).CurrentX = tmCtrls(CHKINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcLogs(ilIndex).Print "Gen"
    pbcLogs(ilIndex).Line (tmCtrls(WRKDATEINDEX).fBoxX - 15, 15)-Step(tmCtrls(WRKDATEINDEX).fBoxW + 15, tmCtrls(WRKDATEINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).Line (tmCtrls(WRKDATEINDEX).fBoxX, 30)-Step(tmCtrls(WRKDATEINDEX).fBoxW - 15, tmCtrls(WRKDATEINDEX).fBoxY - 60), LIGHTBLUE, BF
    pbcLogs(ilIndex).CurrentX = tmCtrls(WRKDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcLogs(ilIndex).Print "Working"
    pbcLogs(ilIndex).CurrentX = tmCtrls(WRKDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = tmCtrls(WRKDATEINDEX).fBoxY / 2 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcLogs(ilIndex).Print "Date"
    pbcLogs(ilIndex).Line (tmCtrls(VEHINDEX).fBoxX - 15, 15)-Step(tmCtrls(VEHINDEX).fBoxW + 15, tmCtrls(VEHINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).Line (tmCtrls(VEHINDEX).fBoxX, 30)-Step(tmCtrls(VEHINDEX).fBoxW - 15, tmCtrls(VEHINDEX).fBoxY - 60), LIGHTBLUE, BF
    pbcLogs(ilIndex).CurrentX = tmCtrls(VEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcLogs(ilIndex).Print "Vehicle"
    pbcLogs(ilIndex).Line (tmCtrls(LLDINDEX).fBoxX - 15, 15)-Step(tmCtrls(LLDINDEX).fBoxW + 15, tmCtrls(LLDINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).Line (tmCtrls(LLDINDEX).fBoxX, 30)-Step(tmCtrls(LLDINDEX).fBoxW - 15, tmCtrls(LLDINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcLogs(ilIndex).CurrentX = tmCtrls(LLDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15
    pbcLogs(ilIndex).Print "Last Log"
    pbcLogs(ilIndex).CurrentX = tmCtrls(LLDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = tmCtrls(LLDINDEX).fBoxY / 2 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcLogs(ilIndex).Print "Date"
    pbcLogs(ilIndex).Line (tmCtrls(LEADTIMEINDEX).fBoxX - 15, 15)-Step(tmCtrls(LEADTIMEINDEX).fBoxW + 15, tmCtrls(LEADTIMEINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).CurrentX = tmCtrls(LEADTIMEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15
    pbcLogs(ilIndex).Print "Lead"
    pbcLogs(ilIndex).CurrentX = tmCtrls(LEADTIMEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = tmCtrls(LEADTIMEINDEX).fBoxY / 2 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcLogs(ilIndex).Print "Time"
    pbcLogs(ilIndex).Line (tmCtrls(CYCLEINDEX).fBoxX - 15, 15)-Step(tmCtrls(CYCLEINDEX).fBoxW + 15, tmCtrls(CYCLEINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).CurrentX = tmCtrls(CYCLEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcLogs(ilIndex).Print "Cycle"
    ilHalfY = tmCtrls(SDATEINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    pbcLogs(ilIndex).Line (tmCtrls(SDATEINDEX).fBoxX - 15, 15)-Step(tmCtrls(SDATEINDEX).fBoxW + tmCtrls(EDATEINDEX).fBoxW + 30, ilHalfY), BLUE, B
    pbcLogs(ilIndex).CurrentX = tmCtrls(SDATEINDEX).fBoxX + tmCtrls(SDATEINDEX).fBoxW / 2 + tmCtrls(EDATEINDEX).fBoxW / 2 - pbcLogs(ilIndex).TextWidth("Closing") / 2 + 15 'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcLogs(ilIndex).Print "Closing"
    pbcLogs(ilIndex).Line (tmCtrls(SDATEINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(SDATEINDEX).fBoxW + 15, tmCtrls(SDATEINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).Line (tmCtrls(SDATEINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(SDATEINDEX).fBoxW - 15, tmCtrls(SDATEINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcLogs(ilIndex).CurrentX = tmCtrls(SDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = ilHalfY + 30
    pbcLogs(ilIndex).Print "Start Date"
    pbcLogs(ilIndex).Line (tmCtrls(EDATEINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(EDATEINDEX).fBoxW + 15, tmCtrls(EDATEINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).Line (tmCtrls(EDATEINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(EDATEINDEX).fBoxW - 15, tmCtrls(EDATEINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcLogs(ilIndex).CurrentX = tmCtrls(EDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = ilHalfY + 30
    pbcLogs(ilIndex).Print "End Date"

    pbcLogs(ilIndex).Line (tmCtrls(LOGINDEX).fBoxX - 15, 15)-Step(tmCtrls(LOGINDEX).fBoxW + 15, tmCtrls(LOGINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).CurrentX = tmCtrls(LOGINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15
    pbcLogs(ilIndex).Print "Log"
    pbcLogs(ilIndex).Line (tmCtrls(CPINDEX).fBoxX - 15, 15)-Step(tmCtrls(CPINDEX).fBoxW + 15, tmCtrls(CPINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).CurrentX = tmCtrls(CPINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15
    pbcLogs(ilIndex).Print "C of P"
    pbcLogs(ilIndex).Line (tmCtrls(LOGOINDEX).fBoxX - 15, 15)-Step(tmCtrls(LOGOINDEX).fBoxW + 15, tmCtrls(LOGOINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).CurrentX = tmCtrls(LOGOINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15
    pbcLogs(ilIndex).Print "Logo"
    pbcLogs(ilIndex).Line (tmCtrls(OTHERINDEX).fBoxX - 15, 15)-Step(tmCtrls(OTHERINDEX).fBoxW + 15, tmCtrls(OTHERINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).CurrentX = tmCtrls(OTHERINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15
    pbcLogs(ilIndex).Print "Other"
    pbcLogs(ilIndex).Line (tmCtrls(ZONEINDEX).fBoxX - 15, 15)-Step(tmCtrls(ZONEINDEX).fBoxW + 15, tmCtrls(ZONEINDEX).fBoxY - 30), BLUE, B
    pbcLogs(ilIndex).CurrentX = tmCtrls(ZONEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcLogs(ilIndex).CurrentY = 15
    pbcLogs(ilIndex).Print "Zone"

    ilLineCount = 0
    llTop = tmCtrls(1).fBoxY
    Do
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            If (ilLoop = WRKDATEINDEX) Or (ilLoop = VEHINDEX) Or (ilLoop = LLDINDEX) Or (ilLoop = SDATEINDEX) Or (ilLoop = EDATEINDEX) Then
                pbcLogs(ilIndex).FillStyle = 0 'Solid
                pbcLogs(ilIndex).FillColor = LIGHTYELLOW
            End If
            pbcLogs(ilIndex).Line (tmCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            If (ilLoop = WRKDATEINDEX) Or (ilLoop = VEHINDEX) Or (ilLoop = LLDINDEX) Or (ilLoop = SDATEINDEX) Or (ilLoop = EDATEINDEX) Then
                pbcLogs(ilIndex).FillStyle = ilFillStyle
                pbcLogs(ilIndex).FillColor = llFillColor
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmCtrls(1).fBoxH + 15
    Loop While llTop + tmCtrls(1).fBoxH < pbcLogs(ilIndex).height
    vbcLogs.LargeChange = ilLineCount - 1
    pbcLogs(ilIndex).FontSize = flFontSize
    pbcLogs(ilIndex).FontName = slFontName
    pbcLogs(ilIndex).FontSize = flFontSize
    pbcLogs(ilIndex).ForeColor = llColor
    pbcLogs(ilIndex).FontBold = True
End Sub

Private Function mObtainSplitReplacments() As Integer
    Dim ilNum As Integer
    Dim slNum As String
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilLast As Integer
    Dim ilRet As Integer

    ReDim tmRBofRec(0 To 0) As SPLITBOFREC
    ReDim tmSplitNetLastFill(0 To 0) As SPLITNETLASTFILL
    ilUpper = LBound(tmRBofRec)
    Randomize
    ilRet = btrGetFirst(hmBof, tmBof, imBofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If tmBof.sType = "R" Then
            If tmBof.lCifCode > 0 Then
                tmCifSrchKey0.lCode = tmBof.lCifCode
                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slStr = Trim$(str$(tmCif.iLen))
                    Do While Len(slStr) < 3
                        slStr = "0" & slStr
                    Loop
                    ilNum = Int(10000 * Rnd + 1)
                    slNum = Trim$(str$(ilNum))
                    Do While Len(slNum) < 5
                        slNum = "0" & slNum
                    Loop
                    tmRBofRec(ilUpper).sKey = slStr & slNum
                    tmRBofRec(ilUpper).tBof = tmBof
                    tmRBofRec(ilUpper).iLen = tmCif.iLen
                    ilUpper = ilUpper + 1
                    ReDim Preserve tmRBofRec(0 To ilUpper) As SPLITBOFREC
                End If
            End If
        End If
        ilRet = btrGetNext(hmBof, tmBof, imBofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilUpper = UBound(tmRBofRec)
    If ilUpper >= 1 Then
        ArraySortTyp fnAV(tmRBofRec(), 0), UBound(tmRBofRec), 0, LenB(tmRBofRec(0)), 0, LenB(tmRBofRec(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmRBofRec) - 1 Step 1
        ilFound = False
        For ilLast = 0 To UBound(tmSplitNetLastFill) - 1 Step 1
            If tmSplitNetLastFill(ilLast).iFillLen = tmRBofRec(ilLoop).iLen Then
                ilFound = True
                Exit For
            End If
        Next ilLast
        If Not ilFound Then
            tmSplitNetLastFill(UBound(tmSplitNetLastFill)).iBofIndex = -1
            tmSplitNetLastFill(UBound(tmSplitNetLastFill)).iFillLen = tmRBofRec(ilLoop).iLen
            ReDim Preserve tmSplitNetLastFill(0 To UBound(tmSplitNetLastFill) + 1) As SPLITNETLASTFILL
        End If
    Next ilLoop
    mObtainSplitReplacments = True
    Exit Function
End Function

Private Function mCreateSplitFill(ilLen As Integer, tlInOdf As ODF, tlOutOdf As ODF) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLast                        ilLenStart                    slCartNo                  *
'*  slProduct                     slISCI                        slCreativeTitle           *
'*  llCrfCsfCode                  llCpfCode                     llLstCode                 *
'*                                                                                        *
'******************************************************************************************

    'tlInOdf(I)- Split Network spot to find fill for

    Dim ilBof As Integer
    Dim ilLoop As Integer
    Dim ilDayOk As Integer
    Dim ilStartBOF As Integer
    Dim ilRet As Integer
    Dim ilLastAssign As Integer
    'Dim slLen As String
    'Dim ilLen As Integer
    Dim llBofStartDate As Long
    Dim llBofEndDate As Long
    Dim llBofStartTime As Long
    Dim llBofEndTime As Long
    Dim slLength As String

    Dim llAirDate As Long
    Dim llAirTime As Long

    If UBound(tmRBofRec) <= LBound(tmRBofRec) Then
        mCreateSplitFill = False
        Exit Function
    End If
    If UBound(tmSplitNetLastFill) <= LBound(tmSplitNetLastFill) Then
        mCreateSplitFill = False
        Exit Function
    End If
    gUnpackDateLong tlInOdf.iAirDate(0), tlInOdf.iAirDate(1), llAirDate
    gUnpackTimeLong tlInOdf.iAirTime(0), tlInOdf.iAirTime(1), True, llAirTime
    'gUnpackLength tlInOdf.iLen(0), tlInOdf.iLen(1), "1", True, slLen
    'ilLen = Val(slLen)
    ilLastAssign = -1
    For ilLoop = 0 To UBound(tmSplitNetLastFill) - 1 Step 1
        If tmSplitNetLastFill(ilLoop).iFillLen = ilLen Then
            ilLastAssign = ilLoop
            Exit For
        End If
    Next ilLoop
    If ilLastAssign = -1 Then
        mCreateSplitFill = False
        Exit Function
    End If
    ilLoop = tmSplitNetLastFill(ilLastAssign).iBofIndex + 1
    If (ilLoop >= UBound(tmRBofRec)) Or (ilLoop < LBound(tmRBofRec)) Then
        ilLoop = LBound(tmRBofRec)
    End If
    ilStartBOF = ilLoop
    ilBof = -1
    Do
        If (tmRBofRec(ilLoop).iLen = ilLen) And ((tmRBofRec(ilLoop).tBof.iVefCode = tlInOdf.iVefCode) Or (tmRBofRec(ilLoop).tBof.iVefCode = 0)) Then
            'Check Dates, Times and Days
            gUnpackDateLong tmRBofRec(ilLoop).tBof.iStartDate(0), tmRBofRec(ilLoop).tBof.iStartDate(1), llBofStartDate
            gUnpackDateLong tmRBofRec(ilLoop).tBof.iEndDate(0), tmRBofRec(ilLoop).tBof.iEndDate(1), llBofEndDate
            If (llAirDate >= llBofStartDate) And (llAirDate <= llBofEndDate) Then
                gUnpackTimeLong tmRBofRec(ilLoop).tBof.iStartTime(0), tmRBofRec(ilLoop).tBof.iStartTime(1), False, llBofStartTime
                gUnpackTimeLong tmRBofRec(ilLoop).tBof.iEndTime(0), tmRBofRec(ilLoop).tBof.iEndTime(1), True, llBofEndTime
                If (llAirTime >= llBofStartTime) And (llAirTime <= llBofEndTime) Then
                    ilDayOk = False
                    If tmRBofRec(ilLoop).tBof.sDays(gWeekDayLong(llAirDate)) = "Y" Then
                        ilDayOk = True
                    End If
                    If ilDayOk Then
                        tmSplitNetLastFill(ilLastAssign).iBofIndex = ilLoop
                        ilBof = ilLoop
                        Exit Do
                    End If
                End If
            End If
        End If
        ilLoop = ilLoop + 1
        If ilLoop >= UBound(tmRBofRec) Then
            ilLoop = LBound(tmRBofRec)
        End If
        If ilLoop = ilStartBOF Then
            Exit Do
        End If
    Loop
    If ilBof = -1 Then
        mCreateSplitFill = False
        Exit Function
    End If

    tlOutOdf = tlInOdf
    tlOutOdf.iAdfCode = tmRBofRec(ilBof).tBof.iAdfCode
    tlOutOdf.lCifCode = tmRBofRec(ilBof).tBof.lCifCode
    tmChfSrchKey0.lCode = tmRBofRec(ilBof).tBof.lRChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mCreateSplitFill = False
        Exit Function
    End If
    tlOutOdf.lCntrNo = tmChf.lCntrNo
    tlOutOdf.sProduct = tmChf.sProduct
    tlOutOdf.lchfcxfCode = tmChf.lCxfCode
    tlOutOdf.sBBDesc = ""
    tlOutOdf.sShortTitle = ""
    tlOutOdf.sSplitNetwork = ""
    slLength = Trim$(str$(ilLen)) & "s"
    gPackLength slLength, tlOutOdf.iLen(0), tlOutOdf.iLen(1)
    mCreateSplitFill = True
End Function

Private Function mMergeWithLog(ilVefCode As Integer) As Integer
    Dim ilVff As Integer
    
    mMergeWithLog = True
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If ilVefCode = tgVff(ilVff).iVefCode Then
            If tgVff(ilVff).sMergeTraffic = "S" Then
                mMergeWithLog = False
            End If
            Exit For
        End If
    Next ilVff

End Function
Private Function mMergeWithAffiliate(ilVefCode As Integer) As Integer
    Dim ilVff As Integer
    
    mMergeWithAffiliate = True
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If ilVefCode = tgVff(ilVff).iVefCode Then
            If tgVff(ilVff).sMergeAffiliate = "S" Then
                mMergeWithAffiliate = False
            End If
            Exit For
        End If
    Next ilVff

End Function

Private Function mGenLogForAirVehicle(tlAirVef As VEF, ilDeleteSTF As Integer, slGenVehName As String, slStartDate As String, slEndDate As String, slCopyStartDate As String, ilZoneExist As Integer, slTZStartDate As String, slTZEndDate As String, slTZStartTime As String, slTZEndTime As String, ilGenLST As Integer, ilExportType As Integer, ilODFVefCode As Integer, ilLSTForLogVeh As Integer, ilMergeExist As Integer, blTestMerge As Boolean) As Integer
    Dim ilLink As Integer
    Dim ilRet As Integer
    Dim llTestDate As Long
    Dim slStartTime As String  'Start Time String
    Dim slEndTime As String    'End Time String
    Dim llCDate As Long
    Dim ilType As Integer
    Dim sLCP As String
    Dim ilCRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    ReDim ilEvtAllowed(0 To 14) As Integer
    Dim tlVef As VEF
    
    
    ilType = 0
    sLCP = "C"
    
    For ilLoop = LBound(ilEvtAllowed) To UBound(ilEvtAllowed) Step 1
        ilEvtAllowed(ilLoop) = True
    Next ilLoop
    ilEvtAllowed(0) = False 'Don't include library names
    
    slStartTime = Trim$(edcTime(0).Text)
    slEndTime = Trim$(edcTime(1).Text)
    
    gBuildLinkArray hmVLF, tlAirVef, slStartDate, igSVefCode() 'Build igSVefCode so that gBuildODFSpotDay can use it
    If (ckcAssignCopy.Value = vbChecked) And (Not rbcLogType(4).Value) Then
        'For ilLink = LBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) To UBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) Step 1
        '    If tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink) > 0 Then
        For ilLink = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                tmVefSrchKey.iCode = igSVefCode(ilLink) 'tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink)
                ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                On Error GoTo mGenLogErr
                gBtrvErrorMsg ilRet, "mGenLog (btrGetEqual)", Logs
                On Error GoTo 0
                If ilDeleteSTF = vbYes Then
                    smStatusCaption = "Deleting Commercial Changes for " & tlVef.sName
                    plcStatus.Cls
                    plcStatus_Paint
                    tmStfSrchKey.iVefCode = tlVef.iCode
                    gPackDate slStartDate, tmStfSrchKey.iLogDate(0), tmStfSrchKey.iLogDate(1)
                    tmStfSrchKey.iLogTime(0) = 0
                    tmStfSrchKey.iLogTime(1) = 0
                    ilRet = btrGetGreaterOrEqual(hmStf, tmStf, imStfRecLen, tmStfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmStf.iVefCode = tlVef.iCode)
                        gUnpackDateLong tmStf.iLogDate(0), tmStf.iLogDate(1), llTestDate
                        If (llTestDate >= gDateValue(slStartDate)) And (llTestDate <= gDateValue(slEndDate)) Then
                            If tmStf.sPrint = "R" Then
                                ilRet = btrGetPosition(hmStf, lmStfRecPos)
                                Do
                                    'tmRec = tmStf
                                    'ilRet = gGetByKeyForUpdate("STF", hmStf, tmRec)
                                    'tmStf = tmRec
                                    'On Error GoTo mGenLogErr
                                    'gBtrvErrorMsg ilRet, "mGenLog (Get by Key)", Logs
                                    'On Error GoTo 0
                                    tmStf.sPrint = "D"
                                    ilRet = btrUpdate(hmStf, tmStf, imStfRecLen)
                                    If ilRet = BTRV_ERR_CONFLICT Then
                                        ilCRet = btrGetDirect(hmStf, tmStf, imStfRecLen, lmStfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                    End If
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                On Error GoTo mGenLogErr
                                gBtrvErrorMsg ilRet, "mGenLog (btrUpdate)", Logs
                                On Error GoTo 0
                            End If
                        Else
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmStf, tmStf, imStfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                smStatusCaption = "Assigning Copy for " & Trim$(tlVef.sName)
                plcStatus.Cls
                plcStatus_Paint
                'Assign copy
                If rbcLogType(0).Value Then
                    'ilRet = gAssignCopyToSpots(slType, tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink), 0, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                    ilRet = gAssignCopyToSpots(ilType, igSVefCode(ilLink), 0, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                Else
                    'ilRet = gAssignCopyToSpots(slType, tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink), 1, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                    ilRet = gAssignCopyToSpots(ilType, igSVefCode(ilLink), 1, slCopyStartDate, slEndDate, slStartTime, slEndTime)
                End If
                If Not ilRet Then
                    gUserActivityLog "E", slGenVehName & ": Log Generation"
                    imTerminate = True
                    mGenLogForAirVehicle = False
                    Exit Function
                End If
                Do
                    tmVpfSrchKey.iVefKCode = igSVefCode(ilLink)  'tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink)
                    ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    On Error GoTo mGenLogErr
                    gBtrvErrorMsg ilRet, "mGenLog (btrGetEqual)", Logs
                    On Error GoTo 0
                    gUnpackDateLong tmVpf.iLLastDateCpyAsgn(0), tmVpf.iLLastDateCpyAsgn(1), llCDate
                    If gDateValue(slEndDate) > llCDate Then
                        gPackDate slEndDate, tmVpf.iLLastDateCpyAsgn(0), tmVpf.iLLastDateCpyAsgn(1)
                        ilRet = btrUpdate(hmVpf, tmVpf, imVpfRecLen)
                    Else
                        ilRet = BTRV_ERR_NONE
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilRet = gBinarySearchVpf(tmVpf.iVefKCode)
                If ilRet <> -1 Then
                    tgVpf(ilRet) = tmVpf
                End If

                If ilZoneExist Then
                    'Assign copy
                    If rbcLogType(0).Value Then
                        ilRet = gAssignCopyToSpots(ilType, igSVefCode(ilLink), 0, slTZStartDate, slTZEndDate, slTZStartTime, slTZEndTime)
                    Else
                        ilRet = gAssignCopyToSpots(ilType, igSVefCode(ilLink), 1, slTZStartDate, slTZEndDate, slTZStartTime, slTZEndTime)
                    End If
                    If Not ilRet Then
                        gUserActivityLog "E", slGenVehName & ": Log Generation"
                        imTerminate = True
                        mGenLogForAirVehicle = False
                        Exit Function
                    End If
                End If
        '    End If
        Next ilLink
        'Change any new assigned Selling to Airing
        gAlertVehicleReplace hmVLF
    End If
    DoEvents
    If imTerminate Then
        gUserActivityLog "E", slGenVehName & ": Log Generation"
        mGenLogForAirVehicle = False
        Exit Function
    End If
    smStatusCaption = "Generating Log for " & Trim$(tlAirVef.sName)
    plcStatus.Cls
    plcStatus_Paint
    ilRet = gBuildODFSpotDay("L", ilType, sLCP, tlAirVef.iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, ilLSTForLogVeh, blTestMerge)
    If Not ilRet Then
        gUserActivityLog "E", slGenVehName & ": Log Generation"
        imTerminate = True
        mGenLogForAirVehicle = False
        Exit Function
    End If
    ilMergeExist = True
    If rbcLogType(0).Value Or rbcLogType(1).Value Then
        'For ilLink = LBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) To UBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) Step 1
        '    If tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink) > 0 Then
        For ilLink = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                Do
                    tmVpfSrchKey.iVefKCode = igSVefCode(ilLink) 'tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink)
                    ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    On Error GoTo mGenLogErr
                    gBtrvErrorMsg ilRet, "mGenLog (btrGetEqual)", Logs
                    On Error GoTo 0
                    If rbcLogType(1).Value Then
                        gUnpackDate tmVpf.iLLD(0), tmVpf.iLLD(1), slDate
                        If gDateValue(slEndDate) > gDateValue(slDate) Then
                            gPackDate slEndDate, tmVpf.iLLD(0), tmVpf.iLLD(1)
                            'gPackDate slSyncDate, tmVpf.iSyncDate(0), tmVpf.iSyncDate(1)
                            'gPackTime slSyncTime, tmVpf.iSyncTime(0), tmVpf.iSyncTime(1)
                            ''tmVpf.iSourceID = tgUrf(0).iRemoteUserID
                        End If
                    End If
                    gPackDate slEndDate, tmVpf.iLPD(0), tmVpf.iLPD(1)
                    ilRet = btrUpdate(hmVpf, tmVpf, imVpfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilRet = gBinarySearchVpf(tmVpf.iVefKCode)
                If ilRet <> -1 Then
                    tgVpf(ilRet) = tmVpf
                End If
            'End If
        Next ilLink
        gAlertVehicleReplace hmVLF
    End If
    mGenLogForAirVehicle = True
    Exit Function
mGenLogErr:
    imTerminate = True
    mGenLogForAirVehicle = False
    Exit Function
End Function

Private Function mObtainAiringClosingDate(tlVef As VEF, ilLoop As Integer, ilFound As Integer) As Long
    Dim slStartDate As String
    Dim ilLink As Integer
    Dim llTempSDate As Long
    Dim llDate As Long
    Dim ilRet As Integer
    
    If Trim$(tgLogSel(ilLoop).sLLD) = "" Then
        slStartDate = Format$(gNow(), "m/d/yy")
    Else
        slStartDate = Trim$(tgLogSel(ilLoop).sLLD)
    End If
    'slStartDate = gIncOneWeek(slStartDate)
    slStartDate = gIncOneDay(slStartDate)
    gBuildLinkArray hmVLF, tlVef, slStartDate, igSVefCode()
    If (UBound(igSVefCode) <= LBound(igSVefCode)) And (Trim$(tgLogSel(ilLoop).sLLD) = "") Then
        tmVlfSrchKey1.iAirCode = tlVef.iCode
        tmVlfSrchKey1.iAirDay = 0
        gPackDate slStartDate, tmVlfSrchKey1.iEffDate(0), tmVlfSrchKey1.iEffDate(1)
        tmVlfSrchKey1.iAirTime(0) = 0
        tmVlfSrchKey1.iAirTime(1) = 0
        tmVlfSrchKey1.iAirPosNo = 0
        ilRet = btrGetGreaterOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = tlVef.iCode) Then
            gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slStartDate
            gBuildLinkArray hmVLF, tlVef, slStartDate, igSVefCode()
        End If
    End If
    If (UBound(igSVefCode) <= LBound(igSVefCode)) Then
        tgLogSel(ilLoop).iStatus = 1
    End If
    llTempSDate = 0
    'For ilLink = LBound(tgVpf(tgLogSel(ilLoop).iVpfIndex).iGLink) To UBound(tgVpf(tgLogSel(ilLoop).iVpfIndex).iGLink) Step 1
    '    If tgVpf(tgLogSel(ilLoop).iVpfIndex).iGLink(ilLink) > 0 Then
    For ilLink = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
            'llDate = mObtainClosingDate(tgVpf(tgLogSel(ilLoop).iVpfIndex).iGLink(ilLink), tgLogSel(ilLoop).lStartDate, tgLogSel(ilLoop).lEndDate, True)
            llDate = mObtainClosingDate(igSVefCode(ilLink), tgLogSel(ilLoop).lStartDate, tgLogSel(ilLoop).lEndDate, True)
            If llDate > 0 Then
                If llTempSDate = 0 Then
                    llTempSDate = llDate
                Else
                    If llDate < llTempSDate Then
                        llTempSDate = llDate
                    End If
                End If
            ElseIf llDate = 0 Then
                ilFound = True
                llTempSDate = 0
                Exit For
            End If
    '    End If
    Next ilLink
    mObtainAiringClosingDate = llTempSDate
End Function

Private Function mInsertCPTT(tlInsertCPTT As CPTT, tlAtt As ATT) As Integer
    Dim tlCPTT As CPTT
    Dim ilRet As Integer
    Do
        tmCPTTSrchKey2.lAtfCode = tlInsertCPTT.lAtfCode
        tmCPTTSrchKey2.iStartDate(0) = tlInsertCPTT.iStartDate(0)
        tmCPTTSrchKey2.iStartDate(1) = tlInsertCPTT.iStartDate(1)
        ilRet = btrGetEqual(hmCPTT, tlCPTT, imCPTTRecLen, tmCPTTSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            tmCPTTSrchKey.lCode = tlCPTT.lCode
            ilRet = btrGetEqual(hmCPTT, tlCPTT, imCPTTRecLen, tmCPTTSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hmCPTT)
            End If
        Else
            Exit Do
        End If
    Loop
    If tlAtt.sServiceAgreement = "Y" Then
        tlInsertCPTT.iStatus = 1
    End If
    ilRet = btrInsert(hmCPTT, tlInsertCPTT, imCPTTRecLen, INDEXKEY0)

End Function

Private Function mSaveAbf(blChangeHold As Boolean) As Integer
    'slCallType: S=Save; C=Change status to G
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim llDate As Long
    Dim llMoDate As Long
    
    If (Not rbcLogType(1).Value) And (Not rbcLogType(2).Value) And (Not rbcLogType(3).Value) Then
        mSaveAbf = True
        Exit Function
    End If
    For llLoop = 0 To UBound(tgAbfInfo) - 1 Step 1
        If tgAbfInfo(llLoop).sStatus = "N" Then
            tmAbfSrchKey2.sStatus = "G"
            tmAbfSrchKey2.iVefCode = tgAbfInfo(llLoop).iVefCode
            tmAbfSrchKey2.iShttCode = 0
            gPackDateLong tgAbfInfo(llLoop).lMondayDate, tmAbfSrchKey2.iMondayDate(0), tmAbfSrchKey2.iMondayDate(1)
            ilRet = btrGetEqual(hmAbf, tmAbf, imAbfRecLen, tmAbfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmAbf.sStatus = "H"
                ilRet = btrUpdate(hmAbf, tmAbf, imAbfRecLen)
                tgAbfInfo(llLoop).lCode = tmAbf.lCode
                tgAbfInfo(llLoop).sStatus = "C"
            End If
        End If
        If tgAbfInfo(llLoop).sStatus = "N" Then
            'Add record
            tmAbf.lCode = 0
            tmAbf.sSource = "L"
            tmAbf.iVefCode = tgAbfInfo(llLoop).iVefCode
            tmAbf.iShttCode = 0
            If blChangeHold Then
                tmAbf.sStatus = "G"
            Else
                tmAbf.sStatus = "H"
            End If
            gPackDateLong tgAbfInfo(llLoop).lMondayDate, tmAbf.iMondayDate(0), tmAbf.iMondayDate(1)
            gPackDateLong tgAbfInfo(llLoop).lStartDate, tmAbf.iGenStartDate(0), tmAbf.iGenStartDate(1)
            gPackDateLong tgAbfInfo(llLoop).lEndDate, tmAbf.iGenEndDate(0), tmAbf.iGenEndDate(1)
            gPackDate Format(Now, "m/d/yy"), tmAbf.iEnteredDate(0), tmAbf.iEnteredDate(1)
            gPackTime Format(Now, "h:mm:ssAM/PM"), tmAbf.iEnteredTime(0), tmAbf.iEnteredTime(1)
            gPackDate "12/31/2069", tmAbf.iCompletedDate(0), tmAbf.iCompletedDate(1)
            gPackTime "12AM", tmAbf.iCompletedTime(0), tmAbf.iCompletedTime(1)
            tmAbf.iUrfCode = tgUrf(0).iCode
            tmAbf.iUstCode = 0
            tmAbf.sUnused = ""
            ilRet = btrInsert(hmAbf, tmAbf, imAbfRecLen, INDEXKEY0)
            tgAbfInfo(llLoop).sStatus = "S"
            tgAbfInfo(llLoop).lCode = tmAbf.lCode
            '3/6/15: Update CPTT to handle case where they are not using the Station Spot Builder
            If blChangeHold Then
                mUpdateCPTTFromAbf tmAbf
            End If
        ElseIf tgAbfInfo(llLoop).sStatus = "C" Then
            'Update record
            Do
                tmAbfSrchKey0.lCode = tgAbfInfo(llLoop).lCode
                ilRet = btrGetEqual(hmAbf, tmAbf, imAbfRecLen, tmAbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    If blChangeHold Then
                        tmAbf.sStatus = "G"
                    Else
                        tmAbf.sStatus = "H"
                    End If
                    gUnpackDateLong tmAbf.iGenStartDate(0), tmAbf.iGenStartDate(1), llDate
                    If tgAbfInfo(llLoop).lStartDate < llDate Then
                        gPackDateLong tgAbfInfo(llLoop).lStartDate, tmAbf.iGenStartDate(0), tmAbf.iGenStartDate(1)
                    End If
                    gUnpackDateLong tmAbf.iGenEndDate(0), tmAbf.iGenEndDate(1), llDate
                    If tgAbfInfo(llLoop).lEndDate > llDate Then
                        gPackDateLong tgAbfInfo(llLoop).lEndDate, tmAbf.iGenEndDate(0), tmAbf.iGenEndDate(1)
                    End If
                    ilRet = btrUpdate(hmAbf, tmAbf, imAbfRecLen)
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            '3/6/15: Update CPTT to handle case where they are not using the Station Spot Builder
            If blChangeHold Then
                mUpdateCPTTFromAbf tmAbf
            End If
        Else
            If blChangeHold Then
                tmAbfSrchKey0.lCode = tgAbfInfo(llLoop).lCode
                ilRet = btrGetEqual(hmAbf, tmAbf, imAbfRecLen, tmAbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    tmAbf.sStatus = "G"
                    ilRet = btrUpdate(hmAbf, tmAbf, imAbfRecLen)
                    '3/6/15: Update CPTT to handle case where they are not using the Station Spot Builder
                    'gUnpackDateLong tmAbf.iMondayDate(0), tmAbf.iMondayDate(1), llMoDate
                    'tmCPTTSrchKey1.iVefCode = tmAbf.iVefCode
                    'gPackDateLong llMoDate, tmCPTTSrchKey1.iStartDate(0), tmCPTTSrchKey1.iStartDate(1)
                    'ilRet = btrGetEqual(hmCPTT, tmCPTT, imCPTTRecLen, tmCPTTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    'Do While (ilRet = BTRV_ERR_NONE) And (tmCPTT.iVefCode = tmAbf.iVefCode)
                    '    gUnpackDateLong tmCPTT.iStartDate(0), tmCPTT.iStartDate(1), llDate
                    '    If llMoDate <> llDate Then
                    '        Exit Do
                    '    End If
                    '    If tmCPTT.sAstStatus = "C" Then
                    '        If (tmAbf.iShttCode = tmCPTT.iShfCode) Or (tmAbf.iShttCode = 0) Then
                    '            If rbcLogType(1).Value Then
                    '                tmCPTT.sAstStatus = "N"
                    '            Else
                    '                tmCPTT.sAstStatus = "R"
                    '            End If
                    '            ilRet = btrUpdate(hmCPTT, tmCPTT, imCPTTRecLen)
                    '        End If
                    '    Else
                    '        ilRet = BTRV_ERR_NONE
                    '    End If
                    '    If ilRet = BTRV_ERR_CONFLICT Then
                    '        tmCPTTSrchKey1.iVefCode = tmAbf.iVefCode
                    '        gPackDateLong llMoDate, tmCPTTSrchKey1.iStartDate(0), tmCPTTSrchKey1.iStartDate(1)
                    '        ilRet = btrGetEqual(hmCPTT, tmCPTT, imCPTTRecLen, tmCPTTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    '    Else
                    '        ilRet = btrGetNext(hmCPTT, tmCPTT, imCPTTRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    '    End If
                    'Loop
                    mUpdateCPTTFromAbf tmAbf
                End If
            End If
        End If
    Next llLoop
    If blChangeHold Then
        ReDim tgAbfInfo(0 To 0) As ABFINFO
    End If
    mSaveAbf = True
End Function

Private Sub mUpdateCPTTFromAbf(tlAbf As ABF)
    Dim llMoDate As Long
    Dim llDate As Long
    Dim ilRet As Integer
    
    gUnpackDateLong tlAbf.iMondayDate(0), tlAbf.iMondayDate(1), llMoDate
    tmCPTTSrchKey1.iVefCode = tlAbf.iVefCode
    gPackDateLong llMoDate, tmCPTTSrchKey1.iStartDate(0), tmCPTTSrchKey1.iStartDate(1)
    ilRet = btrGetEqual(hmCPTT, tmCPTT, imCPTTRecLen, tmCPTTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmCPTT.iVefCode = tlAbf.iVefCode)
        gUnpackDateLong tmCPTT.iStartDate(0), tmCPTT.iStartDate(1), llDate
        If llMoDate <> llDate Then
            Exit Do
        End If
        If tmCPTT.sAstStatus = "C" Then
            If (tlAbf.iShttCode = tmCPTT.iShfCode) Or (tlAbf.iShttCode = 0) Then
                If rbcLogType(1).Value Then
                    tmCPTT.sAstStatus = "N"
                Else
                    tmCPTT.sAstStatus = "R"
                End If
                ilRet = btrUpdate(hmCPTT, tmCPTT, imCPTTRecLen)
            End If
        Else
            ilRet = BTRV_ERR_NONE
        End If
        If ilRet = BTRV_ERR_CONFLICT Then
            tmCPTTSrchKey1.iVefCode = tlAbf.iVefCode
            gPackDateLong llMoDate, tmCPTTSrchKey1.iStartDate(0), tmCPTTSrchKey1.iStartDate(1)
            ilRet = btrGetEqual(hmCPTT, tmCPTT, imCPTTRecLen, tmCPTTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        Else
            ilRet = btrGetNext(hmCPTT, tmCPTT, imCPTTRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        End If
    Loop

End Sub
Private Sub mPopListKey()
    Dim llMaxWidth As Long
    lbcKey.Clear
    lbcKey.AddItem "Background Color"
    lbcKey.AddItem "     Green: Background Station Spot Builder Program Running"
    lbcKey.AddItem "     Red: Background Station Spot Builder Program Not Running"
    lbcKey.AddItem "     Yellow: Unable to Determine Status of the "
    lbcKey.AddItem "             Background Station Spot Builder Program"
    
    Traffic.pbcArial.FontBold = False
    Traffic.pbcArial.FontName = "Arial"
    Traffic.pbcArial.FontBold = False
    Traffic.pbcArial.FontSize = 8
    llMaxWidth = (Traffic.pbcArial.TextWidth("     Red: Background Station Spot Builder Program Not Running")) + 180
    lbcKey.Width = llMaxWidth
    lbcKey.FontBold = False
    lbcKey.FontName = "Arial"
    lbcKey.FontBold = False
    lbcKey.FontSize = 8
    lbcKey.height = (lbcKey.ListCount) * 225
    lbcKey.height = gListBoxHeight(lbcKey.ListCount, 5)
    lbcKey.Move imcKey.Left, imcKey.Top + imcKey.height
End Sub

Public Sub mExportRegionCopyRecord(llOdfCode As Long)
    Dim llLoop As Long
    Dim llOdf As Long
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim blFound As Boolean
    Dim ilShtt As Integer
    Dim ilRsf As Integer
    Dim ilLoop As Integer
    Dim slType As String
    Dim ilAdfCode As Integer
    Dim ilAdf As Integer
    Dim slAdvtName As String
    Dim slCart As String
    Dim slISCI As String
    Dim slProduct As String
    Dim slCreativeTitle As String
    Dim slCallLetters As String
    Dim llStationID As Long
    Dim slRecord As String
    Dim slXDSCue As String
    
    If (tgSpf.sGUseAffSys <> "Y") Then
        Exit Sub
    End If
    For llOdf = 0 To UBound(tgOdfSdfCodes) - 1 Step 1
        If llOdfCode = tgOdfSdfCodes(llOdf).lOdfCode Then
            ReDim tmExportCopyInfo(0 To 0) As EXPORTCOPYINFO
            'Include station list
            ReDim tmStationExportInfo(0 To 0) As STATIONEXPORTINFO
            tmRsfSrchKey1.lCode = tgOdfSdfCodes(llOdf).lSdfCode
            ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tgOdfSdfCodes(llOdf).lSdfCode)
                'Build array of regions spots, then spot in desc RotNo order
                ilUpper = UBound(tmExportCopyInfo)
                tmExportCopyInfo(ilUpper).iRotNo = tmRsf.iRotNo
                tmExportCopyInfo(ilUpper).lRafCode = tmRsf.lRafCode
                tmExportCopyInfo(ilUpper).lCrfCode = tmRsf.lCrfCode
                tmExportCopyInfo(ilUpper).sPtType = tmRsf.sPtType
                tmExportCopyInfo(ilUpper).lCopyCode = tmRsf.lCopyCode
                ReDim Preserve tmExportCopyInfo(0 To ilUpper + 1)
                ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            Loop
            If UBound(tmExportCopyInfo) > 1 Then
                'Descending sort
                ArraySortTyp fnAV(tmExportCopyInfo(), 0), UBound(tmExportCopyInfo) - 1, 1, LenB(tmExportCopyInfo(0)), 0, -1, 0
            End If
            For ilRsf = 0 To UBound(tmExportCopyInfo) - 1 Step 1
                tmSefSrchKey1.lRafCode = tmExportCopyInfo(ilRsf).lRafCode
                tmSefSrchKey1.iSeqNo = 0
                ilRet = btrGetGreaterOrEqual(hmSef, tmSef, imSefRecLen, tmSefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmSef.lRafCode = tmExportCopyInfo(ilRsf).lRafCode)
                    If tmSef.sCategory = "S" Then
                        If tmSef.sInclExcl <> "E" Then
                            blFound = False
                            For ilShtt = 0 To UBound(tmStationExportInfo) - 1 Step 1
                                If tmStationExportInfo(ilShtt).iShttCode = tmSef.iIntCode Then
                                    blFound = True
                                    If tmExportCopyInfo(ilRsf).iRotNo >= tmStationExportInfo(ilShtt).iRotNo Then
                                        tmStationExportInfo(ilShtt).iRotNo = tmExportCopyInfo(ilRsf).iRotNo
                                        tmStationExportInfo(ilShtt).sSource = "I"
                                        tmStationExportInfo(ilShtt).sExport = "Y"
                                    End If
                                    Exit For
                                End If
                            Next ilShtt
                            If Not blFound Then
                                ilUpper = UBound(tmStationExportInfo)
                                tmStationExportInfo(ilUpper).iShttCode = tmSef.iIntCode
                                tmStationExportInfo(ilUpper).iRotNo = tmExportCopyInfo(ilRsf).iRotNo
                                tmStationExportInfo(ilUpper).sSource = "I"
                                tmStationExportInfo(ilUpper).sExport = "Y"
                                tmStationExportInfo(ilUpper).iExportCopyInfoIndex = ilRsf
                                ReDim Preserve tmStationExportInfo(0 To ilUpper + 1) As STATIONEXPORTINFO
                            End If
                        Else
                            'Add all station except the one to be excluded
                            For ilLoop = LBound(tgStations) To UBound(tgStations) - 1 Step 1
                                blFound = False
                                For ilShtt = 0 To UBound(tmStationExportInfo) - 1 Step 1
                                    If tmStationExportInfo(ilShtt).iShttCode = tgStations(ilLoop).iCode Then
                                        blFound = True
                                        If tmExportCopyInfo(ilRsf).iRotNo >= tmStationExportInfo(ilShtt).iRotNo Then
                                            tmStationExportInfo(ilShtt).iRotNo = tmExportCopyInfo(ilRsf).iRotNo
                                            tmStationExportInfo(ilShtt).sSource = "E"
                                            If tmStationExportInfo(ilShtt).iShttCode = tmSef.iIntCode Then
                                                tmStationExportInfo(ilShtt).sExport = "N"
                                            Else
                                                tmStationExportInfo(ilShtt).sExport = "Y"
                                            End If
                                        ElseIf (tmStationExportInfo(ilShtt).sSource = "E") And (tmStationExportInfo(ilShtt).sExport = "Y") Then
                                            tmStationExportInfo(ilShtt).sExport = "N"
                                        End If
                                        Exit For
                                    End If
                                Next ilShtt
                                If Not blFound Then
                                    ilUpper = UBound(tmStationExportInfo)
                                    tmStationExportInfo(ilUpper).iShttCode = tmSef.iIntCode
                                    tmStationExportInfo(ilUpper).iRotNo = tmExportCopyInfo(ilRsf).iRotNo
                                    tmStationExportInfo(ilUpper).sSource = "E"
                                    If tmStationExportInfo(ilUpper).iShttCode = tgStations(ilLoop).iCode Then
                                        tmStationExportInfo(ilUpper).sExport = "N"
                                    Else
                                        tmStationExportInfo(ilUpper).sExport = "Y"
                                    End If
                                    tmStationExportInfo(ilUpper).iExportCopyInfoIndex = ilRsf
                                    ReDim Preserve tmStationExportInfo(0 To ilUpper + 1) As STATIONEXPORTINFO
                                End If
                            Next ilLoop
                        End If
                    End If
                    ilRet = btrGetNext(hmSef, tmSef, imSefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            Next ilRsf
            'Write out all copy records
            'Copy:,Type,Advertiser Name,Advertiser Abbreviation,Cart,ISCI,Product,Creative Title,Call Letters,Station ID
            For ilShtt = 0 To UBound(tmStationExportInfo) - 1 Step 1
                ilRsf = tmStationExportInfo(ilShtt).iExportCopyInfoIndex
                tmCrfSrchKey.lCode = tmExportCopyInfo(ilRsf).lCrfCode
                ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If tmCrf.iBkoutInstAdfCode <= 0 Then
                        slType = "S"    'Standard region
                        ilAdfCode = tmCrf.iAdfCode
                    Else
                        slType = "B"    'Blackout Region
                        ilAdfCode = tmCrf.iBkoutInstAdfCode
                    End If
                    slAdvtName = "AdvertiserName, "
                    ilAdf = gBinarySearchAdf(ilAdfCode)
                    If ilAdf <> -1 Then
                        slAdvtName = """" & Trim$(tgCommAdf(ilAdf).sName) & """" & "," & """" & Trim$(tgCommAdf(ilAdf).sAbbr) & """"
                    End If
                    tmAxfSrchKey1.iCode = ilAdfCode
                    ilRet = btrGetEqual(hmAxf, tmAxf, imAxfRecLen, tmAxfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        slXDSCue = Trim$(tmAxf.sXDSCue)
                    Else
                        slXDSCue = ""
                    End If
                    tmCifSrchKey0.lCode = tmExportCopyInfo(ilRsf).lCopyCode
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        tmMcfSrchKey0.iCode = tmCif.iMcfCode
                        ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            tmMcf.sName = "C"
                            tmMcf.sPrefix = "C"
                        End If
                        If Trim$(tmCif.sCut) = "" Then
                            slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & " "
                        Else
                            slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut) & " "
                        End If
                        tmCpfSrchKey0.lCode = tmCif.lcpfCode     'product/isci/creative title
                        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            slCreativeTitle = """" & Trim$(tmCpf.sCreative) & """"
                            slISCI = """" & Trim$(tmCpf.sISCI) & """"
                            slProduct = """" & Trim$(tmCpf.sName) & """"
                        End If
                        For ilLoop = LBound(tgStations) To UBound(tgStations) - 1 Step 1
                            If tmStationExportInfo(ilShtt).iShttCode = tgStations(ilLoop).iCode Then
                                slCallLetters = Trim$(tgStations(ilLoop).sCallLetters)
                                llStationID = tgStations(ilLoop).lPermStationID
                                Exit For
                            End If
                        Next ilLoop
                        slRecord = "Copy:," & slType & "," & slAdvtName & ","
                        slRecord = slRecord & slCart & "," & slISCI & "," & slProduct & "," & slCreativeTitle & ","
                        slRecord = slRecord & slCallLetters & "," & llStationID & "," & slXDSCue
                        Print #hmTo, slRecord
                    End If
                End If
            Next ilShtt
            Exit For
        End If
    Next llOdf
    On Error Resume Next
    Erase tmExportCopyInfo
    Erase tmStationExportInfo
End Sub


