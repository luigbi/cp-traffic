VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ExpISCIXRef 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6870
   ClientLeft      =   195
   ClientTop       =   1590
   ClientWidth     =   9135
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
   ScaleHeight     =   6870
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar pgcProgress 
      Height          =   225
      Left            =   3600
      TabIndex        =   54
      Top             =   6000
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton rbcFormat 
      Caption         =   "ISCI"
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   52
      Top             =   1200
      Width           =   735
   End
   Begin VB.OptionButton rbcFormat 
      Caption         =   "Break"
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   51
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "&Browse"
      Height          =   345
      Left            =   6105
      TabIndex        =   50
      Top             =   1665
      Width           =   930
   End
   Begin VB.TextBox edcFileName 
      Height          =   330
      Left            =   1710
      TabIndex        =   48
      Top             =   1665
      Width           =   4095
   End
   Begin VB.PictureBox plcCalendarToLog 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   6255
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   40
      Top             =   3345
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton cmcCalUpToLog 
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
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalDnToLog 
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
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendarToLog 
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
         Picture         =   "ExpISCIXRef.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   240
         Width           =   1875
         Begin VB.Label lacDateToLog 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   42
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalNameToLog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   45
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcCalendarLog 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   660
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   34
      Top             =   2535
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendarLog 
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
         Left            =   30
         Picture         =   "ExpISCIXRef.frx":2E1A
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   1875
         Begin VB.Label lacDateLog 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   38
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDnLog 
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
         Left            =   30
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalUpLog 
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
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalNameLog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   345
         TabIndex        =   39
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcCalendarRot 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   2835
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   17
      Top             =   2910
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton cmcCalUpRot 
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
         Left            =   1620
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalDnRot 
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendarRot 
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
         Picture         =   "ExpISCIXRef.frx":5C34
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDateRot 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   19
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalNameRot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   345
         TabIndex        =   16
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcCalendarToRot 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   6720
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   25
      Top             =   2475
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendarToRot 
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
         Picture         =   "ExpISCIXRef.frx":8A4E
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDateToRot 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   20
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDnToRot 
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalUpToRot 
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
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalNameToRot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   360
         TabIndex        =   29
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.CheckBox ckcGeneric 
      Caption         =   "Export &Generic Copy"
      Height          =   270
      Left            =   3045
      TabIndex        =   6
      Top             =   1185
      Width           =   2415
   End
   Begin VB.CheckBox ckcRegional 
      Caption         =   "Export &Regional Copy"
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   1185
      Width           =   2535
   End
   Begin VB.CommandButton cmcEndDateLog 
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
      Left            =   6600
      Picture         =   "ExpISCIXRef.frx":B868
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   810
      Width           =   195
   End
   Begin VB.TextBox edcEndDateLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   3
      Top             =   810
      Width           =   930
   End
   Begin VB.CommandButton cmcStartDateLog 
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
      Left            =   2505
      Picture         =   "ExpISCIXRef.frx":B962
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   810
      Width           =   195
   End
   Begin VB.TextBox edcStartDateLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1590
      MaxLength       =   10
      TabIndex        =   4
      Top             =   810
      Width           =   930
   End
   Begin VB.CommandButton cmcEndDateRot 
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
      Left            =   7695
      Picture         =   "ExpISCIXRef.frx":BA5C
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   375
      Width           =   195
   End
   Begin VB.TextBox edcEndDateRot 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6780
      MaxLength       =   10
      TabIndex        =   2
      Top             =   390
      Width           =   930
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1410
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2475
      Top             =   4545
   End
   Begin VB.ListBox lbcMsg 
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
      Height          =   2760
      Left            =   3720
      MultiSelect     =   2  'Extended
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2595
      Width           =   5235
   End
   Begin VB.CommandButton cmcStartDateRot 
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
      Left            =   3735
      Picture         =   "ExpISCIXRef.frx":BB56
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   375
      Width           =   195
   End
   Begin VB.TextBox edcStartDateRot 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2790
      MaxLength       =   10
      TabIndex        =   1
      Top             =   375
      Width           =   930
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
      Height          =   2760
      ItemData        =   "ExpISCIXRef.frx":BC50
      Left            =   120
      List            =   "ExpISCIXRef.frx":BC52
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   2610
      Width           =   3375
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   6255
      Width           =   1050
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
      Left            =   4785
      TabIndex        =   10
      Top             =   6255
      Width           =   1050
   End
   Begin VB.Label lacFormat 
      Caption         =   "Format:"
      Height          =   255
      Left            =   5880
      TabIndex        =   53
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Export File Name"
      Height          =   300
      Left            =   120
      TabIndex        =   49
      Top             =   1725
      Width           =   1560
   End
   Begin VB.Label Label2 
      Caption         =   "Results"
      Height          =   315
      Left            =   5505
      TabIndex        =   47
      Top             =   2175
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Vehicles"
      Height          =   360
      Left            =   1320
      TabIndex        =   46
      Top             =   2205
      Width           =   900
   End
   Begin VB.Label lacEndDateLog 
      Appearance      =   0  'Flat
      Caption         =   "Log End Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4275
      TabIndex        =   32
      Top             =   795
      Width           =   1260
   End
   Begin VB.Label lacStartDateLog 
      Appearance      =   0  'Flat
      Caption         =   "Log Start Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   30
      Top             =   795
      Width           =   1455
   End
   Begin VB.Label lacScreen 
      Caption         =   "Audio ISCI Title"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   2520
   End
   Begin VB.Label lacEndDateRot 
      Appearance      =   0  'Flat
      Caption         =   "Rotation Entered End Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4275
      TabIndex        =   13
      Top             =   375
      Width           =   2430
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   180
      TabIndex        =   24
      Top             =   5625
      Width           =   8730
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   4500
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   195
      TabIndex        =   22
      Top             =   5415
      Width           =   8730
   End
   Begin VB.Label lacStartDateRot 
      Appearance      =   0  'Flat
      Caption         =   "Rotation Entered Start Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   375
      Width           =   2535
   End
End
Attribute VB_Name = "ExpISCIXRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpISCIXRef.Frm
'
' Release: 1.0
'
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim hmMsg As Integer
Dim lmNowDate As Long
Dim smMessageFilePath As String
Dim hmExport As Integer
Dim lmStartRot As Long
Dim lmEndRot As Long
Dim lmStartLog As Long
Dim lmEndLog As Long
Private Type crfInvExt
    lCode As Long
    iAdfCode As Integer
    lChfCode As Long
    iVefCode As Integer
    lRafCode As Long
    lEndDate(0 To 1) As Integer
    '7557 add blackout adv
    iBlackoutAdv As Integer
End Type
'7557 and I after last L
'Const crfInvExtPK = "LILILL"
Const crfInvExtPK = "LILILLL"
Enum myExpISCICalendars
    rotationStart = 0
    RotationEnd = 1
    logStart = 2
    logEnd = 3
    NoneChosen = 4
End Enum
Enum myExpISCIMessages
    ExpISCIReset
    ExpISCIFailure
    ExpISCIStart
    ExpISCIFinished
    ExpISCIProcess
    ExpISCIFinishedNothingToWrite
    ExpISCIProgress
    ExpISCIWrite
End Enum
Enum myExpISCIFailures
    ExpISCINoMessageFile = 1
    ExpISCIUnknown = 2
    ExpISCINoFailure = 0
    ExpISCINoExportFile = 3
End Enum
Const EXPISCIFILETITLE As String = "ISCICrossReference.csv"
Const FUTUREDATE = "12/30/2050"

'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim imAdfRecLen As Integer  'ANF record length
'Copy Rotation
Dim hmCrf As Integer        'CRF Handle
Dim imCrfRecLen As Integer      'CRF record length
Dim tmCrf As CRF
'5/19/15
Private imCrfVefCode() As Integer
'Copy/Product
Dim hmCpf As Integer
Dim tmCpf As CPF
Dim imCpfRecLen As Integer  'CPF record length
'Copy instruction record information
Dim hmCnf As Integer        'Copy instruction file handle
Dim imCnfRecLen As Integer  'CNF record length
Dim tmCnf As CNF            'CNF record image
'Copy inventory record information
Dim hmCif As Integer        'Copy line file handle
Dim imCifRecLen As Integer  'CIF record length
Dim tmCif As CIF            'CIF record image
'5/19/15: Copy Vehicles
Dim hmCvf As Integer        'Contract header file handle
Dim tmCvf As CVF            'CVF record image
Dim tmCvfSrchKey As LONGKEY0  'CVF key record image
Dim tmCvfSrchKey1 As LONGKEY0  'CVF key record image
Dim imCvfRecLen As Integer      'CVF record length
''Contract record information
Dim hmCHF As Integer        'Contract header file handle
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
Dim hmClf As Integer
Dim imClfRecLen As Integer
Dim tmClf As CLF
''Short Title Vehicle Table record information  change to vff
Dim hmVsf As Integer        'Short Title Vehicle Table file handle
Dim imVsfRecLen As Integer  'VSF record length
Dim tmVsf As VSF            'VSF record image
Dim hmSif As Integer        'Short Title Vehicle Table file handle
Dim imSifRecLen As Integer  'VSF record length
Dim tmSif As SIF            'VSF record image
' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim imVefRecLen As Integer     'VEF record length
Dim hmMcf As Integer
Dim tmMcf As MCF
Dim imMcfRecLen As Integer  'mcf record length
Dim hmVff As Integer        'Vehicle file handle
Dim tmVff As VFF            'VfF record image
Dim imVffRecLen As Integer     'VfF record length
Dim hmRaf As Integer
Dim tmRaf As RAF
Dim imRafRecLen As Integer  'raf record length
Dim hmVlf As Integer        '
Dim tmVlf As VLF            'vlf record image
Dim imVlfRecLen As Integer     'vlf record length

Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer

Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)

' MsgBox parameters
Const vbOkOnly = 0                 ' OK button only
Const vbCritical = 16          ' Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0


'EVENTS

Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    If lbcVehicle.ListCount <= 0 Then
        Exit Sub
    End If
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
        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicle.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.NoneChosen)
End Sub

Private Sub ckcGeneric_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.NoneChosen)
End Sub

Private Sub ckcGeneric_Validate(Cancel As Boolean)
    mAtLeastOneCheckbox ckcGeneric.Name
End Sub

Private Sub ckcRegional_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.NoneChosen)
End Sub

Private Sub ckcRegional_Validate(Cancel As Boolean)
    mAtLeastOneCheckbox ckcRegional.Name
End Sub

'CALENDAR/DATE EVENTS
Private Sub cmcCalDnLog_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendarLog_Paint
    edcStartDateLog.SelStart = 0
    edcStartDateLog.SelLength = Len(edcStartDateLog.Text)
    edcStartDateLog.SetFocus
End Sub

Private Sub cmcCalDnToLog_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendarToLog_Paint
    edcEndDateLog.SelStart = 0
    edcEndDateLog.SelLength = Len(edcEndDateLog.Text)
    edcEndDateLog.SetFocus
End Sub

Private Sub cmcCalUpLog_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendarLog_Paint
    edcStartDateLog.SelStart = 0
    edcStartDateLog.SelLength = Len(edcStartDateLog.Text)
    edcStartDateLog.SetFocus
End Sub

Private Sub cmcCalUpToLog_Click()
    mSetCalendarVisibility (myExpISCICalendars.logEnd)
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendarToLog_Paint
    edcEndDateLog.SelStart = 0
    edcEndDateLog.SelLength = Len(edcEndDateLog.Text)
    edcEndDateLog.SetFocus
End Sub
Private Sub cmcEndDateLog_Click()
    plcCalendarToLog.Visible = Not plcCalendarToLog.Visible
    edcEndDateLog.SelStart = 0
    edcEndDateLog.SelLength = Len(edcEndDateLog.Text)
    edcEndDateLog.SetFocus
    mSetCommands
End Sub

Private Sub cmcEndDateLog_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.logEnd)
    gCtrlGotFocus ActiveControl
    mSetCommands
End Sub
Private Sub cmcStartDateLog_Click()
    plcCalendarLog.Visible = Not plcCalendarLog.Visible
    edcStartDateLog.SelStart = 0
    edcStartDateLog.SelLength = Len(edcStartDateLog.Text)
    edcStartDateLog.SetFocus
    mSetCommands
End Sub
Private Sub cmcStartDateLog_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetCommands
End Sub

Private Sub cmcBrowse_Click()
    CommonDialog1.ShowSave
    If CommonDialog1.fileName <> "" Then
        edcFileName.Text = CommonDialog1.fileName
        mSetCommands
    End If
End Sub

Private Sub edcEndDateLog_Change()
    Dim slStr As String
    
    slStr = edcEndDateLog.Text
    If Not gValidDate(slStr) Then
        lacDateToLog.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendarToLog_Paint   'mBoxCalDate called within paint
    mSetCommands
End Sub
Private Sub edcEndDateLog_Click()
    mSetCalendarVisibility (myExpISCICalendars.logEnd)
    mSetCommands
End Sub

Private Sub edcEndDateLog_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.logEnd)
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDateLog
    mSetCommands
End Sub

Private Sub edcEndDateLog_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    mSetCommands
End Sub

Private Sub edcEndDateLog_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcEndDateLog.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    mSetCommands
End Sub

Private Sub edcEndDateLog_KeyUp(KeyCode As Integer, Shift As Integer)
Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendarToLog.Visible = Not plcCalendarToLog.Visible
        Else
            slDate = edcEndDateLog.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDateLog.Text = slDate
            End If
        End If
        edcEndDateLog.SelStart = 0
        edcEndDateLog.SelLength = Len(edcEndDateLog.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcEndDateLog.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDateLog.Text = slDate
            End If
        End If
        edcEndDateLog.SelStart = 0
        edcEndDateLog.SelLength = Len(edcEndDateLog.Text)
    End If
    mSetCommands
End Sub

Private Sub edcFileName_Change()
    mSetCommands
End Sub

Private Sub edcStartDateLog_Change()
    Dim slStr As String
    slStr = edcStartDateLog.Text
    If Not gValidDate(slStr) Then
        lacDateLog.Visible = False
        mEnableEndDate False, myExpISCICalendars.logEnd
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendarLog_Paint   'mBoxCalDate called within paint
    mEnableEndDate True, myExpISCICalendars.logEnd
    mSetCommands
End Sub
Private Sub edcStartDateLog_Click()
    mSetCommands
End Sub
Private Sub edcStartDateLog_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.logStart)
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDateLog
    mSetCommands
End Sub
Private Sub edcStartDateLog_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    mSetCommands
End Sub
Private Sub edcStartDateLog_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcStartDateLog.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    mSetCommands
End Sub
Private Sub edcStartDateLog_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendarLog.Visible = Not plcCalendarLog.Visible
        Else
            slDate = edcStartDateLog.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDateLog.Text = slDate
            End If
        End If
        edcStartDateLog.SelStart = 0
        edcStartDateLog.SelLength = Len(edcStartDateLog.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcStartDateLog.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDateLog.Text = slDate
            End If
        End If
        edcStartDateLog.SelStart = 0
        edcStartDateLog.SelLength = Len(edcStartDateLog.Text)
    End If
    mSetCommands
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    '5/19/15
    On Error Resume Next
    
    Erase imCrfVefCode
    
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmCnf)
    btrDestroy hmCnf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    '5/19/15
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    ilRet = btrClose(hmVff)
    btrDestroy hmVff
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    
    Set ExpISCIXRef = Nothing   'Remove data segment
    
End Sub

Private Sub lacEndDateLog_Click()
    mSetCommands
End Sub

Private Sub lacStartDateLog_Click()
    mSetCommands
End Sub

Private Sub lbcMsg_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.NoneChosen)
End Sub

Private Sub pbcCalendarLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                edcStartDateLog.Text = Format$(llDate, "m/d/yy")
                edcStartDateLog.SelStart = 0
                edcStartDateLog.SelLength = Len(edcStartDateLog.Text)
                imBypassFocus = True
                edcStartDateLog.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcStartDateLog.SetFocus
End Sub
Private Sub pbcCalendarLog_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalNameLog.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendarLog, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate edcStartDateLog, lacDateLog
End Sub

Private Sub pbcCalendarToLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                edcEndDateLog.Text = Format$(llDate, "m/d/yy")
                edcEndDateLog.SelStart = 0
                edcEndDateLog.SelLength = Len(edcEndDateLog.Text)
                imBypassFocus = True
                edcEndDateLog.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcEndDateLog.SetFocus
End Sub
Private Sub pbcCalendarToLog_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalNameToLog.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendarToLog, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate edcEndDateLog, lacDateToLog
End Sub
Private Sub cmcCalDnRot_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendarRot_Paint
    edcStartDateRot.SelStart = 0
    edcStartDateRot.SelLength = Len(edcStartDateRot.Text)
    edcStartDateRot.SetFocus
End Sub

Private Sub cmcCalDnToRot_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendarToRot_Paint
    edcEndDateRot.SelStart = 0
    edcEndDateRot.SelLength = Len(edcEndDateRot.Text)
    edcEndDateRot.SetFocus
End Sub

Private Sub cmcCalUpRot_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendarRot_Paint
    edcStartDateRot.SelStart = 0
    edcStartDateRot.SelLength = Len(edcStartDateRot.Text)
    edcStartDateRot.SetFocus
End Sub

Private Sub cmcCalUpToRot_Click()
    mSetCalendarVisibility (myExpISCICalendars.RotationEnd)
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendarToRot_Paint
    edcEndDateRot.SelStart = 0
    edcEndDateRot.SelLength = Len(edcEndDateRot.Text)
    edcEndDateRot.SetFocus
End Sub
Private Sub cmcEndDateRot_Click()
    plcCalendarToRot.Visible = Not plcCalendarToRot.Visible
    edcEndDateRot.SelStart = 0
    edcEndDateRot.SelLength = Len(edcEndDateRot.Text)
    edcEndDateRot.SetFocus
    mSetCommands
End Sub

Private Sub cmcEndDateRot_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.RotationEnd)
    gCtrlGotFocus ActiveControl
    mSetCommands
End Sub
Private Sub cmcStartDateRot_Click()
    plcCalendarRot.Visible = Not plcCalendarRot.Visible
    edcStartDateRot.SelStart = 0
    edcStartDateRot.SelLength = Len(edcStartDateRot.Text)
    edcStartDateRot.SetFocus
    mSetCommands
End Sub
Private Sub cmcStartDateRot_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetCommands
End Sub

Private Sub edcEndDateRot_Change()
    Dim slStr As String
    mSetCalendarVisibility (myExpISCICalendars.RotationEnd)
    slStr = edcEndDateRot.Text
    If Not gValidDate(slStr) Then
        lacDateToRot.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendarToRot_Paint   'mBoxCalDate called within paint
    mSetCommands
End Sub
Private Sub edcEndDateRot_Click()
    mSetCalendarVisibility (myExpISCICalendars.RotationEnd)
    mSetCommands
End Sub

Private Sub edcEndDateRot_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.RotationEnd)
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDateRot
    mSetCommands
End Sub

Private Sub edcEndDateRot_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    mSetCommands
End Sub

Private Sub edcEndDateRot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcEndDateRot.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    mSetCommands
End Sub

Private Sub edcEndDateRot_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendarToRot.Visible = Not plcCalendarToRot.Visible
        Else
            slDate = edcEndDateRot.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDateRot.Text = slDate
            End If
        End If
        edcEndDateRot.SelStart = 0
        edcEndDateRot.SelLength = Len(edcEndDateRot.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcEndDateRot.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDateRot.Text = slDate
            End If
        End If
        edcEndDateRot.SelStart = 0
        edcEndDateRot.SelLength = Len(edcEndDateRot.Text)
    End If
    mSetCommands
End Sub
Private Sub edcStartDateRot_Change()
    Dim slStr As String
    slStr = edcStartDateRot.Text
    If Not gValidDate(slStr) Then
        lacDateRot.Visible = False
        mEnableEndDate False, myExpISCICalendars.RotationEnd
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendarRot_Paint   'mBoxCalDate called within paint
    mEnableEndDate True, RotationEnd
    mSetCommands
End Sub
Private Sub edcStartDateRot_Click()
    mSetCommands
End Sub
Private Sub edcStartDateRot_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.rotationStart)
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDateRot
    mSetCommands
End Sub
Private Sub edcStartDateRot_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    mSetCommands
End Sub
Private Sub edcStartDateRot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcStartDateRot.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    mSetCommands
End Sub
Private Sub edcStartDateRot_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendarRot.Visible = Not plcCalendarRot.Visible
        Else
            slDate = edcStartDateRot.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDateRot.Text = slDate
            End If
        End If
        edcStartDateRot.SelStart = 0
        edcStartDateRot.SelLength = Len(edcStartDateRot.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcStartDateRot.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDateRot.Text = slDate
            End If
        End If
        edcStartDateRot.SelStart = 0
        edcStartDateRot.SelLength = Len(edcStartDateRot.Text)
    End If
    mSetCommands
End Sub
Private Sub lacEndDateRot_Click()
    mSetCommands
End Sub

Private Sub lacStartDateRot_Click()
    mSetCommands
End Sub
Private Sub pbcCalendarRot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                edcStartDateRot.Text = Format$(llDate, "m/d/yy")
                edcStartDateRot.SelStart = 0
                edcStartDateRot.SelLength = Len(edcStartDateRot.Text)
                imBypassFocus = True
                edcStartDateRot.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcStartDateRot.SetFocus
End Sub
Private Sub pbcCalendarRot_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalNameRot.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendarRot, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate edcStartDateRot, lacDateRot
End Sub

Private Sub pbcCalendarToRot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                edcEndDateRot.Text = Format$(llDate, "m/d/yy")
                edcEndDateRot.SelStart = 0
                edcEndDateRot.SelLength = Len(edcEndDateRot.Text)
                imBypassFocus = True
                edcEndDateRot.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcEndDateRot.SetFocus
End Sub
Private Sub pbcCalendarToRot_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalNameToRot.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendarToRot, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate edcEndDateRot, lacDateToRot
End Sub
'END CALENDAR/DATE EVENTS
Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.NoneChosen)
End Sub

Private Sub cmcExport_Click()
    mExportMain
End Sub
Private Sub cmcExport_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.NoneChosen)
End Sub



Private Sub Form_Activate()

    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.NoneChosen)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        mSetCalendarVisibility (myExpISCICalendars.NoneChosen)
        gFunctionKeyBranch KeyCode
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
        tmcCancel.Enabled = True
        Me.Left = 2 * Screen.Width  'move off the screen so screen won't flash
    End If
End Sub


Private Sub lbcVehicle_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
    End If
    mSetCommands
End Sub
Private Sub lbcVehicle_GotFocus()
    mSetCalendarVisibility (myExpISCICalendars.NoneChosen)
End Sub


Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False
    cmcCancel_Click
End Sub
'END EVENTS
'PRIVATE METHODS
Private Sub mSetCalendarVisibility(ilShowThis As myExpISCICalendars)
    Select Case ilShowThis
        Case myExpISCICalendars.rotationStart
            plcCalendarToRot.Visible = False
            plcCalendarLog.Visible = False
            plcCalendarToLog.Visible = False
        Case myExpISCICalendars.RotationEnd
            plcCalendarRot.Visible = False
            plcCalendarLog.Visible = False
            plcCalendarToLog.Visible = False
        Case myExpISCICalendars.logStart
            plcCalendarRot.Visible = False
            plcCalendarToRot.Visible = False
            plcCalendarToLog.Visible = False
        Case myExpISCICalendars.logEnd
            plcCalendarRot.Visible = False
            plcCalendarToRot.Visible = False
            plcCalendarLog.Visible = False
        Case Else
            plcCalendarRot.Visible = False
            plcCalendarToRot.Visible = False
            plcCalendarLog.Visible = False
            plcCalendarToLog.Visible = False
    End Select
End Sub
Private Sub mAtLeastOneCheckbox(slBoxName As String)
    
    If ckcGeneric.Value <> vbChecked And ckcRegional.Value <> vbChecked Then
        If StrComp(slBoxName, ckcGeneric.Name, vbTextCompare) = 0 Then
            ckcRegional.Value = vbChecked
        Else
            ckcGeneric.Value = vbChecked
        End If
    End If
End Sub
Private Function mFixFileName(slFileName As String) As String
    Dim slExt As String
    Dim blExtExists As Boolean
    
    slFileName = mSetPath(slFileName)
    slExt = ""
    blExtExists = True
    If InStr(slFileName, ".") = 0 Then 'no extension specified
        blExtExists = False
    End If
    If blExtExists Then
        slExt = ""
    Else
        slExt = ".csv"
    End If
    mFixFileName = slFileName & slExt
End Function
   Private Function mSetPath(ByRef slFileName As String) As String
      Dim slReptDest As String
      If (InStr(slFileName, ":") = 0) And (Left(slFileName, 2) <> "\\") Then
         slReptDest = sgExportPath & slFileName
      Else
         slReptDest = slFileName
      End If
      mSetPath = slReptDest
   End Function
Private Function mOpenExportFile(slMessage As String) As Boolean
    Dim ilRet As Integer
    Dim slStr As String
    Dim slDateTime As String
    
    mOpenExportFile = True
    'On Error GoTo mOpenExportFileErr
    ilRet = 0
    slStr = mFixFileName(Trim$(edcFileName.Text))
    'slDateTime = FileDateTime(slStr)
    ilRet = gFileExist(slStr)
    If ilRet = 0 Then
        Kill slStr
    End If
    ilRet = 0
    'hmExport = FreeFile
    'Open slStr For Output As hmExport
    ilRet = gFileOpen(slStr, "Output", hmExport)
    If ilRet <> 0 Then
        Close #hmExport
        slMessage = str$(ilRet)
        ''gMsgBox "Open " & slStr & ", Error #" & str$(ilRet), -1 , ""
        gAutomationAlertAndLogHandler "Open " & slStr & ", Error #" & str$(ilRet), , "Open Error"
        mOpenExportFile = False
        Exit Function
    End If
    Exit Function
'mOpenExportFileErr:
'    ilRet = Err.Number
'    Resume Next

End Function
Private Sub mSetMessages(ilAction As myExpISCIMessages, Optional ilFailureType As myExpISCIFailures, Optional slMessage As String)
   'errors sent to lbcMsg.
   'writing to message file must be ok
    Dim slTab As String
    Dim llNoRec As Long
    
    slTab = "     "
    Select Case ilAction
        Case myExpISCIMessages.ExpISCIReset
            lacProcessing.Caption = ""
            lacMsg.Caption = ""
            lbcMsg.Clear
            pgcProgress.Visible = False
            lbcMsg.ForeColor = vbGreen
        Case myExpISCIMessages.ExpISCIFailure
            lbcMsg.ForeColor = vbRed
            lacProcessing.Visible = False
            lacProcessing.Caption = " "
            pgcProgress.Visible = False
            lacMsg.Caption = "Messages sent to " & smMessageFilePath
            Select Case ilFailureType
                Case myExpISCIFailures.ExpISCINoMessageFile
                    lbcMsg.AddItem ("error opening message file: error#" & slMessage)
                Case myExpISCIFailures.ExpISCINoExportFile
                    lbcMsg.AddItem ("error opening export file: error#" & slMessage)
                Case myExpISCIFailures.ExpISCIUnknown
                    lbcMsg.AddItem (slMessage)
            End Select
        Case myExpISCIMessages.ExpISCIStart
        On Error Resume Next
            llNoRec = btrRecords(hmCrf) 'Obtain number of records
            If llNoRec > 0 Then
                pgcProgress.Value = 1
                pgcProgress.Visible = True
                pgcProgress.Max = llNoRec
                lacProcessing.Caption = "Getting Rotations: 1 of " & llNoRec
            Else
                lacProcessing.Caption = "Getting Rotations"
            End If
            'Print #hmMsg, "** Export ISCI CROSS REFERENCE: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            'gAutomationAlertAndLogHandler "** Export ISCI CROSS REFERENCE: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            'Print #hmMsg, slTab & "Getting Rotations"
            gAutomationAlertAndLogHandler slTab & "Getting Rotations"
        Case myExpISCIMessages.ExpISCIProcess
            pgcProgress.Visible = False
            lacProcessing.Caption = slMessage
            'Print #hmMsg, slTab & slMessage
            gAutomationAlertAndLogHandler slTab & slMessage
        Case myExpISCIMessages.ExpISCIWrite
            llNoRec = Val(slMessage)
            If llNoRec > 0 Then
                pgcProgress.Value = 1
                pgcProgress.Visible = True
                pgcProgress.Max = llNoRec
                lacProcessing.Caption = "Writing copy information: 1 of " & llNoRec
             Else
                lacProcessing.Caption = "Writing copy information"
           End If
        Case myExpISCIMessages.ExpISCIProgress
        'it increments itself by 1.
            On Error Resume Next
            pgcProgress.Value = pgcProgress.Value + 1
            On Error GoTo 0
            If pgcProgress.Value >= pgcProgress.Max Then
                pgcProgress.Visible = False
                lacProcessing.Caption = "Continuing"
            Else
                If InStr(1, lacProcessing.Caption, "Writing copy information") > 0 Then
                    lacProcessing.Caption = "writing copy information: " & pgcProgress.Value & " of " & pgcProgress.Max
                Else
                    lacProcessing.Caption = "Getting Rotations: " & pgcProgress.Value & " of " & pgcProgress.Max
                End If
            End If
        Case myExpISCIMessages.ExpISCIFinished
            lacProcessing.Caption = ""
            lbcMsg.AddItem ("Completed")
            pgcProgress.Visible = False
            lacMsg.Caption = "Messages sent to " & smMessageFilePath
            'Print #hmMsg, "** Completed Export ISCI CROSS REFERENCE: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            gAutomationAlertAndLogHandler "** Completed Export ISCI CROSS REFERENCE: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        Case myExpISCIMessages.ExpISCIFinishedNothingToWrite
            lacProcessing.Caption = ""
            lacMsg.Caption = "Messages sent to " & smMessageFilePath
            lbcMsg.AddItem ("No records to export.")
            pgcProgress.Visible = False
            'Print #hmMsg, "No records exported."
            gAutomationAlertAndLogHandler "No records exported."
            'Print #hmMsg, "** Completed Export ISCI CROSS REFERENCE: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            gAutomationAlertAndLogHandler "** Completed Export ISCI CROSS REFERENCE: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"

    End Select
End Sub
Private Sub mExportMain()
'type,adv rotationdate,regionName created earlier in routine
'product,cart,isci,creative and filename created later in routine, when exporting
    Dim ilRet As Integer
    Dim slMessage As String
    Dim blEmpty As Boolean
    Dim tlInputExt As crfInvExt
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim llRecPos As Long
    ' user selected vehicles, limited by crf vehicle, or vehicles in airing/package crf vehicle
    Dim myVehRs As ADODB.Recordset
    ' crf and selected vehicles unique information to export for each cif
    Dim myUniqueRs As ADODB.Recordset
    Dim myExportRs As ADODB.Recordset
    Dim llRet As Long 'why use this?
    Dim llCifCodes() As Long
    Dim c As Integer
    Dim slAdvName As String * 30
    Dim slRegionName As String * 80
    Dim slShortTitle As String * 45
    Dim slISCIPrefix As String * 6
    Dim slISCI As String * 26
    Dim slFileName As String
    Dim slCifName As String 'part of cartnumber
    Dim slCrfType As String
    Dim llCpfCode As Long
    Dim ilMcfCode As Integer
    Dim llRotationDate As Long
    Dim slRotationDate As String
    Dim blUseShortTitle As Boolean
    Dim slInfo As String
    Dim slComma As String
    '7549
    Dim blIsNationalModel As Boolean
    '5/19/15
    Dim ilVef As Integer
    '7557
    Dim blIsBlackout As Boolean
    Dim slProductReturn As String
    
    Const BLACKOUTTEST As String = "Dan_And-Dick--BlACKOUttesT"
    
    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    
    Screen.MousePointer = vbHourglass
    If Not mOpenMsgFile(slMessage) Then
        mExpISCIResetSelectivity
        cmcCancel.SetFocus
        mSetMessages myExpISCIMessages.ExpISCIFailure, myExpISCIFailures.ExpISCINoMessageFile, slMessage
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    gAutomationAlertAndLogHandler "Export ISCICrossReference"
    gAutomationAlertAndLogHandler " * RotationEntStartDate = " & edcStartDateRot.Text
    gAutomationAlertAndLogHandler " * RotationEntEndDate = " & edcEndDateRot.Text
    gAutomationAlertAndLogHandler " * LogStartDate = " & edcStartDateLog.Text
    gAutomationAlertAndLogHandler " * LogEndDate = " & edcEndDateLog.Text
    If ckcRegional.Value = vbChecked Then
        gAutomationAlertAndLogHandler " * Export Regional Copy = True"
    Else
        gAutomationAlertAndLogHandler " * Export Regional Copy = False"
    End If
    If ckcGeneric.Value = vbChecked Then
        gAutomationAlertAndLogHandler " * Export Generic Copy = True"
    Else
        gAutomationAlertAndLogHandler " * Export Generic Copy = false"
    End If
    If rbcFormat(0).Value = True Then gAutomationAlertAndLogHandler " * Format = Break"
    If rbcFormat(1).Value = True Then gAutomationAlertAndLogHandler " * Format = ISCI"
    If ckcAll.Value = vbChecked Then
        gAutomationAlertAndLogHandler " * All Vehicles = True"
    Else
        gAutomationAlertAndLogHandler " * All Vehicles = False"
    End If
    
    
    imExporting = True
    mSetMessages ExpISCIStart
    slComma = ","
    lmStartLog = 0
    lmEndLog = 0
    lmStartRot = 0
    lmEndRot = 0
    blEmpty = True
    slAdvName = String(30, " ")
    slRegionName = String(80, " ")
    slShortTitle = String(45, " ")
    slISCIPrefix = String(6, " ")
    slISCI = String(26, " ")
    
    Set myVehRs = mFillVehRecordset
    Set myExportRs = mPrepRecordset(True)
    '7933
    If rbcFormat(1).Value Then
        blIsNationalModel = True
    Else
        blIsNationalModel = False
    End If
   ' blIsNationalModel = mSetAsIsci()
On Error GoTo CRFExtendErr
    ilExtLen = mPrepCrfExtend()
    llNoRec = gExtNoRec(ilExtLen)
    ilRet = mLimitCrfExtend(llNoRec)
    If ilRet = BTRV_ERR_NONE Then
        ilRet = btrExtGetNext(hmCrf, tlInputExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo CRFExtendErr
            gBtrvErrorMsg ilRet, "mExportMain (btrExtGetNextExt):" & "Crf.Btr", ExpISCIXRef
            On Error GoTo 0
            ilExtLen = Len(tlInputExt)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmCrf, tlInputExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                '7557 bring in blackout advertisert and replace if > 0 than adv
                blIsBlackout = False
                If tlInputExt.iBlackoutAdv > 0 Then
                    tlInputExt.iAdfCode = tlInputExt.iBlackoutAdv
                    blIsBlackout = True
                End If
                If Not myUniqueRs Is Nothing Then
                    If (myUniqueRs.State And adStateOpen) <> 0 Then
                        myUniqueRs.Close
                    End If
                End If
                '5/19/15: Moved into the for ilVef loop
                ''not sure why I loop this over and over: I got an error when I tried to just close and reopen
                'Set myUniqueRs = mPrepRecordset(False)
                'slShortTitle = ""
                ''limit vehicles to subset of user selected and crf vehicles
                ''warning: tmchf filled here...used in mGetShortTitleByVehicle below.
                
                '5/19/15
                mObtainCrfVehicle tlInputExt.lCode, tlInputExt.iVefCode
                For ilVef = 0 To UBound(imCrfVefCode) - 1 Step 1
                    tlInputExt.iVefCode = imCrfVefCode(ilVef)
                    'move from above
                    Set myUniqueRs = mPrepRecordset(False)
                    slShortTitle = ""
                    
                    myVehRs.Filter = mVehicleLimitAndFillContract(tlInputExt.iVefCode, tlInputExt.lChfCode)
                    If Not (myVehRs.EOF And myVehRs.BOF) Then
                        myVehRs.MoveFirst
                        slCrfType = mTestType(tlInputExt.lRafCode)
                        If slCrfType = "R" Then
                            slRegionName = mGetRegionName(tlInputExt.lRafCode)
                        Else
                            slRegionName = ""
                        End If
                        
                        'fill tmadf for mgetShortTitleByVehicle
                        slAdvName = mFillAdv(tlInputExt.iAdfCode)
                        Do While Not myVehRs.EOF
                            '7549
                            'slISCIPrefix = mGetISCIPrefixAndSetTitleTest(myVehRs("vefCode"), blUseShortTitle)
                            'Dan M 52615 blUseShortTitle is old code looking at 'p'...old ISCI. it always returns false for cue model, and is no longer used for national
                            slISCIPrefix = mGetISCIPrefixAndSetTitleTest(myVehRs("vefCode"), blUseShortTitle, blIsNationalModel)
                            If Not blIsNationalModel Then
                                'break rules
                                'Dan 5/25/15
                                If ((Asc(tgSpf.sUsingFeatures10) And ADDADVTTOISCI) = ADDADVTTOISCI) Then
                                    slShortTitle = gXDSShortTitle(tmAdf, "", False, False)
                                End If
'                                If (blUseShortTitle) Or ((Asc(tgSpf.sUsingFeatures10) And ADDADVTTOISCI) = ADDADVTTOISCI) Then
'                                    slShortTitle = mGetShortTitleByVehicle(hmVsf, hmSif, tmChf, tmAdf, myVehRs("vefCode"))
'                                    slShortTitle = UCase$(gFileNameFilter(Trim$(slShortTitle)))
'                                    If ((Asc(tgSpf.sUsingFeatures10) And ADDADVTTOISCI) = ADDADVTTOISCI) Then
'                                        '2/7/13: Use Advertiser name only because XDS limited to 32 characters
'                                        'slShortTitle = UCase$(Left$(slShortTitle, 15))
'                                        'Dan M 7219
'                                        slShortTitle = gXDSShortTitle(tmAdf, "", False, False)
'        '                                If Trim$(tmAdf.sAbbr) <> "" Then
'        '                                    slShortTitle = UCase$(mFileNameFilter(Left$(UCase(Trim(tmAdf.sAbbr)), 6)))
'        '                                Else
'        '                                    slShortTitle = UCase(mFileNameFilter(Trim(Left(tmAdf.sName, 6))))
'        '                                End If
'                                    End If
'                                End If
                            Else
                                '7557
                                If blIsBlackout Then
                                    slShortTitle = gXDSShortTitle(tmAdf, BLACKOUTTEST, False, True)
                                Else
                                    slShortTitle = mGetShortTitleByVehicle(hmVsf, hmSif, tmChf, tmAdf, myVehRs("vefCode"))
                                    slShortTitle = UCase$(gFileNameFilter(Trim$(slShortTitle)))
                                End If
                            End If
                            On Error GoTo FINDEXPORTERROR
                            If mUniqueForCrf(slISCIPrefix, slShortTitle, myUniqueRs) Then
                                myUniqueRs.AddNew Array("Prefix", "ShortTitle"), Array(slISCIPrefix, slShortTitle)
                            End If
                            myVehRs.MoveNext
                        Loop
                        ReDim llCifCodes(0)
                        'for each cif that corresponds to crf, test each myUniqueRs + cif info to see if unique: if it is, write to export rs.
                        If Not (myUniqueRs.BOF And myUniqueRs.EOF) Then
                            llRet = mFillCif(tlInputExt.lCode, llCifCodes)
                            If llRet >= 0 Then
                                For c = 0 To llRet
                                    
                                    mGetCifInfo llCifCodes(c), llCpfCode, llRotationDate, ilMcfCode, slCifName, tlInputExt.lEndDate
                                    myUniqueRs.MoveFirst
                                    Do While Not myUniqueRs.EOF
                                        If mUniqueExport(slRegionName, myUniqueRs("prefix"), myUniqueRs("ShortTitle"), llCifCodes(c), llCpfCode, myExportRs) Then
                                            myExportRs.AddNew Array("RegionName", "Prefix", "ShortTitle", "AdvName", "Type", "CpfCode", "RotationDate", "McfCode", "CifName", "CifCode"), Array(slRegionName, myUniqueRs("prefix"), myUniqueRs("ShortTitle"), slAdvName, slCrfType, llCpfCode, llRotationDate, ilMcfCode, slCifName, llCifCodes(c))
                                        Else
                                            'if types different, change to "B"
                                            If StrComp(myExportRs("Type"), slCrfType, vbTextCompare) <> 0 Then
                                                myExportRs("Type") = "B"
                                            End If
                                        End If
                                        myUniqueRs.MoveNext
                                   Loop
                                Next c
                            End If
                        End If
                    End If
                Next ilVef
                ilRet = btrExtGetNext(hmCrf, tlInputExt, ilExtLen, llRecPos)
                mSetMessages ExpISCIProgress
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmCrf, tlInputExt, ilExtLen, llRecPos)
                Loop
            Loop
            myExportRs.Filter = adFilterNone
            If Not (myExportRs.EOF And myExportRs.BOF) Then
                mSetMessages ExpISCIWrite, ExpISCINoFailure, CStr(myExportRs.RecordCount)
                slMessage = ""
                If Not mOpenExportFile(slMessage) Then
                    mSetMessages myExpISCIMessages.ExpISCIFailure, myExpISCIFailures.ExpISCINoExportFile, slMessage
                    GoTo Cleanup
                End If
                On Error GoTo PRINTERR
                Print #hmExport, mWriteExportHeader()
                blEmpty = False
                myExportRs.MoveFirst
                Do While Not myExportRs.EOF
                    ''product,cart,isci,creative..fill slisci for mBuldFileName
                    slInfo = mGetExportInfo(myExportRs("mcfCode"), myExportRs("CpfCode"), myExportRs("Prefix"), myExportRs("cifName"), slISCI, slProductReturn)
                    '7557 with slProductReturn added above, also
                    If InStr(Trim$(myExportRs("ShortTitle")), BLACKOUTTEST) > 0 Then
                        slProductReturn = gFileNameFilter(Trim$(slProductReturn))
                        myExportRs("ShortTitle") = Replace(myExportRs("ShortTitle"), BLACKOUTTEST, UCase$(slProductReturn))
                    End If
                    slFileName = mBuildFileName(myExportRs("ShortTitle"), slISCI)
                    slRotationDate = Format(myExportRs("rotationdate"), "General Date")
                    Print #hmExport, mAddQuotes(myExportRs("Type")) & slComma & mAddQuotes(myExportRs("AdvName")) & slComma & slInfo & slComma & mAddQuotes(slRotationDate) & slComma _
                     & mAddQuotes(myExportRs("regionName")) & slComma & mAddQuotes(slFileName)
                    myExportRs.MoveNext
                    mSetMessages ExpISCIProgress
                Loop
            End If
        End If
    Else
        GoTo CRFExtendErr
    End If
On Error GoTo 0
    If blEmpty Then
        mSetMessages ExpISCIFinishedNothingToWrite
    Else
        mSetMessages ExpISCIFinished
    End If
Cleanup:
On Error Resume Next
    Close #hmMsg
    If Not blEmpty Then
        Close #hmExport
    End If
On Error GoTo 0
    Screen.MousePointer = vbDefault
    imExporting = False
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    mExpISCIResetSelectivity
    If Not myUniqueRs Is Nothing Then
        If (myUniqueRs.State And adStateOpen) <> 0 Then
            myUniqueRs.Close
            Set myUniqueRs = Nothing
        End If
    End If
    If Not myVehRs Is Nothing Then
        If (myVehRs.State And adStateOpen) <> 0 Then
            myVehRs.Close
            Set myVehRs = Nothing
        End If
    End If
    If Not myExportRs Is Nothing Then
        If (myExportRs.State And adStateOpen) <> 0 Then
            myExportRs.Close
            Set myExportRs = Nothing
        End If
    End If
    Erase llCifCodes
    Exit Sub

CRFExtendErr:
   ' gDbg_HandleError "ExptISCIXRef: mExportMain"
    slMessage = "Problem Reading tables and building export. Could not create list of rotations."
    mSetMessages myExpISCIMessages.ExpISCIFailure, myExpISCIFailures.ExpISCIUnknown, slMessage
    imExporting = False
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    mExpISCIResetSelectivity
    Exit Sub
FINDEXPORTERROR:
   ' gDbg_HandleError "ExptISCIXRef: mExportMain"
    slMessage = "Problem Reading tables and building export. File not created."
    mSetMessages myExpISCIMessages.ExpISCIFailure, myExpISCIFailures.ExpISCIUnknown, slMessage
    GoTo Cleanup
    Exit Sub
PRINTERR:
   ' gDbg_HandleError "ExptISCIXRef: mExportMain"
    slMessage = "problem writing records to be exported.  File not created."
    mSetMessages myExpISCIMessages.ExpISCIFailure, myExpISCIFailures.ExpISCIUnknown, slMessage
    GoTo Cleanup
    
ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub
Private Function mBuildFileName(ByVal slShortTitle As String, ByVal slISCI As String) As String
  ' ShortTitle(ISCI).mp2 or ISCI.mp2
    slISCI = UCase$(gFileNameFilter(Trim$(slISCI)))
    If StrComp(slShortTitle, String(45, " "), vbBinaryCompare) = 0 Then
        '7496
        'mBuildFileName = Trim$(slISCI) & ".mp2"
        mBuildFileName = slISCI & UCase(sgAudioExtension)
    Else
        '7496
        'mBuildFileName = Trim$(slShortTitle) & "(" & Trim$(slISCI) & ").mp2"
        mBuildFileName = Trim$(slShortTitle) & "(" & slISCI & ")" & UCase(sgAudioExtension)
    End If
End Function
Private Function mGetExportInfo(ilMcfCode As Integer, llCpfCode As Long, slISCIPrefix As String, slCifName As String, slISCI As String, slProductReturn) As String
'7557 return product
'Private Function mGetExportInfo(ilMcfCode As Integer, llCpfCode As Long, slISCIPrefix As String, slCifName As String, slISCI As String) As String
'product,cart,isci,creative
    Dim ilRet As Integer
    Dim slCreative As String * 30
    Dim slProduct As String * 35
    Dim slCartNumber As String * 11
    Dim slUnfilteredIsci As String
    Dim slComma As String
    
    slComma = ","
    slCreative = ""
    slProduct = ""
    slCartNumber = ""
    slUnfilteredIsci = ""
    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, llCpfCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        slProduct = Trim$(tmCpf.sName)
        slCreative = Trim$(tmCpf.sCreative)
        slUnfilteredIsci = Trim$(tmCpf.sISCI)
        slISCI = gFileNameFilter(Trim$(slISCIPrefix) & slUnfilteredIsci)
    End If
    ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, ilMcfCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        slCartNumber = Trim$(tmMcf.sName) & slCifName
    End If
    slProductReturn = slProduct
    mGetExportInfo = mAddQuotes(slProduct) & slComma & mAddQuotes(slCartNumber) & slComma & mAddQuotes(slISCI) & slComma & mAddQuotes(slCreative)
End Function
Private Sub mGetCifInfo(llCifCode As Long, llCpfCode As Long, llRotationDate As Long, ilMcfCode As Integer, slCifName As String, ilRot() As Integer)
    Dim ilRet As Integer
    
    llCpfCode = 0
    ilMcfCode = 0
    slCifName = " "
    llRotationDate = 0
    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, llCifCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        With tmCif
            llCpfCode = .lcpfCode
            ilMcfCode = .iMcfCode
            slCifName = Trim$(.sName)
            gUnpackDateLong .iRotEndDate(0), .iRotEndDate(1), llRotationDate
        End With
    End If
    'no rotation date in cif? take from crf
    If llRotationDate = 0 Then
        gUnpackDateLong ilRot(0), ilRot(1), llRotationDate
    End If
End Sub
Private Function mUniqueExport(ByVal slRegionName As String, ByVal slPrefix As String, ByVal slShortTitle As String, llCifCode As Long, llCpfCode As Long, myRs As ADODB.Recordset) As Boolean
    Dim slLimit As String
    
    If myRs.BOF And myRs.EOF Then
        mUniqueExport = True
        Exit Function
    End If
    slLimit = "RegionName = '" & mFixQuote(slRegionName) & "' AND Prefix = '" & mFixQuote(slPrefix) & "' AND ShortTitle = '" & mFixQuote(slShortTitle) & "' AND cifCode = " & llCifCode _
    & " AND cpfCode = " & llCpfCode
    myRs.Filter = slLimit
    If myRs.BOF And myRs.EOF Then
        mUniqueExport = True
    Else
        mUniqueExport = False
    End If
End Function
Private Function mUniqueForCrf(ByVal slISCIPrefix As String, ByVal slShortTitle As String, myRs As ADODB.Recordset) As Boolean
        Dim slLimit As String
        
        If myRs.EOF And myRs.BOF Then
            mUniqueForCrf = True
            Exit Function
        End If
        slLimit = "Prefix = '" & mFixQuote(Trim$(slISCIPrefix)) & "' And ShortTitle = '" & mFixQuote(Trim$(slShortTitle)) & "'"
        myRs.Filter = slLimit
        If myRs.EOF And myRs.BOF Then
            mUniqueForCrf = True
        Else
            mUniqueForCrf = False
        End If
End Function
Private Function mTestType(llRafCode As Long) As String
    If llRafCode = 0 Then
        mTestType = "G"
    Else
        mTestType = "R"
    End If
End Function
Private Function mGetISCIPrefixAndSetTitleTest(ilVefCode As Integer, blUseShortTitle As Boolean, blIsNationalModel As Boolean) As String
    Dim ilRet As Integer
    
    blUseShortTitle = False
    ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, ilVefCode, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        If blIsNationalModel Then
            mGetISCIPrefixAndSetTitleTest = Trim$(tmVff.sXDSISCIPrefix)
        Else
            mGetISCIPrefixAndSetTitleTest = Trim$(tmVff.sXDISCIPrefix)
        End If
        If StrComp(tmVff.sXDXMLForm, "P", vbTextCompare) = 0 Then
            blUseShortTitle = True
        End If
    End If
End Function
Private Function mGetRegionName(llRafCode As Long) As String
    Dim ilRet As Integer
    
    ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, llRafCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        mGetRegionName = Trim$(tmRaf.sName)
    End If
End Function
Private Function mFillAdv(ilAdfCode As Integer) As String
    Dim ilRet As Integer
    
    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, ilAdfCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        mFillAdv = Trim$(tmAdf.sName)
    End If
End Function

Private Function mVehicleLimitAndFillContract(ilVefCode As Integer, llChfCode As Long) As String
    Dim slLimit As String
    Dim ilRet As Integer
    Dim ilVehForAiring() As Integer
    Dim blFirstVehicle As Boolean
    Dim slDate As String
    Dim c As Integer
    
    ilRet = mVefType(ilVefCode)
    'package
    If ilRet = 1 Then
        slLimit = mPackageScanContractLines(ilVefCode, llChfCode)
    'airing
    ElseIf ilRet = 2 Then
        blFirstVehicle = True
        slDate = mGetAiringDate()
        gBuildLinkArray hmVlf, tmVef, slDate, ilVehForAiring
        For c = 0 To UBound(ilVehForAiring) - 1
            If blFirstVehicle Then
                slLimit = "VefCode =  " & ilVehForAiring(c)
                blFirstVehicle = False
            Else
                slLimit = slLimit & " OR vefCode =  " & ilVehForAiring(c)
            End If
        Next c
    Else
        slLimit = "VefCode = " & ilVefCode
    End If
    If ilRet <> 1 Then
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, llChfCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            slLimit = ""
        End If
    End If
    Erase ilVehForAiring
    mVehicleLimitAndFillContract = slLimit
End Function
Private Function mGetAiringDate() As String
'use log start date.  If it wasn't set by user, use rot start date.
    Dim llMyDate As Long
    
    If lmStartLog <> 0 Then
        llMyDate = lmStartLog
    Else
        llMyDate = lmStartRot
    End If
    mGetAiringDate = mLongToString(llMyDate)
End Function
Private Function mPackageScanContractLines(ilVefCode As Integer, llChfCode As Long) As String
    Dim ilRet As Integer
    Dim tlMyClf() As CLFLIST
    Dim c As Integer
    Dim i As Integer
    Dim ilMax As Integer
    Dim ilMyPackageLine As Integer
    Dim slLimit As String
    Dim blFirstVehicle As Boolean
    ' test clf line for vehicle know to be package; get its line #; go back through array and search pkLineNo for that line #--these are hidden lines
    ' write out these vehicles:  "vefcode = 1 OR vefcode =2"
    blFirstVehicle = True
    ilRet = gObtainChfClf(hmCHF, hmClf, llChfCode, 0, tmChf, tlMyClf)
    If ilRet = -1 Then
        ilMax = UBound(tlMyClf)
        If ilMax > 0 Then
            ilMax = ilMax - 1
            For c = 0 To ilMax
                If tlMyClf(c).ClfRec.iVefCode = ilVefCode Then
                    ilMyPackageLine = tlMyClf(c).ClfRec.iLine
                    For i = 0 To ilMax
                         If tlMyClf(i).ClfRec.iPkLineNo = ilMyPackageLine Then
                            If blFirstVehicle Then
                                slLimit = "VefCode =  ~" & tlMyClf(i).ClfRec.iVefCode & "~"
                                blFirstVehicle = False
                            Else
                         'does vehicle already exist in string? Use ~ to help search; remove later
                                If InStr(1, slLimit, "~" & tlMyClf(i).ClfRec.iVefCode & "~", vbTextCompare) = 0 Then
                                    slLimit = slLimit & " OR vefCode =  ~" & tlMyClf(i).ClfRec.iVefCode & "~"
                                End If
                            End If
                         End If
                    Next i
                End If
            Next c
        End If
    End If
    Erase tlMyClf
    slLimit = Replace(slLimit, "~", " ")
    mPackageScanContractLines = slLimit
End Function
Private Function mVefType(ilVefCode As Integer) As Integer
    Dim ilRet As Integer
    'change to binary search on array
    mVefType = 0
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, ilVefCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        If StrComp(tmVef.sType, "P", vbTextCompare) = 0 Then
            mVefType = 1
        ElseIf StrComp(tmVef.sType, "A", vbTextCompare) = 0 Then
            mVefType = 2
        End If
    End If
End Function

Private Function mLimitCrfExtend(llNoRec As Long) As Integer
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim tlDateTypeBuff As POPDATETYPE
    Dim tlLongTypeBuff As POPLCODE
    Dim tlStringTypeBuff As POPCHARTYPE
    Dim ilCommand As Integer
    Dim slEndDate As String
    
    Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC", "crfInvExtPK", crfInvExtPK)
    If ckcGeneric.Value = vbChecked Xor ckcRegional.Value = vbChecked Then
        ilOffSet = gFieldOffset("Crf", "CrfRafCode")
        tlLongTypeBuff.lCode = "0"
        If ckcGeneric.Value = vbChecked Then
            ilCommand = BTRV_EXT_EQUAL
        Else
            ilCommand = BTRV_EXT_NOT_EQUAL
        End If
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_INT, ilOffSet, 4, ilCommand, BTRV_EXT_AND, tlLongTypeBuff, 4)
    Else
        ilRet = BTRV_ERR_NONE
    End If
    If ilRet = BTRV_ERR_NONE Then
        If edcStartDateLog.Text <> "" Then
            slEndDate = FUTUREDATE
            If edcEndDateLog.Text <> "" Then
                slEndDate = edcEndDateLog.Text
            End If
            gPackDate edcStartDateLog.Text, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Crf", "CrfEndDate")
            ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            gUnpackDateLong tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1, lmStartLog
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Crf", "CrfStartDate")
            gUnpackDateLong tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1, lmEndLog
            ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        End If
        'write last of logic for log after decide if more logic coming
        If edcStartDateRot.Text <> "" Then
            slEndDate = FUTUREDATE
            If edcEndDateRot.Text <> "" Then
                slEndDate = edcEndDateRot.Text
            End If
            gPackDate edcStartDateRot.Text, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            gUnpackDateLong tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1, lmStartRot
            ilOffSet = gFieldOffset("Crf", "CrfEntryDate")
            ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            gUnpackDateLong tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1, lmEndRot
        End If
        'not active
        tlStringTypeBuff.sType = "D"
        ilOffSet = gFieldOffset("Crf", "CrfState")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlStringTypeBuff, 1)
    'problem adding constant
    Else
        ilRet = 111
    End If
    mLimitCrfExtend = ilRet
End Function
Private Function mPrepCrfExtend() As Integer
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim tlBuffer As crfInvExt
    
    btrExtClear hmCrf
    ilRet = btrGetFirst(hmCrf, tmCrf, imCrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_END_OF_FILE Then
        gBtrvErrorMsg ilRet, "mPrepCrfExtend (btrReset):" & "Crf.Btr", ExpISCIXRef
        Exit Function
    Else
        gBtrvErrorMsg ilRet, "mPrepCrfExtend (btrGetFirst):" & "Crf.Btr", ExpISCIXRef
    End If
    'select
    ilOffSet = gFieldOffset("Crf", "CrfCode")
    ilRet = btrExtAddField(hmCrf, ilOffSet, 4)
    gBtrvErrorMsg ilRet, "mPrepCrfExtend (btrExtAddField):" & "Crf.Btr", ExpISCIXRef
    ilOffSet = gFieldOffset("Crf", "CrfAdfCode")
    ilRet = btrExtAddField(hmCrf, ilOffSet, 2)
    gBtrvErrorMsg ilRet, "mPrepCrfExtend (btrExtAddField):" & "Crf.Btr", ExpISCIXRef
    ilOffSet = gFieldOffset("Crf", "CrfChfCode")
    ilRet = btrExtAddField(hmCrf, ilOffSet, 4)
    gBtrvErrorMsg ilRet, "mPrepCrfExtend (btrExtAddField):" & "Crf.Btr", ExpISCIXRef
    ilOffSet = gFieldOffset("Crf", "CrfVefCode")
    ilRet = btrExtAddField(hmCrf, ilOffSet, 2)
    gBtrvErrorMsg ilRet, "mPrepCrfExtend (btrExtAddField):" & "Crf.Btr", ExpISCIXRef
    ilOffSet = gFieldOffset("Crf", "CrfRafCode")
    ilRet = btrExtAddField(hmCrf, ilOffSet, 4)
    gBtrvErrorMsg ilRet, "mPrepCrfExtend (btrExtAddField):" & "Crf.Btr", ExpISCIXRef
    ilOffSet = gFieldOffset("Crf", "CrfEndDate")
    ilRet = btrExtAddField(hmCrf, ilOffSet, 4)
    gBtrvErrorMsg ilRet, "mPrepCrfExtend (btrExtAddField):" & "Crf.Btr", ExpISCIXRef
    '7557 add blackout adv
    ilOffSet = gFieldOffset("Crf", "CrfBkoutInstAdfCode")
    ilRet = btrExtAddField(hmCrf, ilOffSet, 2)
    gBtrvErrorMsg ilRet, "mPrepCrfExtend (btrExtAddField):" & "Crf.Btr", ExpISCIXRef

     mPrepCrfExtend = Len(tlBuffer)

End Function
Private Function mWriteExportHeader() As String
    Dim slHeader1 As String
    Dim slHeader2 As String
    Dim slQuote As String
    Dim slComma As String
    Dim slStartRot As String
    Dim slEndRot As String
    Dim slStartLog As String
    Dim slEndLog As String
    
    Dim ilDay1 As Integer
    Dim ilDay2 As Integer
    Dim llFutureDate As Long
    
    gPackDate FUTUREDATE, ilDay1, ilDay2
    gUnpackDateLong ilDay1, ilDay2, llFutureDate
    slComma = ","
    slQuote = """"
    slStartRot = "Not Specified"
    slEndRot = "Not Specified"
    slStartLog = "Not Specified"
    slEndLog = "Not Specified"
    If lmStartRot > 0 Then
        slStartRot = mLongToString(lmStartRot)
        If lmEndRot <> 0 And lmEndRot <> llFutureDate Then
            slEndRot = mLongToString(lmEndRot)
        End If
    End If
    If lmStartLog > 0 Then
        slStartLog = mLongToString(lmStartLog)
        If lmEndLog <> 0 And lmEndLog <> llFutureDate Then
            slEndLog = mLongToString(lmEndLog)
        End If
    End If
    slHeader1 = slQuote & "Entered Start Date: " & slStartRot & slQuote & slComma & slQuote & "Entered End Date: " & slEndRot & slQuote & slComma _
    & slQuote & "Log Start Date: " & slStartLog & slQuote & slComma & slQuote & "Log End Date: " & slEndLog & slQuote
    slHeader2 = slQuote & "Type" & slQuote & slComma & slQuote & "Advertiser Name" & slQuote & slComma & slQuote & "Product Name" & slQuote & slComma _
    & slQuote & "Cart Number" & slQuote & slComma & slQuote & "ISCI" & slQuote & slComma & slQuote & "Creative Title" & slQuote & slComma _
    & slQuote & "Latest Rotation Date" & slQuote & slComma & slQuote & "Region Name" & slQuote & slComma & slQuote & "File Name" & slQuote
    
    mWriteExportHeader = slHeader1 & vbCrLf & slHeader2
End Function

Private Function mFillCif(llCrfCode As Long, llCifCodes() As Long) As Long
'return ubound - 1
    Dim ilRet As Integer
    Dim tlcnfSearch As CNFKEY0

    tlcnfSearch.lCrfCode = llCrfCode
    tlcnfSearch.iInstrNo = 0
    'find all cif's for crf in cnf
    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tlcnfSearch, INDEXKEY0, BTRV_LOCK_NONE)
    Do While ilRet = BTRV_ERR_NONE And tmCnf.lCrfCode = llCrfCode
        llCifCodes(UBound(llCifCodes)) = tmCnf.lCifCode
        ReDim Preserve llCifCodes(0 To UBound(llCifCodes) + 1)
        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
     mFillCif = UBound(llCifCodes) - 1
End Function
Private Function mFillVehRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    Dim c As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    
    Set myRs = New ADODB.Recordset
    myRs.Fields.Append "VefCode", adSmallInt
    myRs.Open
    'Dan M 9/14/10 optimize for faster filtering
    myRs("VefCode").Properties("optimize") = True
    For c = 0 To lbcVehicle.ListCount - 1
        If lbcVehicle.Selected(c) Then
            slNameCode = tgUserVehicle(c).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            myRs.AddNew "VefCode", slCode
        End If
    Next c
    Set mFillVehRecordset = myRs
End Function
Private Function mPrepRecordset(blExport As Boolean) As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
    If blExport Then
        With myRs.Fields
            .Append "RegionName", adChar, 80    'from crf
            .Append "Prefix", adChar, 6         'for ISCI
            .Append "ShortTitle", adChar, 45    'for filename
            .Append "CpfCode", adInteger        'for ProductName,ISCI,CreativeTitle
            .Append "AdvName", adChar, 30     'from crf
            .Append "RotationDate", adInteger   'from cif
            .Append "Type", adChar, 1           'from crfs
            .Append "McfCode", adSmallInt       'from cif
            .Append "CifName", adChar, 5        'from cif
            .Append "CifCode", adInteger        'for comparisons
        End With
    Else
        With myRs.Fields
            .Append "Prefix", adChar, 6         'for ISCI
            .Append "ShortTitle", adChar, 45    'for filename
        End With
    End If
    myRs.Open
    'Dan M 9/14/10 optimize for faster filtering
    myRs("ShortTitle").Properties("optimize") = True
    myRs("prefix").Properties("optimize") = True
    Set mPrepRecordset = myRs
End Function

Public Function mFixQuote(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = "'" Then
                sOutStr = sOutStr & "''"
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    mFixQuote = sOutStr
End Function
Private Function mAddQuotes(slvalue As String) As String

    mAddQuotes = """" & Trim$(slvalue) & """"
End Function
Private Function mLongToString(llInpDate As Long) As String
    mLongToString = Format$(llInpDate, "m/d/yy")
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
Private Sub mBoxCalDate(EditDate As Control, LabelDate As Control)
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = EditDate.Text   'edcStartDateRot.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    LabelDate.Caption = slDay
                    LabelDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    LabelDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            LabelDate.Visible = False
        Else
            LabelDate.Visible = False
        End If
    Else
        LabelDate.Visible = False
    End If
End Sub


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
    Dim ilRet As Integer

    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    smMessageFilePath = sgDBPath & "Messages\" & "ExptISCICrossReference.Txt"
    sgMessageFile = smMessageFilePath
    mPrepDialogBox
    mPrepSelectivity
    If Not mPrepTables Then
        Screen.MousePointer = vbDefault
        imTerminate = True
        mTerminate
        Exit Sub
    End If
    
    gCenterStdAlone ExpISCIXRef
    Screen.MousePointer = vbDefault
    
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)

End Sub

Private Sub mPrepDialogBox()
    
    With CommonDialog1
        .DefaultExt = ".csv"
        .fileName = EXPISCIFILETITLE
        .Filter = "Comma |*.csv|All Files (*.*)|*.*"
        .FilterIndex = 1
        .DialogTitle = "ISCI Cross-Reference Export"
        .InitDir = sgExportPath
    End With
End Sub
Private Function mPrepTables() As Boolean
    Dim ilRet As Integer
    
    mPrepTables = True
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imCrfRecLen = Len(tmCrf)
    hmCif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCif, "", sgDBPath & "cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imCifRecLen = Len(tmCif)
    '5/19/15
    hmCvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCvf, "", sgDBPath & "cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imCvfRecLen = Len(tmCvf)
    
    hmCnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCnf, "", sgDBPath & "cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imCnfRecLen = Len(tmCnf)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMcf, "", sgDBPath & "mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmVff = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVff, "", sgDBPath & "vff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imVffRecLen = Len(tmVff)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRaf, "", sgDBPath & "raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imRafRecLen = Len(tmRaf)
    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSif, "", sgDBPath & "sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imSifRecLen = Len(tmSif)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVsf, "", sgDBPath & "vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
On Error GoTo mPrepTableErr
    gBtrvErrorMsg ilRet, "mPrepTables (btrOpen)", ExpISCIXRef
On Error GoTo 0
    imVlfRecLen = Len(tmVlf)


    Exit Function
mPrepTableErr:
    mPrepTables = False
End Function

Private Sub mPrepSelectivity()
    Dim slStr As String
    
    imSetAll = True
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    slStr = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slStr)
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    mPositionCalendars
    pbcCalendarRot_Paint   'mBoxCalDate called within paint
    pbcCalendarToRot_Paint
    mEnableEndDate False, myExpISCICalendars.NoneChosen
    mVehPop
    edcStartDateRot.TabIndex = 1
    edcEndDateRot.TabIndex = 2
    edcStartDateLog.TabIndex = 3
    edcEndDateLog.TabIndex = 4
    ckcRegional.TabIndex = 5
    ckcGeneric.TabIndex = 6
    edcFileName.TabIndex = 7
    cmcBrowse.TabIndex = 8
    lbcVehicle.TabIndex = 9
    ckcAll.TabStop = True
    ckcAll.TabIndex = 10
    cmcExport.TabIndex = 11
    cmcCancel.TabIndex = 12
    mExpISCIResetSelectivity
    '7933
    mEnableFormatButtons
End Sub
Private Sub mExpISCIResetSelectivity()
    Dim i As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVff As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    
    imAllClicked = False
    ckcGeneric.Value = vbChecked
    ckcRegional.Value = vbChecked
    'not sure I need labels set
    lacDateRot.Visible = False
    lacDateToRot.Visible = False
    lacDateLog.Visible = False
    lacDateToLog.Visible = False
    edcStartDateRot.Text = ""
    edcEndDateRot.Text = ""
    edcStartDateLog.Text = ""
    edcEndDateLog.Text = ""
    ckcAll.Value = vbUnchecked
    edcFileName.Text = sgExportPath & EXPISCIFILETITLE
    For i = 0 To lbcVehicle.ListCount - 1
        lbcVehicle.Selected(i) = False
    Next
    For ilLoop = LBound(tgUserVehicle) To UBound(tgUserVehicle) - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If ilVefCode = tgVff(ilVff).iVefCode Then
                If tgVff(ilVff).sExportAudio = "Y" Then
                    lbcVehicle.Selected(ilLoop) = True
                End If
                Exit For
            End If
        Next ilVff
    Next ilLoop
End Sub
Private Sub mPositionCalendars()

    mCommandRightPosition edcStartDateRot, plcCalendarRot
    mCommandRightPosition edcEndDateRot, plcCalendarToRot
    mCommandRightPosition edcStartDateLog, plcCalendarLog
    mCommandRightPosition edcEndDateLog, plcCalendarToLog
    plcCalendarRot.ZOrder
    plcCalendarToRot.ZOrder
    plcCalendarLog.ZOrder
    plcCalendarToLog.ZOrder
    
End Sub
Private Sub mCommandRightPosition(edcCurrent As TextBox, calCurrent As PictureBox)
    
    calCurrent.Left = edcCurrent.Left
    calCurrent.Top = edcCurrent.Top + edcCurrent.Height
End Sub
Private Sub mEnableEndDate(blEnable As Boolean, ilThisBox As myExpISCICalendars)
    
    Select Case ilThisBox
        Case myExpISCICalendars.RotationEnd, myExpISCICalendars.rotationStart
            edcEndDateRot.Enabled = blEnable
            cmcEndDateRot.Enabled = blEnable
        Case myExpISCICalendars.logEnd, myExpISCICalendars.logStart
            edcEndDateLog.Enabled = blEnable
            cmcEndDateLog.Enabled = blEnable
        Case myExpISCICalendars.NoneChosen
            edcEndDateRot.Enabled = blEnable
            edcEndDateLog.Enabled = blEnable
            cmcEndDateRot.Enabled = blEnable
            cmcEndDateLog.Enabled = blEnable
    End Select
End Sub
Private Function mValidDateSequence(ilDates As myExpISCICalendars) As Boolean
    Dim myStartBox As TextBox
    Dim myEndBox As TextBox
    
    mValidDateSequence = True
    Select Case ilDates
        Case myExpISCICalendars.RotationEnd, myExpISCICalendars.rotationStart
            Set myStartBox = edcStartDateRot
            Set myEndBox = edcEndDateRot
        Case myExpISCICalendars.logEnd, myExpISCICalendars.logStart
                Set myStartBox = edcStartDateLog
                Set myEndBox = edcEndDateLog
    End Select
    
    If myEndBox.Enabled And Trim(myEndBox.Text) <> "" Then
        If Trim(myStartBox.Text) <> "" Then
            If IsDate(myStartBox.Text) And IsDate(myEndBox.Text) Then
                If DateDiff("d", myStartBox.Text, myEndBox.Text) >= 0 Then
                    mValidDateSequence = True
                Else
                    mValidDateSequence = False
                End If
            End If
        End If
    End If
End Function
Private Function mValidDates() As Boolean
    
    mValidDates = False
    If Not ((Trim$(edcStartDateRot.Text) = "") And (Trim$(edcStartDateLog.Text) = "")) And (mValidDateSequence(myExpISCICalendars.rotationStart) And mValidDateSequence(myExpISCICalendars.logStart)) Then
        mValidDates = True
    End If
End Function
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
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified: 6/22/10     By:D.Michaelson    *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile(slErrorMessage As String)
'don't display message box: write to log. write title in different routine. don't set mouse. Write to error message
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    
    mOpenMsgFile = True
    'On Error GoTo mOpenMsgFileErr:
    slToFile = smMessageFilePath
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
            'ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                mOpenMsgFile = False
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
                mOpenMsgFile = False
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
            mOpenMsgFile = False
        End If
    End If
    If mOpenMsgFile = False Then
        slErrorMessage = str$(ilRet)
        ''gMsgBox "Open " & slToFile & ", Error #" & str$(ilRet), -1, ""
        gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly, "Open MsgFile Error"
    End If
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
Private Sub mSetCommands()
Dim ilEnabled As Integer
Dim ilLoop As Integer
    
    ilEnabled = False
    If edcFileName.Text <> "" Then
        'at least one vehicle must be selected
        For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
            If lbcVehicle.Selected(ilLoop) Then
                ilEnabled = True
                Exit For
            End If
        Next ilLoop
        If ilEnabled Then
            If Not mValidDates Then
                ilEnabled = False
            End If
        End If
        If ilEnabled Then
            mSetMessages ExpISCIReset
        End If
    End If
    cmcExport.Enabled = ilEnabled
End Sub
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

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpISCIXRef
    igManUnload = NO
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
    Dim ilVff As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    
    ilRet = gPopUserVehicleBox(ExpISCIXRef, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHSPORTMINUELIVE + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)

    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExpISCIXRef
        On Error GoTo 0
    End If


    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'Private Function mFileNameFilter(slInName As String) As String
'Dan 10/27/14 mFileNameFilter to gFileNameFilter

'    'taken from affiliate expt isci xreference
'    Dim slName As String
'    Dim ilPos As Integer
'    Dim ilFound As Integer
'    slName = slInName
'    'Remove " and '
'    Do
'        ilFound = False
'        ilPos = InStr(1, slName, "'", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
'            ilFound = True
'        End If
'    Loop While ilFound
'    Do
'        ilFound = False
'        ilPos = InStr(1, slName, """", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
'            ilFound = True
'        End If
'    Loop While ilFound
'    Do
'        ilFound = False
'        ilPos = InStr(1, slName, "&", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "/", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "\", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "*", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, ":", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "?", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "%", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        'ilPos = InStr(1, slName, """", 1)
'        'If ilPos > 0 Then
'        '    Mid$(slName, ilPos, 1) = "'"
'        '    ilFound = True
'        'End If
'        ilPos = InStr(1, slName, "=", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "+", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "<", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, ">", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "|", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, ";", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "@", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "[", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "]", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "{", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "}", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "^", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, ".", 1)    'If period, use underscore
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "_"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, ",", 1)    'If comma, use underscore
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "_"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, " ", 1)    'If space, use underscore
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "_"
'            ilFound = True
'        End If
'    Loop While ilFound
'    mFileNameFilter = slName
'End Function
Private Function mGetShortTitleByVehicle(hlVsf As Integer, hlSif As Integer, tlChf As CHF, tlAdf As ADF, ilSchVefCode As Integer) As String
    'Copy rotation record information-modeled on gGetShorTitle
    Dim tlVsfSrchKey0 As LONGKEY0 'VSF key record image
    Dim ilVsfReclen As Integer  'VSF record length
    Dim tlVsf As VSF            'VSF record image
    Dim ilRet As Integer
    Dim ilVsf As Integer
    Dim llSifCode As Long
    Dim ilFound As Integer

    ilVsfReclen = Len(tlVsf)
    llSifCode = 0
    If tgSpf.sUseProdSptScr = "P" Then
        If tlChf.lVefCode < 0 Then
            ilFound = False
            tlVsfSrchKey0.lCode = -tlChf.lVefCode
            ilRet = btrGetEqual(hlVsf, tlVsf, ilVsfReclen, tlVsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While ilRet = BTRV_ERR_NONE
                For ilVsf = LBound(tlVsf.iFSCode) To UBound(tlVsf.iFSCode) Step 1
                    If tlVsf.iFSCode(ilVsf) = ilSchVefCode Then
                        ilFound = True
                        If tlVsf.lFSComm(ilVsf) > 0 Then
                            llSifCode = tlVsf.lFSComm(ilVsf)
                        End If
                        Exit For
                    End If
                Next ilVsf
                If ilFound Then
                    Exit Do
                End If
                If tlVsf.lLkVsfCode <= 0 Then
                    Exit Do
                End If
                tlVsfSrchKey0.lCode = tlVsf.lLkVsfCode
                ilRet = btrGetEqual(hlVsf, tlVsf, ilVsfReclen, tlVsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        mGetShortTitleByVehicle = Trim$(gGetProdOrShtTitle(hlSif, llSifCode, tlChf, tlAdf, 6))
    Else
        mGetShortTitleByVehicle = Trim$(gGetProdOrShtTitle(hlSif, llSifCode, tlChf, tlAdf, 6))
    End If
End Function
Private Function mSetAsIsci(Optional ilRet) As Boolean
    '7933 added out ilRet 1 isci, 2 break, 3 both
    Dim blRet As Boolean
  '  Dim ilRet As Integer
    
    ilRet = 0
    blRet = False
    If (Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT Then
        ilRet = 1
    End If
    If (Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT Then
        ilRet = ilRet + 2
    End If
    If ilRet = 1 Then
        blRet = True
    End If
    mSetAsIsci = blRet
End Function

Private Sub mObtainCrfVehicle(llCrfCode As Long, ilCrfVefCode As Integer)
    Dim ilRet As Integer
    Dim ilCvf As Integer
        
    If ilCrfVefCode > 0 Then
        ReDim imCrfVefCode(0 To 1) As Integer
        imCrfVefCode(0) = ilCrfVefCode
        Exit Sub
    End If
    ReDim imCrfVefCode(0 To 0) As Integer
    imCvfRecLen = Len(tmCvf)
    tmCvfSrchKey1.lCode = llCrfCode
    ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmCvf.lCrfCode = llCrfCode)
        For ilCvf = 0 To 99 Step 1
            If tmCvf.iVefCode(ilCvf) > 0 Then
                imCrfVefCode(UBound(imCrfVefCode)) = tmCvf.iVefCode(ilCvf)
                ReDim Preserve imCrfVefCode(0 To UBound(imCrfVefCode) + 1) As Integer
            End If
        Next ilCvf
        ilRet = btrGetNext(hmCvf, tmCvf, imCvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub
Private Sub mEnableFormatButtons()
    Dim ilFormat As Integer
    lacFormat.Visible = False
    rbcFormat(0).Value = True
    rbcFormat(0).Visible = False
    rbcFormat(1).Visible = False
    If mSetAsIsci(ilFormat) Then
        rbcFormat(1).Value = True
    End If
    If ilFormat = 3 Then
        rbcFormat(0).Value = True
        lacFormat.Visible = True
        rbcFormat(0).Visible = True
        rbcFormat(1).Visible = True
        
    End If
End Sub
