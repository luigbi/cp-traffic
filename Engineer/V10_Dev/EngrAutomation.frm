VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form EngrAutomation 
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrAutomation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11790
   Begin VB.Frame frcNotUsed 
      Caption         =   "Not Used"
      Height          =   2370
      Left            =   9855
      TabIndex        =   74
      Top             =   5895
      Visible         =   0   'False
      Width           =   9915
      Begin VB.TextBox edcImportFileFormat 
         Height          =   285
         Left            =   2670
         MaxLength       =   20
         TabIndex        =   79
         Top             =   795
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.TextBox edcImportExt 
         Height          =   285
         Left            =   6885
         MaxLength       =   3
         TabIndex        =   78
         Top             =   795
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox edcServerImportPath 
         Height          =   285
         Left            =   2670
         MaxLength       =   100
         TabIndex        =   77
         Top             =   1290
         Visible         =   0   'False
         Width           =   8385
      End
      Begin VB.TextBox edcClientImportPath 
         Height          =   285
         Left            =   2655
         MaxLength       =   100
         TabIndex        =   76
         Top             =   1800
         Visible         =   0   'False
         Width           =   8385
      End
      Begin VB.TextBox edcDelay 
         Height          =   285
         Left            =   4395
         TabIndex        =   75
         Top             =   330
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lacImportFileFormat 
         Caption         =   "Import File Name Format:"
         Height          =   255
         Left            =   255
         TabIndex        =   84
         Top             =   795
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label lacServerImportPath 
         Caption         =   "Server Import Path:"
         Height          =   255
         Left            =   255
         TabIndex        =   83
         Top             =   1290
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label lacClientImportPath 
         Caption         =   "Client Import Path:"
         Height          =   255
         Left            =   255
         TabIndex        =   82
         Top             =   1800
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label lacImportExt 
         Caption         =   "Extension:"
         Height          =   240
         Left            =   5250
         TabIndex        =   81
         Top             =   465
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lacDelay 
         Caption         =   "Delay Time after Schedule Set to Test if Removed:"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   330
         Visible         =   0   'False
         Width           =   3765
      End
   End
   Begin VB.CommandButton cmcErase 
      Caption         =   "&Erase"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7725
      TabIndex        =   32
      Top             =   6615
      Width           =   1335
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   6045
      TabIndex        =   31
      Top             =   6615
      Width           =   1335
   End
   Begin VB.Frame frcTab 
      Caption         =   "Paths"
      Height          =   4620
      Index           =   1
      Left            =   630
      TabIndex        =   46
      Top             =   1515
      Visible         =   0   'False
      Width           =   11145
      Begin VB.Frame frcEPTest 
         Caption         =   "Test System"
         Height          =   1275
         Left            =   120
         TabIndex        =   68
         Top             =   3045
         Width           =   10965
         Begin VB.TextBox edcClientExportPathTest 
            Height          =   285
            Left            =   2475
            MaxLength       =   100
            TabIndex        =   72
            Top             =   825
            Width           =   8385
         End
         Begin VB.TextBox edcServerExportPathTest 
            Height          =   285
            Left            =   2475
            MaxLength       =   100
            TabIndex        =   70
            Top             =   315
            Width           =   8385
         End
         Begin VB.Label lacClientExportPathTest 
            Caption         =   "Client Export Path:"
            Height          =   255
            Left            =   60
            TabIndex        =   71
            Top             =   825
            Width           =   2340
         End
         Begin VB.Label lacServerExportPathTest 
            Caption         =   "Server Export Path:"
            Height          =   255
            Left            =   60
            TabIndex        =   69
            Top             =   315
            Width           =   2340
         End
      End
      Begin VB.Frame frcEPProd 
         Caption         =   "Production System"
         Height          =   1275
         Left            =   120
         TabIndex        =   63
         Top             =   1650
         Width           =   10965
         Begin VB.TextBox edcServerExportPath 
            Height          =   285
            Left            =   2475
            MaxLength       =   100
            TabIndex        =   65
            Top             =   315
            Width           =   8385
         End
         Begin VB.TextBox edcClientExportPath 
            Height          =   285
            Left            =   2475
            MaxLength       =   100
            TabIndex        =   67
            Top             =   825
            Width           =   8385
         End
         Begin VB.Label lacServerExportPath 
            Caption         =   "Server Export Path:"
            Height          =   255
            Left            =   60
            TabIndex        =   64
            Top             =   315
            Width           =   2340
         End
         Begin VB.Label lacClientExportPath 
            Caption         =   "Client Export Path:"
            Height          =   255
            Left            =   60
            TabIndex        =   66
            Top             =   825
            Width           =   2340
         End
      End
      Begin VB.TextBox edcExportChgFileFormat 
         Height          =   285
         Left            =   5715
         MaxLength       =   20
         TabIndex        =   54
         Top             =   750
         Width           =   2205
      End
      Begin VB.TextBox edcExportDelFileFormat 
         Height          =   285
         Left            =   8790
         MaxLength       =   20
         TabIndex        =   56
         Top             =   750
         Width           =   2205
      End
      Begin VB.TextBox edcDateFormat 
         Height          =   285
         Left            =   2085
         MaxLength       =   20
         TabIndex        =   48
         Top             =   285
         Width           =   945
      End
      Begin VB.TextBox edcTimeFormat 
         Height          =   285
         Left            =   5415
         MaxLength       =   20
         TabIndex        =   50
         Top             =   285
         Width           =   945
      End
      Begin VB.TextBox edcExportExtDel 
         Height          =   285
         Left            =   5100
         MaxLength       =   3
         TabIndex        =   62
         Top             =   1185
         Width           =   660
      End
      Begin VB.TextBox edcExportExtChg 
         Height          =   285
         Left            =   3330
         MaxLength       =   3
         TabIndex        =   60
         Top             =   1185
         Width           =   660
      End
      Begin VB.TextBox edcExportExtNew 
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   58
         Top             =   1185
         Width           =   660
      End
      Begin VB.TextBox edcExportNewFileFormat 
         Height          =   285
         Left            =   2535
         MaxLength       =   20
         TabIndex        =   52
         Top             =   750
         Width           =   2205
      End
      Begin VB.Label Label3 
         Caption         =   "Change:"
         Height          =   255
         Left            =   4950
         TabIndex        =   53
         Top             =   750
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Delete:"
         Height          =   255
         Left            =   8115
         TabIndex        =   55
         Top             =   750
         Width           =   795
      End
      Begin VB.Label lacDateFormat 
         Caption         =   "File Name Date Format:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   285
         Width           =   1845
      End
      Begin VB.Label lacFileNameTimeFormat 
         Caption         =   "File Name Time Format:"
         Height          =   255
         Left            =   3465
         TabIndex        =   49
         Top             =   285
         Width           =   1815
      End
      Begin VB.Label lacExportExtDel 
         Caption         =   "Delete:"
         Height          =   255
         Left            =   4335
         TabIndex        =   61
         Top             =   1185
         Width           =   795
      End
      Begin VB.Label lacExportExtChg 
         Caption         =   "Change:"
         Height          =   255
         Left            =   2565
         TabIndex        =   59
         Top             =   1185
         Width           =   795
      End
      Begin VB.Label lacExportExtNew 
         Caption         =   "Extension- New:"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1185
         Width           =   1605
      End
      Begin VB.Label lacExportFileFormat 
         Caption         =   "Export File Name Format- New"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   750
         Width           =   2565
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "Format"
      Height          =   3630
      Index           =   2
      Left            =   10725
      TabIndex        =   33
      Top             =   4740
      Visible         =   0   'False
      Width           =   11625
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   8115
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   330
         Width           =   60
      End
      Begin VB.PictureBox pbcImportSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   60
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   41
         Top             =   2130
         Width           =   60
      End
      Begin VB.PictureBox pbcImportTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   90
         Left            =   90
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   43
         Top             =   3405
         Width           =   60
      End
      Begin VB.TextBox edcImport 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3180
         TabIndex        =   42
         Top             =   2805
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox edcExport 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3765
         TabIndex        =   38
         Top             =   1245
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.PictureBox pbcExportSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   330
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   37
         Top             =   675
         Width           =   60
      End
      Begin VB.PictureBox pbcExportTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   90
         Left            =   45
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   39
         Top             =   1995
         Width           =   60
      End
      Begin VB.TextBox edcFixedTimeChar 
         Height          =   285
         Left            =   3090
         MaxLength       =   1
         TabIndex        =   35
         Top             =   210
         Width           =   540
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdExport 
         Height          =   1305
         Left            =   510
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   705
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   5
         Cols            =   39
         FixedRows       =   2
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   39
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdImport 
         Height          =   1065
         Left            =   90
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   1879
         _Version        =   393216
         Rows            =   4
         Cols            =   17
         FixedRows       =   2
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   17
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lacFixedTimeChar 
         Caption         =   "Fixed Time Export Symbol (Character):"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   210
         Width           =   3105
      End
      Begin VB.Label lacSchStartCol 
         Caption         =   "Col #"
         Height          =   180
         Left            =   0
         TabIndex        =   44
         Top             =   1470
         Width           =   1650
      End
      Begin VB.Label lacSchNoChars 
         Caption         =   "# Char"
         Height          =   180
         Left            =   -15
         TabIndex        =   45
         Top             =   1725
         Width           =   1125
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "Information"
      Height          =   3825
      Index           =   0
      Left            =   11055
      TabIndex        =   3
      Top             =   4080
      Width           =   11310
      Begin VB.TextBox edcSecFax 
         Height          =   285
         Left            =   870
         MaxLength       =   20
         TabIndex        =   23
         Top             =   3000
         Width           =   2475
      End
      Begin VB.TextBox edcSecEMail 
         Height          =   285
         Left            =   4890
         MaxLength       =   20
         TabIndex        =   25
         Top             =   3000
         Width           =   5295
      End
      Begin VB.TextBox edcPriFax 
         Height          =   285
         Left            =   870
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2100
         Width           =   2475
      End
      Begin VB.TextBox edcPriEMail 
         Height          =   285
         Left            =   4890
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2100
         Width           =   5295
      End
      Begin VB.TextBox edcSecPhone 
         Height          =   285
         Left            =   7710
         MaxLength       =   20
         TabIndex        =   21
         Top             =   2535
         Width           =   2475
      End
      Begin VB.TextBox edcSecContactName 
         Height          =   285
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   19
         Top             =   2535
         Width           =   3765
      End
      Begin VB.TextBox edcPriPhone 
         Height          =   285
         Left            =   7710
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1650
         Width           =   2475
      End
      Begin VB.TextBox edcPriContactName 
         Height          =   285
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   11
         Top             =   1650
         Width           =   3765
      End
      Begin VB.TextBox edcManufacture 
         Height          =   285
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1200
         Width           =   8475
      End
      Begin VB.Frame frcState 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   330
         Left            =   120
         TabIndex        =   26
         Top             =   3465
         Width           =   2220
         Begin VB.OptionButton rbcState 
            Caption         =   "Active"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton rbcState 
            Caption         =   "Dormant"
            Height          =   255
            Index           =   1
            Left            =   975
            TabIndex        =   28
            Top             =   0
            Width           =   990
         End
      End
      Begin VB.TextBox edcName 
         Height          =   285
         Left            =   2535
         MaxLength       =   20
         TabIndex        =   5
         Top             =   345
         Width           =   2475
      End
      Begin VB.TextBox edcDescription 
         Height          =   285
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   7
         Top             =   765
         Width           =   8475
      End
      Begin VB.Label lacSecFax 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   660
      End
      Begin VB.Label lacSecEMail 
         Caption         =   "E-Mail:"
         Height          =   255
         Left            =   3900
         TabIndex        =   24
         Top             =   3000
         Width           =   1035
      End
      Begin VB.Label lacPriFax 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2100
         Width           =   660
      End
      Begin VB.Label lacPriEMail 
         Caption         =   "E-Mail:"
         Height          =   255
         Left            =   3900
         TabIndex        =   16
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label lacSecPhone 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   6465
         TabIndex        =   20
         Top             =   2535
         Width           =   1035
      End
      Begin VB.Label lacSecContactName 
         Caption         =   "Secondary Contact Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2535
         Width           =   2565
      End
      Begin VB.Label lacPriPhone 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   6465
         TabIndex        =   12
         Top             =   1650
         Width           =   1035
      End
      Begin VB.Label lacPriContactName 
         Caption         =   "Primary Contact Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1650
         Width           =   2565
      End
      Begin VB.Label lacManufacture 
         Caption         =   "Manfacturer:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Label lacName 
         Caption         =   "Automation System Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   345
         Width           =   2565
      End
      Begin VB.Label lacDescription 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   765
         Width           =   1380
      End
   End
   Begin VB.Frame frcSelect 
      Caption         =   "Step 1: Select Automation"
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3525
      Begin VB.ComboBox cbcSelect 
         BackColor       =   &H00FFFF80&
         Height          =   315
         ItemData        =   "EngrAutomation.frx":030A
         Left            =   150
         List            =   "EngrAutomation.frx":030C
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3180
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   315
      Top             =   6525
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7170
      FormDesignWidth =   11790
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4395
      TabIndex        =   30
      Top             =   6615
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   2730
      TabIndex        =   29
      Top             =   6615
      Width           =   1335
   End
   Begin ComctlLib.TabStrip tabAuto 
      Height          =   5445
      Left            =   60
      TabIndex        =   2
      Top             =   885
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   9604
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Name"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Export Paths"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Format"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "EngrAutomation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrAutomation - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private imInChg As Integer
Private imBSMode As Integer
Private imAeeCode As Integer
Private smUsedFlag As String
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private smCurrACEStamp As String
Private smCurrADEStamp As String
Private smCurrAFEStamp As String
Private smCurrAPEStamp As String
Private imMaxChar(0 To 37) As Integer
Private imMaxCols As Integer
Private lmEEnableRow As Long
Private lmEEnableCol As Long
Private lmIEnableRow As Long
Private lmIEnableCol As Long


Private tmAEE As AEE
Private tmCurrACE() As ACE
Private tmCurrADE() As ADE
Private tmCurrAFE() As AFE
Private tmCurrAPE() As APE

Private imTabIndex As Integer

Const BUSNAMEINDEX = 0
Const BUSCTRLINDEX = 1
Const EVENTTYPEINDEX = 2
Const TIMEINDEX = 3
Const STARTTYPEINDEX = 4
Const FIXEDINDEX = 5
Const ENDTYPEINDEX = 6
Const DURATIONINDEX = 7
Const ENDTIMEINDEX = 8
Const MATERIALINDEX = 9
Const AUDIONAMEINDEX = 10
Const AUDIOITEMIDINDEX = 11
Const AUDIOISCIINDEX = 12
Const AUDIOCTRLINDEX = 13
Const BACKUPNAMEINDEX = 14  '16
Const BACKUPCTRLINDEX = 15  '17
Const PROTNAMEINDEX = 16    '13
Const PROTITEMIDINDEX = 17  '14
Const PROTISCIINDEX = 18  '14
Const PROTCTRLINDEX = 19    '15
Const RELAY1INDEX = 20
Const RELAY2INDEX = 21
Const FOLLOWINDEX = 22
Const SILENCETIMEINDEX = 23
Const SILENCE1INDEX = 24
Const SILENCE2INDEX = 25
Const SILENCE3INDEX = 26
Const SILENCE4INDEX = 27
Const NETCUE1INDEX = 28
Const NETCUE2INDEX = 29
Const TITLE1INDEX = 30
Const TITLE2INDEX = 31
Const DATEINDEX = 32
Const EVENTIDINDEX = 33
Const ABCFORMATINDEX = 34
Const ABCPGMCODEINDEX = 35
Const ABCXDSMODEINDEX = 36
Const ABCRECORDITEMINDEX = 37
Const CODEINDEX = 38

Const ECHOSTARTINDEX = 0
Const DATESTARTINDEX = 1
Const DATENOCHARINDEX = 2
Const TIMESTARTINDEX = 3
Const TIMENOCHARINDEX = 4
Const AUTOOFFINDEX = 5
Const DATAINDEX = 6
Const SCHEDULEINDEX = 7
Const TRUETIMEINDEX = 8
Const SRCECONFLICTINDEX = 9
Const SRCEUNAVAILINDEX = 10
Const SRCEITEMINDEX = 11
Const BKUPUNAVAILINDEX = 14
Const BKUPITEMINDEX = 15
Const PROTUNAVAILINDEX = 12
Const PROTITEMINDEX = 13
Const ADECODEINDEX = 16

Private Sub mClearControls()
    Dim ilCol As Integer
    imVersion = -1
    smUsedFlag = "N"
    gClearControls EngrAutomation
    For ilCol = 0 To grdExport.Cols - 1 Step 1
        grdExport.TextMatrix(3, ilCol) = ""
        grdExport.TextMatrix(4, ilCol) = ""
        'grdExport.TextMatrix(6, ilCol) = ""
        'grdExport.TextMatrix(7, ilCol) = ""
    Next ilCol
    ReDim tmCurrACE(0 To 0) As ACE
    ReDim tmCurrADE(0 To 0) As ADE
    ReDim tmCurrAFE(0 To 0) As AFE
    ReDim tmCurrAPE(0 To 0) As APE
    imFieldChgd = False
End Sub

Private Function mCompare(tlNew As AEE, tlOld As AEE) As Integer
    If StrComp(tlNew.sName, tlOld.sName, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sDescription, tlOld.sDescription, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sManufacture, tlOld.sManufacture, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sFixedTimeChar, tlOld.sFixedTimeChar, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If tlNew.lAlertSchdDelay <> tlOld.lAlertSchdDelay Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sState, tlOld.sState, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    mCompare = True
End Function

Private Function mCompareACE(ilCode As Integer) As Integer
    Dim ilACENew As Integer
    Dim ilACEOld As Integer
    
    If ilCode > 0 Then
        For ilACENew = LBound(tmCurrACE) To UBound(tmCurrACE) - 1 Step 1
            If ilCode = tmCurrACE(ilACENew).iCode Then
                For ilACEOld = LBound(tgCurrACE) To UBound(tgCurrACE) - 1 Step 1
                    If ilCode = tgCurrACE(ilACEOld).iCode Then
                        'Compare fields
                        If tmCurrACE(ilACENew).sType <> tgCurrACE(ilACEOld).sType Then
                            mCompareACE = False
                            Exit Function
                        End If
                        If tmCurrACE(ilACENew).sContact <> tgCurrACE(ilACEOld).sContact Then
                            mCompareACE = False
                            Exit Function
                        End If
                        If tmCurrACE(ilACENew).sPhone <> tgCurrACE(ilACEOld).sPhone Then
                            mCompareACE = False
                            Exit Function
                        End If
                        If tmCurrACE(ilACENew).sFax <> tgCurrACE(ilACEOld).sFax Then
                            mCompareACE = False
                            Exit Function
                        End If
                        If tmCurrACE(ilACENew).sEMail <> tgCurrACE(ilACEOld).sEMail Then
                            mCompareACE = False
                            Exit Function
                        End If
                        mCompareACE = True
                        Exit Function
                    End If
                Next ilACEOld
                mCompareACE = True
                Exit Function
            End If
        Next ilACENew
    Else
        mCompareACE = True
    End If
End Function


Private Function mCompareAPE(ilCode As Integer) As Integer
    Dim ilAPENew As Integer
    Dim ilAPEOld As Integer
    
    If ilCode > 0 Then
        For ilAPENew = LBound(tmCurrAPE) To UBound(tmCurrAPE) - 1 Step 1
            If ilCode = tmCurrAPE(ilAPENew).iCode Then
                For ilAPEOld = LBound(tgCurrAPE) To UBound(tgCurrAPE) - 1 Step 1
                    If ilCode = tgCurrAPE(ilAPEOld).iCode Then
                        'Compare fields
                        If tmCurrAPE(ilAPENew).sType <> tgCurrAPE(ilAPEOld).sType Then
                            mCompareAPE = False
                            Exit Function
                        End If
                        If (tmCurrAPE(ilAPENew).sType = "SE") Or (tmCurrAPE(ilAPENew).sType = "SI") Then
                            If tmCurrAPE(ilAPENew).sNewFileName <> tgCurrAPE(ilAPEOld).sNewFileName Then
                                mCompareAPE = False
                                Exit Function
                            End If
                            If tmCurrAPE(ilAPENew).sNewFileExt <> tgCurrAPE(ilAPEOld).sNewFileExt Then
                                mCompareAPE = False
                                Exit Function
                            End If
                        End If
                        If (tmCurrAPE(ilAPENew).sType = "SE") Then
                            If tmCurrAPE(ilAPENew).sChgFileExt <> tgCurrAPE(ilAPEOld).sChgFileExt Then
                                mCompareAPE = False
                                Exit Function
                            End If
                            If tmCurrAPE(ilAPENew).sDelFileExt <> tgCurrAPE(ilAPEOld).sDelFileExt Then
                                mCompareAPE = False
                                Exit Function
                            End If
                        End If
                        If tmCurrAPE(ilAPENew).sPath <> tgCurrAPE(ilAPEOld).sPath Then
                            mCompareAPE = False
                            Exit Function
                        End If
                        mCompareAPE = True
                        Exit Function
                    End If
                Next ilAPEOld
                mCompareAPE = True
                Exit Function
            End If
        Next ilAPENew
    Else
        mCompareAPE = True
    End If
    
    
    
End Function

Private Sub cbcSelect_Change()
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilLen As Integer
    Dim ilSel As Integer
    Dim llRow As Long
    
    If imInChg Then
        Exit Sub
    End If
    imInChg = True
    Screen.MousePointer = vbHourglass
    slName = LTrim$(cbcSelect.text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cbcSelect.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        cbcSelect.ListIndex = llRow
        cbcSelect.SelStart = ilLen
        cbcSelect.SelLength = Len(cbcSelect.text)
        imAeeCode = cbcSelect.ItemData(cbcSelect.ListIndex)
        If imAeeCode <= 0 Then
            mClearControls
            rbcState(0).Value = True
            mMoveAFERecToCtrls
            mMoveADERecToCtrls
            ReDim tmCurrACE(0 To 2) As ACE
            ReDim tmCurrAPE(0 To 4) As APE
        Else
            'Load existing data
            mClearControls
            For ilLoop = LBound(tgCurrAEE) To UBound(tgCurrAEE) - 1 Step 1
                If imAeeCode = tgCurrAEE(ilLoop).iCode Then
                    mMoveRecToCtrls ilLoop
                    Exit For
                End If
            Next ilLoop
        End If
    End If
    imInChg = False
    imFieldChgd = False
    mSetCommands
    Screen.MousePointer = vbDefault
    Exit Sub

End Sub

Private Sub cbcSelect_Click()
    cbcSelect_Change
End Sub

Private Sub cbcSelect_GotFocus()
    mISetShow
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmIEnableRow = -1
    lmIEnableCol = -1
End Sub

Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cbcSelect.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmcCancel_Click()
    Unload EngrAutomation
End Sub

Private Sub cmcCancel_GotFocus()
    mISetShow
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmIEnableRow = -1
    lmIEnableCol = -1
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        Unload EngrAutomation
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        ilRet = mSave()
        If Not ilRet Then
            Exit Sub
        End If
        sgCurrAEEStamp = ""
        sgCurrACEStamp = ""
        sgCurrADEStamp = ""
        sgCurrAFEStamp = ""
        sgCurrAPEStamp = ""
        ilRet = gGetTypeOfRecs_AEE_AutoEquip("C", sgCurrAEEStamp, "EngrAutomation-mPopulate", tgCurrAEE())
        ilRet = gGetTypeOfRecs_ACE_AutoContact("C", sgCurrACEStamp, "EngrEventType-mPopulate", tgCurrACE())
        ilRet = gGetTypeOfRecs_ADE_AutoDataFlags("C", sgCurrADEStamp, "EngrEventType-mPopulate", tgCurrADE())
        ilRet = gGetTypeOfRecs_AFE_AutoFormat("C", sgCurrAFEStamp, "EngrEventType-mPopulate", tgCurrAFE())
        ilRet = gGetTypeOfRecs_APE_AutoPath("C", sgCurrAPEStamp, "EngrEventType-mPopulate", tgCurrAPE())
    End If
    
    Screen.MousePointer = vbDefault
    Unload EngrAutomation
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mISetShow
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmIEnableRow = -1
    lmIEnableCol = -1
End Sub

Private Sub cmcErase_Click()
    Dim slStr As String
    Dim slMsg As String
    Dim ilRet As Integer
    
    If cmcDone.Enabled = False Then
        Exit Sub
    End If
    slStr = Trim$(edcName.text)
    If smUsedFlag <> "N" Then
        MsgBox slStr & " used or was used, unable to delete", vbInformation + vbOKCancel, "Erase"
        Exit Sub
    End If
    slMsg = "Delete " & slStr
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilRet = gPutDelete_AEE_AutoEquip(imAeeCode, "EngrAutomation- Delete")
    sgCurrAEEStamp = ""
    sgCurrACEStamp = ""
    sgCurrADEStamp = ""
    sgCurrAFEStamp = ""
    sgCurrAPEStamp = ""
    mPopulate
    If cbcSelect.ListCount = 2 Then
        cbcSelect.ListIndex = 1
    ElseIf cbcSelect.ListCount >= 1 Then
        cbcSelect.ListIndex = 0
    End If
    imFieldChgd = False
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcErase_GotFocus()
    mISetShow
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmIEnableRow = -1
    lmIEnableCol = -1
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    
    If imFieldChgd = True Then
        slName = Trim$(edcName.text)
        ilRet = mSave()
        If Not ilRet Then
            Exit Sub
        End If
        sgCurrAEEStamp = ""
        sgCurrACEStamp = ""
        sgCurrADEStamp = ""
        sgCurrAFEStamp = ""
        sgCurrAPEStamp = ""
        mPopulate
        cbcSelect.text = slName
        imFieldChgd = False
        mSetCommands
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mISetShow
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmIEnableRow = -1
    lmIEnableCol = -1
End Sub

Private Sub edcClientExportPath_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcClientExportPath_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcClientExportPathTest_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcClientExportPathTest_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcClientImportPath_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcClientImportPath_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDateFormat_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDateFormat_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDelay_Change()
    Dim slStr As String
    slStr = edcDelay.text
    If gIsLength(slStr) Then
        imFieldChgd = True
        mSetCommands
    End If
End Sub

Private Sub edcDelay_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDescription_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDescription_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcExport_Change()
    If grdExport.text <> edcExport.text Then
        imFieldChgd = True
    End If
    grdExport.text = edcExport.text
    grdExport.CellForeColor = vbBlack
    mSetCommands
End Sub

Private Sub edcExport_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcExport_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    If lmEEnableRow = 4 Then
        slStr = edcExport.text
        slStr = Left$(slStr, edcExport.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcExport.SelStart - edcExport.SelLength)
        If Val(slStr) > imMaxChar(lmEEnableCol) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcExportChgFileFormat_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcExportChgFileFormat_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcExportDelFileFormat_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcExportDelFileFormat_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcExportExtChg_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcExportExtChg_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcExportExtDel_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcExportExtDel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcExportExtNew_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcExportExtNew_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcExportNewFileFormat_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcExportNewFileFormat_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFixedTimeChar_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcFixedTimeChar_GotFocus()
    mISetShow
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmIEnableRow = -1
    lmIEnableCol = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcImport_Change()
    If grdImport.text <> edcImport.text Then
        imFieldChgd = True
    End If
    grdImport.text = edcImport.text
    grdImport.CellForeColor = vbBlack
    mSetCommands
End Sub

Private Sub edcImport_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcImport_KeyPress(KeyAscii As Integer)
    
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If

End Sub

Private Sub edcImportExt_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcImportExt_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcImportFileFormat_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcImportFileFormat_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcManufacture_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcManufacture_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPriContactName_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPriContactName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPriEMail_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPriEMail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPriFax_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPriFax_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPriPhone_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcPriPhone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSecContactName_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcSecContactName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSecEMail_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcSecEMail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSecFax_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcSecFax_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSecPhone_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcSecPhone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcServerExportPath_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcServerExportPath_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcServerExportPathTest_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcServerExportPathTest_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcServerImportPath_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcServerImportPath_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcTimeFormat_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcTimeFormat_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    mGridColumns
    lacSchStartCol.Top = grdExport.Top + 3 * grdExport.RowHeight(0)
    lacSchNoChars.Top = grdExport.Top + 4 * grdExport.RowHeight(0)
'    lacAsAirStartCol.Top = grdExport.Top + 6 * grdExport.RowHeight(0)
'    lacAsAirNoChars.Top = grdExport.Top + 7 * grdExport.RowHeight(0)
    mSetTab
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrAutomation
    gCenterFormModal EngrAutomation
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    mInit
    imTabIndex = 1
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Bus Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Bus Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub Form_Resize()
    mSetTab
    mGridColumnWidth
    grdExport.Height = 8 * grdExport.RowHeight(0) + 15
    gGrid_IntegralHeight grdExport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EngrAutomation = Nothing
End Sub

Private Sub frcTab_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 2 Then
        mISetShow
        mESetShow
        lmEEnableRow = -1
        lmEEnableCol = -1
        lmIEnableRow = -1
        lmIEnableCol = -1
    End If
End Sub

Private Sub grdExport_EnterCell()
    mISetShow
    mESetShow
End Sub

Private Sub grdExport_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdExport.RowHeight(0) Then
        'mSortCol grdExport.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    'ilFound = gGrid_DetermineRowCol(grdExport, X, Y)
    'If Not ilFound Then
    '    pbcClickFocus.SetFocus
    '    Exit Sub
    'End If
    If grdExport.Col >= grdExport.Cols - 1 Then
        Exit Sub
    End If
    mEEnableBox
End Sub

Private Sub grdExport_Scroll()
    mISetShow
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmIEnableRow = -1
    lmIEnableCol = -1
End Sub

Private Sub grdImport_EnterCell()
    mISetShow
    mESetShow
End Sub

Private Sub grdImport_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdImport.RowHeight(0) Then
        'mSortCol grdExport.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdImport, x, y)
    If Not ilFound Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdImport.Col >= grdImport.Cols - 1 Then
        Exit Sub
    End If
    mIEnableBox
End Sub

Private Sub pbcClickFocus_GotFocus()
    mISetShow
    mESetShow
    lmIEnableRow = -1
    lmIEnableCol = -1
    lmEEnableRow = -1
    lmEEnableCol = -1
End Sub

Private Sub pbcExportSTab_GotFocus()
    If GetFocus() <> pbcExportSTab.hwnd Then
        Exit Sub
    End If
    mISetShow
    If edcExport.Visible Then
        mESetShow
        If grdExport.Col = BUSNAMEINDEX Then
            If grdExport.Row > grdExport.FixedRows Then
                grdExport.Row = grdExport.Row - 1
                grdExport.Col = imMaxCols
                mEEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdExport.Col = grdExport.Col - 1
            mEEnableBox
        End If
    Else
        grdExport.Col = BUSNAMEINDEX
        grdExport.Row = grdExport.FixedRows + 1
        mEEnableBox
    End If
End Sub

Private Sub pbcExportTab_GotFocus()
    Dim llRow As Long
    
    If GetFocus() <> pbcExportTab.hwnd Then
        Exit Sub
    End If
    mISetShow
    If edcExport.Visible Then
        mESetShow
        If grdExport.Col = imMaxCols Then
            If grdExport.Row >= grdExport.Rows - 1 Then
                pbcImportSTab.SetFocus
            Else
                grdExport.Row = grdExport.Row + 1
                grdExport.LeftCol = BUSNAMEINDEX
                grdExport.Col = BUSNAMEINDEX
                mEEnableBox
            End If
        Else
            grdExport.Col = grdExport.Col + 1
            mEEnableBox
        End If
    Else
        grdExport.Col = imMaxCols    'BUSNAMEINDEX
        grdExport.Row = grdExport.FixedRows + 1
        mEEnableBox
    End If
End Sub

Private Sub pbcImportSTab_GotFocus()
    If GetFocus() <> pbcImportSTab.hwnd Then
        Exit Sub
    End If
    mESetShow
    If edcImport.Visible Then
        mISetShow
        If grdExport.Col = BUSNAMEINDEX Then
            pbcExportTab.SetFocus
        Else
            grdExport.Col = grdExport.Col - 1
            mEEnableBox
        End If
    Else
        grdImport.Col = ECHOSTARTINDEX
        grdImport.Row = grdImport.FixedRows + 1
        mIEnableBox
    End If
End Sub

Private Sub pbcImportTab_GotFocus()
    Dim llRow As Long
    
    If GetFocus() <> pbcImportTab.hwnd Then
        Exit Sub
    End If
    mESetShow
    If edcImport.Visible Then
        mISetShow
        If grdImport.Col = imMaxCols Then
            cmcCancel.SetFocus
        Else
            grdImport.Col = grdImport.Col + 1
            mIEnableBox
        End If
    Else
        grdImport.Col = BUSNAMEINDEX
        grdImport.Row = grdImport.FixedRows + 1
        mIEnableBox
    End If
End Sub

Private Sub rbcState_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub edcName_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub




Private Sub mGridColumns()
    Dim ilCol As Integer
    
    gGrid_AlignAllColsLeft grdExport
    gGrid_AlignAllColsLeft grdImport
    mGridColumnWidth
    
    'Set Titles
    For ilCol = BUSNAMEINDEX To BUSCTRLINDEX Step 1
        grdExport.TextMatrix(0, ilCol) = "Bus"
    Next ilCol
'    For ilCol = TIMEINDEX To ENDTIMEINDEX Step 1
'        grdExport.TextMatrix(0, ilCol) = "Time"
'    Next ilCol
    For ilCol = AUDIONAMEINDEX To AUDIOCTRLINDEX Step 1
        grdExport.TextMatrix(0, ilCol) = "Audio"
    Next ilCol
    For ilCol = BACKUPNAMEINDEX To BACKUPCTRLINDEX Step 1
        grdExport.TextMatrix(0, ilCol) = "Backup"
    Next ilCol
    For ilCol = PROTNAMEINDEX To PROTCTRLINDEX Step 1
        grdExport.TextMatrix(0, ilCol) = "Protection"
    Next ilCol
    For ilCol = RELAY1INDEX To RELAY2INDEX Step 1
        grdExport.TextMatrix(0, ilCol) = "Relay"
    Next ilCol
    For ilCol = SILENCETIMEINDEX To SILENCE4INDEX Step 1
        grdExport.TextMatrix(0, ilCol) = "Silence"
    Next ilCol
    For ilCol = NETCUE1INDEX To NETCUE2INDEX Step 1
        grdExport.TextMatrix(0, ilCol) = "Netcue"
    Next ilCol
    For ilCol = TITLE1INDEX To TITLE2INDEX Step 1
        grdExport.TextMatrix(0, ilCol) = "Title"
    Next ilCol
    grdExport.TextMatrix(1, BUSNAMEINDEX) = "Name"
    grdExport.TextMatrix(1, BUSCTRLINDEX) = "C"
    grdExport.TextMatrix(0, EVENTTYPEINDEX) = "Evt"
    grdExport.TextMatrix(1, EVENTTYPEINDEX) = "Type"
    grdExport.TextMatrix(0, TIMEINDEX) = "Start"
    grdExport.TextMatrix(1, TIMEINDEX) = "Time"
    grdExport.TextMatrix(0, STARTTYPEINDEX) = "Start "
    grdExport.TextMatrix(1, STARTTYPEINDEX) = "Type"
    grdExport.TextMatrix(0, FIXEDINDEX) = "Fix"
    grdExport.TextMatrix(0, ENDTYPEINDEX) = "End"
    grdExport.TextMatrix(1, ENDTYPEINDEX) = "Type"
    grdExport.TextMatrix(0, DURATIONINDEX) = "Dur"
    grdExport.TextMatrix(0, ENDTIMEINDEX) = "Out"
    grdExport.TextMatrix(1, ENDTIMEINDEX) = "Time"
    grdExport.TextMatrix(0, MATERIALINDEX) = "Mat"
    grdExport.TextMatrix(1, MATERIALINDEX) = "Type"
    grdExport.TextMatrix(1, AUDIONAMEINDEX) = "Name"
    grdExport.TextMatrix(1, AUDIOITEMIDINDEX) = "Item"
    grdExport.TextMatrix(1, AUDIOISCIINDEX) = "ISCI"
    grdExport.TextMatrix(1, AUDIOCTRLINDEX) = "C"
    grdExport.TextMatrix(1, BACKUPNAMEINDEX) = "Name"
    grdExport.TextMatrix(1, BACKUPCTRLINDEX) = "C"
    grdExport.TextMatrix(1, PROTNAMEINDEX) = "Name"
    grdExport.TextMatrix(1, PROTITEMIDINDEX) = "Item"
    grdExport.TextMatrix(1, PROTISCIINDEX) = "ISCI"
    grdExport.TextMatrix(1, PROTCTRLINDEX) = "C"
    grdExport.TextMatrix(1, RELAY1INDEX) = "1"
    grdExport.TextMatrix(1, RELAY2INDEX) = "2"
    grdExport.TextMatrix(0, FOLLOWINDEX) = "Fol-"
    grdExport.TextMatrix(1, FOLLOWINDEX) = "low"
    grdExport.TextMatrix(1, SILENCETIMEINDEX) = "Time"
    grdExport.TextMatrix(1, SILENCE1INDEX) = "1"
    grdExport.TextMatrix(1, SILENCE2INDEX) = "2"
    grdExport.TextMatrix(1, SILENCE3INDEX) = "3"
    grdExport.TextMatrix(1, SILENCE4INDEX) = "4"
    grdExport.TextMatrix(1, NETCUE1INDEX) = "Start"
    grdExport.TextMatrix(1, NETCUE2INDEX) = "Stop"
    grdExport.TextMatrix(1, TITLE1INDEX) = "1"
    grdExport.TextMatrix(1, TITLE2INDEX) = "2"
    grdExport.TextMatrix(0, DATEINDEX) = "Date"
    grdExport.TextMatrix(0, EVENTIDINDEX) = "Evt"
    grdExport.TextMatrix(1, EVENTIDINDEX) = "ID"
    grdExport.TextMatrix(0, ABCFORMATINDEX) = "For-"
    grdExport.TextMatrix(1, ABCFORMATINDEX) = "mat"
    grdExport.TextMatrix(0, ABCPGMCODEINDEX) = "Pgm"
    grdExport.TextMatrix(1, ABCPGMCODEINDEX) = "Code"
    grdExport.TextMatrix(0, ABCXDSMODEINDEX) = "XDS"
    grdExport.TextMatrix(1, ABCXDSMODEINDEX) = "Mode"
    grdExport.TextMatrix(0, ABCRECORDITEMINDEX) = "Rec'd"
    grdExport.TextMatrix(1, ABCRECORDITEMINDEX) = "Item"
    
    grdImport.TextMatrix(0, ECHOSTARTINDEX) = "Echo Schd"
    grdImport.TextMatrix(1, ECHOSTARTINDEX) = "Start Col"
    For ilCol = DATESTARTINDEX To DATENOCHARINDEX Step 1
        grdImport.TextMatrix(0, ilCol) = "Date"
    Next ilCol
    grdImport.TextMatrix(1, DATESTARTINDEX) = "Start Col"
    grdImport.TextMatrix(1, DATENOCHARINDEX) = "# Chars"
    For ilCol = TIMESTARTINDEX To TIMENOCHARINDEX Step 1
        grdImport.TextMatrix(0, ilCol) = "Time"
    Next ilCol
    grdImport.TextMatrix(1, TIMESTARTINDEX) = "Start Col"
    grdImport.TextMatrix(1, TIMENOCHARINDEX) = "# Chars"
    grdImport.TextMatrix(0, AUTOOFFINDEX) = "Auto-Off"
    grdImport.TextMatrix(1, AUTOOFFINDEX) = "Error"
    grdImport.TextMatrix(0, DATAINDEX) = "Data"
    grdImport.TextMatrix(1, DATAINDEX) = "Error"
    grdImport.TextMatrix(0, SCHEDULEINDEX) = "Schd"
    grdImport.TextMatrix(1, SCHEDULEINDEX) = "Error"
    grdImport.TextMatrix(0, TRUETIMEINDEX) = "True Time"
    grdImport.TextMatrix(1, TRUETIMEINDEX) = "Error"
    For ilCol = SRCECONFLICTINDEX To SRCEITEMINDEX Step 1
        grdImport.TextMatrix(0, ilCol) = "Source"
    Next ilCol
    grdImport.TextMatrix(1, SRCECONFLICTINDEX) = "Conflict"
    grdImport.TextMatrix(1, SRCEUNAVAILINDEX) = "Not Avail."
    grdImport.TextMatrix(1, SRCEITEMINDEX) = "Item"
    For ilCol = BKUPUNAVAILINDEX To BKUPITEMINDEX Step 1
        grdImport.TextMatrix(0, ilCol) = "Backup"
    Next ilCol
    grdImport.TextMatrix(1, BKUPUNAVAILINDEX) = "Not Avail."
    grdImport.TextMatrix(1, BKUPITEMINDEX) = "Item"
    For ilCol = PROTUNAVAILINDEX To PROTITEMINDEX Step 1
        grdImport.TextMatrix(0, ilCol) = "Protection"
    Next ilCol
    grdImport.TextMatrix(1, PROTUNAVAILINDEX) = "Not Avail."
    grdImport.TextMatrix(1, PROTITEMINDEX) = "Item"
    
    
    grdExport.Row = 1
    For ilCol = 0 To grdExport.Cols - 1 Step 1
        grdExport.Col = ilCol
        grdExport.CellAlignment = flexAlignLeftCenter
    Next ilCol
    'grdExport.Row = 0
    grdExport.MergeCells = flexMergeRestrictRows
    grdExport.MergeRow(0) = True
    grdExport.Row = 0
    grdExport.Col = BUSNAMEINDEX
    grdExport.CellAlignment = flexAlignCenterCenter
    grdExport.Row = 0
    grdExport.Col = AUDIONAMEINDEX
    grdExport.CellAlignment = flexAlignCenterCenter
    grdExport.Row = 0
    grdExport.Col = BACKUPNAMEINDEX
    grdExport.CellAlignment = flexAlignCenterCenter
    grdExport.Row = 0
    grdExport.Col = PROTNAMEINDEX
    grdExport.CellAlignment = flexAlignCenterCenter
    grdExport.Row = 0
    grdExport.Col = RELAY1INDEX
    grdExport.CellAlignment = flexAlignCenterCenter
    grdExport.Row = 0
    grdExport.Col = SILENCETIMEINDEX
    grdExport.CellAlignment = flexAlignCenterCenter
    grdExport.Row = 0
    grdExport.Col = NETCUE1INDEX
    grdExport.CellAlignment = flexAlignCenterCenter
    grdExport.Row = 0
    grdExport.Col = TITLE1INDEX
    grdExport.CellAlignment = flexAlignCenterCenter
    grdExport.Height = 8 * grdExport.RowHeight(0) + 30
    gGrid_IntegralHeight grdExport
    'gGrid_Clear grdExport, True
    'Create Titles
    For ilCol = 0 To grdExport.Cols - 1 Step 1
        grdExport.TextMatrix(2, ilCol) = "Schedule"
'        grdExport.TextMatrix(5, ilCol) = "As Aired"
        grdExport.Row = 2
        grdExport.Col = ilCol
        grdExport.CellBackColor = LIGHTYELLOW
'        grdExport.Row = 5
'        grdExport.Col = ilCol
'        grdExport.CellBackColor = LIGHTYELLOW
    Next ilCol
    'grdExport.Row = 2
    grdExport.Col = 0
    grdExport.MergeCells = flexMergeRestrictRows
    grdExport.MergeRow(2) = True
    'grdExport.Row = 5
'    grdExport.MergeCells = flexMergeRestrictRows
'    grdExport.MergeRow(5) = True
    'grdExport.Row = grdExport.FixedRows + 1
'    grdExport.Height = 8 * grdExport.RowHeight(0) + 15
    grdExport.Height = 5 * grdExport.RowHeight(0) + 15
    gGrid_IntegralHeight grdExport
    
    grdImport.Row = 1
    For ilCol = 0 To grdImport.Cols - 1 Step 1
        grdImport.Col = ilCol
        grdImport.CellAlignment = flexAlignLeftCenter
    Next ilCol
    For ilCol = 0 To grdImport.Cols - 1 Step 1
        grdImport.TextMatrix(2, ilCol) = "As Aired"
        grdImport.Row = 2
        grdImport.Col = ilCol
        grdImport.CellBackColor = LIGHTYELLOW
    Next ilCol
    grdImport.MergeCells = flexMergeRestrictRows
    grdImport.MergeRow(0) = True
    grdImport.Row = 0
    grdImport.Col = DATESTARTINDEX
    grdImport.CellAlignment = flexAlignCenterCenter
    grdImport.Row = 0
    grdImport.Col = TIMESTARTINDEX
    grdImport.CellAlignment = flexAlignCenterCenter
    grdImport.Row = 0
    grdImport.Col = SRCECONFLICTINDEX
    grdImport.CellAlignment = flexAlignCenterCenter
    grdImport.Row = 0
    grdImport.Col = BKUPUNAVAILINDEX
    grdImport.CellAlignment = flexAlignCenterCenter
    grdImport.Row = 0
    grdImport.Col = PROTUNAVAILINDEX
    grdImport.CellAlignment = flexAlignCenterCenter
    
    grdImport.Col = 0
    grdImport.MergeCells = flexMergeRestrictRows
    grdImport.MergeRow(2) = True
    
    grdImport.Height = 4 * grdImport.RowHeight(0) + 15
    gGrid_IntegralHeight grdImport
    
    
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdExport.ColWidth(CODEINDEX) = 0
    grdExport.ColWidth(BUSNAMEINDEX) = grdExport.Width / 28
    grdExport.ColWidth(BUSCTRLINDEX) = grdExport.Width / 40
    grdExport.ColWidth(EVENTTYPEINDEX) = grdExport.Width / 33
    grdExport.ColWidth(STARTTYPEINDEX) = grdExport.Width / 33
    grdExport.ColWidth(FIXEDINDEX) = grdExport.Width / 40
    grdExport.ColWidth(ENDTYPEINDEX) = grdExport.Width / 33
    grdExport.ColWidth(DURATIONINDEX) = grdExport.Width / 35
    grdExport.ColWidth(ENDTIMEINDEX) = grdExport.Width / 36
    grdExport.ColWidth(MATERIALINDEX) = grdExport.Width / 33
    grdExport.ColWidth(AUDIONAMEINDEX) = grdExport.Width / 28
    grdExport.ColWidth(AUDIOITEMIDINDEX) = grdExport.Width / 30
    grdExport.ColWidth(AUDIOISCIINDEX) = grdExport.Width / 30
    grdExport.ColWidth(AUDIOCTRLINDEX) = grdExport.Width / 42
    grdExport.ColWidth(BACKUPNAMEINDEX) = grdExport.Width / 28
    grdExport.ColWidth(BACKUPCTRLINDEX) = grdExport.Width / 42
    grdExport.ColWidth(PROTNAMEINDEX) = grdExport.Width / 28
    grdExport.ColWidth(PROTITEMIDINDEX) = grdExport.Width / 30
    grdExport.ColWidth(PROTISCIINDEX) = grdExport.Width / 30
    grdExport.ColWidth(PROTCTRLINDEX) = grdExport.Width / 42
    grdExport.ColWidth(RELAY1INDEX) = grdExport.Width / 37
    grdExport.ColWidth(RELAY2INDEX) = grdExport.Width / 37
    grdExport.ColWidth(FOLLOWINDEX) = grdExport.Width / 35
    grdExport.ColWidth(SILENCETIMEINDEX) = grdExport.Width / 29
    grdExport.ColWidth(SILENCE1INDEX) = grdExport.Width / 40
    grdExport.ColWidth(SILENCE2INDEX) = grdExport.Width / 40
    grdExport.ColWidth(SILENCE3INDEX) = grdExport.Width / 40
    grdExport.ColWidth(SILENCE4INDEX) = grdExport.Width / 40
    grdExport.ColWidth(NETCUE1INDEX) = grdExport.Width / 30
    grdExport.ColWidth(NETCUE2INDEX) = grdExport.Width / 30
    grdExport.ColWidth(TITLE1INDEX) = grdExport.Width / 40
    grdExport.ColWidth(TITLE2INDEX) = grdExport.Width / 40
    grdExport.ColWidth(DATEINDEX) = grdExport.Width / 32
    grdExport.ColWidth(EVENTIDINDEX) = grdExport.Width / 38
    If sgClientFields = "A" Then
        grdExport.ColWidth(ABCFORMATINDEX) = grdExport.Width / 28
        grdExport.ColWidth(ABCPGMCODEINDEX) = grdExport.Width / 28
        grdExport.ColWidth(ABCXDSMODEINDEX) = grdExport.Width / 28
        grdExport.ColWidth(ABCRECORDITEMINDEX) = grdExport.Width / 28
    Else
        grdExport.ColWidth(ABCFORMATINDEX) = 0
        grdExport.ColWidth(ABCPGMCODEINDEX) = 0
        grdExport.ColWidth(ABCXDSMODEINDEX) = 0
        grdExport.ColWidth(ABCRECORDITEMINDEX) = 0
    End If
    grdExport.ColWidth(TIMEINDEX) = grdExport.Width '- GRIDSCROLLWIDTH
    For ilCol = BUSNAMEINDEX To EVENTIDINDEX Step 1
        If ilCol <> TIMEINDEX Then
            If grdExport.ColWidth(TIMEINDEX) > grdExport.ColWidth(ilCol) Then
                grdExport.ColWidth(TIMEINDEX) = grdExport.ColWidth(TIMEINDEX) - grdExport.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol

    grdImport.ColWidth(ADECODEINDEX) = 0
    grdImport.ColWidth(ECHOSTARTINDEX) = grdImport.Width / 13
    grdImport.ColWidth(DATESTARTINDEX) = grdImport.Width / 16
    grdImport.ColWidth(DATENOCHARINDEX) = grdImport.Width / 16
    'grdImport.ColWidth(TIMESTARTINDEX) = grdImport.Width / 15
    grdImport.ColWidth(TIMENOCHARINDEX) = grdImport.Width / 16
    grdImport.ColWidth(AUTOOFFINDEX) = grdImport.Width / 16
    grdImport.ColWidth(DATAINDEX) = grdImport.Width / 20
    grdImport.ColWidth(SCHEDULEINDEX) = grdImport.Width / 20
    grdImport.ColWidth(TRUETIMEINDEX) = grdImport.Width / 15
    grdImport.ColWidth(SRCECONFLICTINDEX) = grdImport.Width / 16
    grdImport.ColWidth(SRCEUNAVAILINDEX) = grdImport.Width / 16
    grdImport.ColWidth(SRCEITEMINDEX) = grdImport.Width / 16
    grdImport.ColWidth(BKUPUNAVAILINDEX) = grdImport.Width / 16
    grdImport.ColWidth(BKUPITEMINDEX) = grdImport.Width / 16
    grdImport.ColWidth(PROTUNAVAILINDEX) = grdImport.Width / 16
    grdImport.ColWidth(PROTITEMINDEX) = grdImport.Width / 16
    grdImport.ColWidth(TIMESTARTINDEX) = grdImport.Width '- GRIDSCROLLWIDTH
    For ilCol = ECHOSTARTINDEX To PROTITEMINDEX Step 1
        If ilCol <> TIMESTARTINDEX Then
            If grdImport.ColWidth(TIMESTARTINDEX) > grdImport.ColWidth(ilCol) Then
                grdImport.ColWidth(TIMESTARTINDEX) = grdImport.ColWidth(TIMESTARTINDEX) - grdImport.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol

End Sub

Private Sub tabAuto_Click()
    If imTabIndex = tabAuto.SelectedItem.Index Then
        Exit Sub
    End If
    mISetShow
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    lmIEnableRow = -1
    lmIEnableCol = -1
    frcTab(tabAuto.SelectedItem.Index - 1).Visible = True
    frcTab(imTabIndex - 1).Visible = False
    imTabIndex = tabAuto.SelectedItem.Index
End Sub

Private Sub mSetTab()
    tabAuto.Left = frcSelect.Left
    tabAuto.Height = cmcCancel.Top - (frcSelect.Top + frcSelect.Height + 300)  'TabAuto.ClientTop - TabAuto.Top + (10 * frcTab(0).Height) / 9
    frcTab(0).Move tabAuto.ClientLeft, tabAuto.ClientTop, tabAuto.ClientWidth, tabAuto.ClientHeight
    frcTab(1).Move tabAuto.ClientLeft, tabAuto.ClientTop, tabAuto.ClientWidth, tabAuto.ClientHeight
    frcTab(2).Move tabAuto.ClientLeft, tabAuto.ClientTop, tabAuto.ClientWidth, tabAuto.ClientHeight
    frcTab(0).BorderStyle = 0
    frcTab(1).BorderStyle = 0
    frcTab(2).BorderStyle = 0
End Sub

Private Sub mInit()
    mPopulate
    mSetMaxChar
'    If ((StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0) And (Len(sgSpecialPassword) = 5)) Or _
'       (((StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0) And (Len(sgSpecialPassword) = 4))) Then
'    Else
'        'tabAuto.Tabs.Remove 2
'        grdExport.Visible = False
'        grdImport.Visible = False
'        lacSchStartCol.Visible = False
'        lacSchNoChars.Visible = False
''        lacAsAirStartCol.Visible = False
''        lacAsAirNoChars.Visible = False
'    End If
    If sgClientFields = "A" Then
        grdExport.ScrollBars = flexScrollBarHorizontal
        imMaxCols = ABCRECORDITEMINDEX
    Else
        imMaxCols = EVENTIDINDEX
    End If
    If cbcSelect.ListCount = 2 Then
        cbcSelect.ListIndex = 1
    ElseIf cbcSelect.ListCount >= 1 Then
        cbcSelect.ListIndex = 0
    End If
    imFieldChgd = False
    mSetCommands
End Sub

Private Sub mInitAFE(tlAFE As AFE, slType As String, slSubType As String)
    If slSubType = "S" Then 'Start Column
        tlAFE.iBus = 1
        tlAFE.iBusControl = 91
        tlAFE.iDate = 6
        tlAFE.iEventID = 14
        tlAFE.iEventType = 22
        tlAFE.iTime = 23
        tlAFE.iStartType = 44
        tlAFE.iFixedTime = 33
        tlAFE.iEndType = 47
        tlAFE.iDuration = 50
        tlAFE.iEndTime = 34
        tlAFE.iMaterialType = 60
        tlAFE.iAudioName = 63
        tlAFE.iAudioItemID = 69
        tlAFE.iAudioControl = 68
        tlAFE.iBkupAudioName = 74
        tlAFE.iBkupAudioControl = 79
        tlAFE.iProtAudioName = 80
        tlAFE.iProtItemID = 86
        tlAFE.iProtAudioControl = 85
        tlAFE.iRelay1 = 98
        tlAFE.iRelay2 = 103
        tlAFE.iFollow = 108
        tlAFE.iSilenceTime = 115
        tlAFE.iSilence1 = 120
        tlAFE.iSilence2 = 121
        tlAFE.iSilence3 = 122
        tlAFE.iSilence4 = 123
        tlAFE.iStartNetcue = 92
        tlAFE.iStopNetcue = 95
        tlAFE.iTitle1 = 124
        tlAFE.iTitle2 = 190
        tlAFE.iAudioISCI = 0
        tlAFE.iProtISCI = 0
        tlAFE.iABCFormat = 0
        tlAFE.iABCPgmCode = 0
        tlAFE.iABCXDSMode = 0
        tlAFE.iABCRecordItem = 0
        tlAFE.iCode = 0
    Else        'Number of characters
        tlAFE.iBus = 5
        tlAFE.iBusControl = 1
        tlAFE.iDate = 8
        tlAFE.iEventID = 8
        tlAFE.iEventType = 1
        tlAFE.iTime = 10
        tlAFE.iStartType = 3
        tlAFE.iFixedTime = 1
        tlAFE.iEndType = 3
        tlAFE.iDuration = 10
        tlAFE.iEndTime = 10
        tlAFE.iMaterialType = 3
        tlAFE.iAudioName = 5
        tlAFE.iAudioItemID = 5
        tlAFE.iAudioControl = 1
        tlAFE.iBkupAudioName = 5
        tlAFE.iBkupAudioControl = 1
        tlAFE.iProtAudioName = 5
        tlAFE.iProtItemID = 5
        tlAFE.iProtAudioControl = 1
        tlAFE.iRelay1 = 5
        tlAFE.iRelay2 = 5
        tlAFE.iFollow = 7
        tlAFE.iSilenceTime = 5
        tlAFE.iSilence1 = 1
        tlAFE.iSilence2 = 1
        tlAFE.iSilence3 = 1
        tlAFE.iSilence4 = 1
        tlAFE.iStartNetcue = 3
        tlAFE.iStopNetcue = 3
        tlAFE.iTitle1 = 66
        tlAFE.iTitle2 = 66
        tlAFE.iAudioISCI = 20
        tlAFE.iProtISCI = 20
        tlAFE.iABCFormat = 1
        tlAFE.iABCPgmCode = 25
        tlAFE.iABCXDSMode = 2
        tlAFE.iABCRecordItem = 5
        tlAFE.iCode = 0
    End If

End Sub


Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    
    ilRet = gGetTypeOfRecs_AEE_AutoEquip("C", sgCurrAEEStamp, "EngrAutomation-mPopulate", tgCurrAEE())
    ilRet = gGetTypeOfRecs_ACE_AutoContact("C", sgCurrACEStamp, "EngrEventType-mPopulate", tgCurrACE())
    ilRet = gGetTypeOfRecs_ADE_AutoDataFlags("C", sgCurrADEStamp, "EngrEventType-mPopulate", tgCurrADE())
    ilRet = gGetTypeOfRecs_AFE_AutoFormat("C", sgCurrAFEStamp, "EngrEventType-mPopulate", tgCurrAFE())
    ilRet = gGetTypeOfRecs_APE_AutoPath("C", sgCurrAPEStamp, "EngrEventType-mPopulate", tgCurrAPE())
    
    cbcSelect.Clear
    cbcSelect.text = ""
    '
    'At the current time only allow one Automation equipment definition
    'When we want to allow multi-automation equipment:
    '1.  Add password to this function
    '2.  Determine at which point we need to ask for which automation equipment
    '    Are library and templates independent of the automation equipment
    '    (they might not be as the size might be different)
    '    If the schedule different?
    '    Is it only the export that would be different?
    '    If the Libraries, template and schedule are different, do we just
    '    code for a combination (one does not require silence control but the
    '    other does, therefore we permit the silence control to be defined.  The export would not include it)
    '
    If UBound(tgCurrAEE) > LBound(tgCurrAEE) Then
        For ilLoop = 0 To UBound(tgCurrAEE) - 1 Step 1
            cbcSelect.AddItem Trim$(tgCurrAEE(ilLoop).sName)
            cbcSelect.ItemData(cbcSelect.NewIndex) = tgCurrAEE(ilLoop).iCode
        Next ilLoop
    Else
        cbcSelect.AddItem "[New]", 0
        cbcSelect.ItemData(cbcSelect.NewIndex) = 0
    End If
End Sub

Private Sub mMoveCtrlsToRec()
    Dim ilLoop As Integer
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    tmAEE.iCode = imAeeCode
    tmAEE.sName = edcName.text
    tmAEE.sDescription = edcDescription.text
    tmAEE.sManufacture = edcManufacture.text
    tmAEE.sFixedTimeChar = edcFixedTimeChar.text
    tmAEE.lAlertSchdDelay = gLengthToLong(edcDelay.text)
    If rbcState(1).Value Then
        tmAEE.sState = "D"
    Else
        tmAEE.sState = "A"
    End If
    tmAEE.sUsedFlag = smUsedFlag
    tmAEE.iVersion = imVersion + 1
    tmAEE.iOrigAeeCode = imAeeCode
    tmAEE.sCurrent = "Y"
    'tmAEE.sEnteredDate = smNowDate
    'tmAEE.sEnteredTime = smNowTime
    tmAEE.sEnteredDate = Format(Now, sgShowDateForm)
    tmAEE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmAEE.iUieCode = tgUIE.iCode
    tmAEE.sUnused = ""
    For ilLoop = 0 To UBound(tmCurrACE) - 1 Step 1
        If ilLoop = 0 Then
            tmCurrACE(ilLoop).sType = "P"
            tmCurrACE(ilLoop).sContact = edcPriContactName.text
            tmCurrACE(ilLoop).sPhone = edcPriPhone.text
            tmCurrACE(ilLoop).sFax = edcPriFax.text
            tmCurrACE(ilLoop).sEMail = edcPriEMail.text
        Else
            tmCurrACE(ilLoop).sType = "S"
            tmCurrACE(ilLoop).sContact = edcSecContactName.text
            tmCurrACE(ilLoop).sPhone = edcSecPhone.text
            tmCurrACE(ilLoop).sFax = edcSecFax.text
            tmCurrACE(ilLoop).sEMail = edcSecEMail.text
        End If
        tmCurrACE(ilLoop).sUnused = ""
    Next ilLoop
    
    For ilLoop = 0 To UBound(tmCurrAPE) - 1 Step 1
        If ilLoop = 0 Then
            tmCurrAPE(ilLoop).sType = "SE"
            tmCurrAPE(ilLoop).sSubType = "P"
            tmCurrAPE(ilLoop).sDateFormat = edcDateFormat.text
            tmCurrAPE(ilLoop).sTimeFormat = edcTimeFormat.text
            tmCurrAPE(ilLoop).sNewFileName = edcExportNewFileFormat.text
            tmCurrAPE(ilLoop).sChgFileName = edcExportChgFileFormat.text
            tmCurrAPE(ilLoop).sDelFileName = edcExportDelFileFormat.text
            tmCurrAPE(ilLoop).sNewFileExt = edcExportExtNew.text
            tmCurrAPE(ilLoop).sChgFileExt = edcExportExtChg.text
            tmCurrAPE(ilLoop).sDelFileExt = edcExportExtDel.text
            tmCurrAPE(ilLoop).sPath = edcServerExportPath.text
        ElseIf ilLoop = 1 Then
            tmCurrAPE(ilLoop).sType = "SI"
            tmCurrAPE(ilLoop).sSubType = "P"
            tmCurrAPE(ilLoop).sDateFormat = edcDateFormat.text
            tmCurrAPE(ilLoop).sTimeFormat = edcTimeFormat.text
            tmCurrAPE(ilLoop).sNewFileName = edcImportFileFormat.text
            tmCurrAPE(ilLoop).sChgFileName = ""
            tmCurrAPE(ilLoop).sDelFileName = ""
            tmCurrAPE(ilLoop).sNewFileExt = edcImportExt.text
            tmCurrAPE(ilLoop).sChgFileExt = ""
            tmCurrAPE(ilLoop).sDelFileExt = ""
            tmCurrAPE(ilLoop).sPath = edcServerImportPath.text
        ElseIf ilLoop = 2 Then
            tmCurrAPE(ilLoop).sType = "CE"
            tmCurrAPE(ilLoop).sSubType = "P"
            tmCurrAPE(ilLoop).sDateFormat = edcDateFormat.text
            tmCurrAPE(ilLoop).sTimeFormat = edcTimeFormat.text
            tmCurrAPE(ilLoop).sNewFileName = edcExportNewFileFormat.text
            tmCurrAPE(ilLoop).sChgFileName = edcExportChgFileFormat.text
            tmCurrAPE(ilLoop).sDelFileName = edcExportDelFileFormat.text
            tmCurrAPE(ilLoop).sNewFileExt = edcExportExtNew.text
            tmCurrAPE(ilLoop).sChgFileExt = edcExportExtChg.text
            tmCurrAPE(ilLoop).sDelFileExt = edcExportExtDel.text
            tmCurrAPE(ilLoop).sPath = edcClientExportPath.text
        ElseIf ilLoop = 3 Then
            tmCurrAPE(ilLoop).sType = "CI"
            tmCurrAPE(ilLoop).sSubType = "P"
            tmCurrAPE(ilLoop).sDateFormat = edcDateFormat.text
            tmCurrAPE(ilLoop).sTimeFormat = edcTimeFormat.text
            tmCurrAPE(ilLoop).sNewFileName = edcImportFileFormat.text
            tmCurrAPE(ilLoop).sChgFileName = ""
            tmCurrAPE(ilLoop).sDelFileName = ""
            tmCurrAPE(ilLoop).sNewFileExt = edcImportExt.text
            tmCurrAPE(ilLoop).sChgFileExt = ""
            tmCurrAPE(ilLoop).sDelFileExt = ""
            tmCurrAPE(ilLoop).sPath = edcClientImportPath.text
        ElseIf ilLoop = 4 Then
            tmCurrAPE(ilLoop).sType = "SE"
            tmCurrAPE(ilLoop).sSubType = "T"
            tmCurrAPE(ilLoop).sDateFormat = edcDateFormat.text
            tmCurrAPE(ilLoop).sTimeFormat = edcTimeFormat.text
            tmCurrAPE(ilLoop).sNewFileName = edcExportNewFileFormat.text
            tmCurrAPE(ilLoop).sChgFileName = edcExportChgFileFormat.text
            tmCurrAPE(ilLoop).sDelFileName = edcExportDelFileFormat.text
            tmCurrAPE(ilLoop).sNewFileExt = edcExportExtNew.text
            tmCurrAPE(ilLoop).sChgFileExt = edcExportExtChg.text
            tmCurrAPE(ilLoop).sDelFileExt = edcExportExtDel.text
            tmCurrAPE(ilLoop).sPath = edcServerExportPathTest.text
        ElseIf ilLoop = 5 Then
            tmCurrAPE(ilLoop).sType = "CE"
            tmCurrAPE(ilLoop).sSubType = "T"
            tmCurrAPE(ilLoop).sDateFormat = edcDateFormat.text
            tmCurrAPE(ilLoop).sTimeFormat = edcTimeFormat.text
            tmCurrAPE(ilLoop).sNewFileName = edcExportNewFileFormat.text
            tmCurrAPE(ilLoop).sChgFileName = edcExportChgFileFormat.text
            tmCurrAPE(ilLoop).sDelFileName = edcExportDelFileFormat.text
            tmCurrAPE(ilLoop).sNewFileExt = edcExportExtNew.text
            tmCurrAPE(ilLoop).sChgFileExt = edcExportExtChg.text
            tmCurrAPE(ilLoop).sDelFileExt = edcExportExtDel.text
            tmCurrAPE(ilLoop).sPath = edcClientExportPathTest.text
        End If
        tmCurrAPE(ilLoop).sUnused = ""
    Next ilLoop

End Sub

Private Sub mMoveADECtrlsToRec()
    Dim ilPass As Long
    Dim llPRow As Long
    
    ilPass = 0
    llPRow = 3
    tmCurrADE(ilPass).iScheduleData = Val(grdImport.TextMatrix(llPRow, ECHOSTARTINDEX))
    tmCurrADE(ilPass).iDate = Val(grdImport.TextMatrix(llPRow, DATESTARTINDEX))
    tmCurrADE(ilPass).iDateNoChar = Val(grdImport.TextMatrix(llPRow, DATENOCHARINDEX))
    tmCurrADE(ilPass).iTime = Val(grdImport.TextMatrix(llPRow, TIMESTARTINDEX))
    tmCurrADE(ilPass).iTimeNoChar = Val(grdImport.TextMatrix(llPRow, TIMENOCHARINDEX))
    tmCurrADE(ilPass).iAutoOff = Val(grdImport.TextMatrix(llPRow, AUTOOFFINDEX))
    tmCurrADE(ilPass).iData = Val(grdImport.TextMatrix(llPRow, DATAINDEX))
    tmCurrADE(ilPass).iSchedule = Val(grdImport.TextMatrix(llPRow, SCHEDULEINDEX))
    tmCurrADE(ilPass).iTrueTime = Val(grdImport.TextMatrix(llPRow, TRUETIMEINDEX))
    tmCurrADE(ilPass).iSourceConflict = Val(grdImport.TextMatrix(llPRow, SRCECONFLICTINDEX))
    tmCurrADE(ilPass).iSourceUnavail = Val(grdImport.TextMatrix(llPRow, SRCEUNAVAILINDEX))
    tmCurrADE(ilPass).iSourceItem = Val(grdImport.TextMatrix(llPRow, SRCEITEMINDEX))
    tmCurrADE(ilPass).iBkupSrceUnavail = Val(grdImport.TextMatrix(llPRow, BKUPUNAVAILINDEX))
    tmCurrADE(ilPass).iBkupSrceItem = Val(grdImport.TextMatrix(llPRow, BKUPITEMINDEX))
    tmCurrADE(ilPass).iProtSrceUnavail = Val(grdImport.TextMatrix(llPRow, PROTUNAVAILINDEX))
    tmCurrADE(ilPass).iProtSrceItem = Val(grdImport.TextMatrix(llPRow, PROTITEMINDEX))
    tmCurrADE(ilPass).sUnused = ""
    tmCurrADE(ilPass).iCode = Val(grdImport.TextMatrix(llPRow, ADECODEINDEX))
    
End Sub
Private Sub mMoveAFECtrlsToRec()
    Dim ilPass As Long
    Dim llPRow As Long
    
'    For ilPass = 0 To 3 Step 1
    For ilPass = 0 To 1 Step 1
        If ilPass = 0 Then
            llPRow = 3
        ElseIf ilPass = 1 Then
            llPRow = 4
'        ElseIf ilPass = 2 Then
'            llPRow = 6
'        Else
'            llPRow = 7
        End If
        tmCurrAFE(ilPass).iBus = Val(grdExport.TextMatrix(llPRow, BUSNAMEINDEX))
        tmCurrAFE(ilPass).iBusControl = Val(grdExport.TextMatrix(llPRow, BUSCTRLINDEX))
        tmCurrAFE(ilPass).iEventType = Val(grdExport.TextMatrix(llPRow, EVENTTYPEINDEX))
        tmCurrAFE(ilPass).iTime = Val(grdExport.TextMatrix(llPRow, TIMEINDEX))
        tmCurrAFE(ilPass).iStartType = Val(grdExport.TextMatrix(llPRow, STARTTYPEINDEX))
        tmCurrAFE(ilPass).iFixedTime = Val(grdExport.TextMatrix(llPRow, FIXEDINDEX))
        tmCurrAFE(ilPass).iEndType = Val(grdExport.TextMatrix(llPRow, ENDTYPEINDEX))
        tmCurrAFE(ilPass).iDuration = Val(grdExport.TextMatrix(llPRow, DURATIONINDEX))
        tmCurrAFE(ilPass).iEndTime = Val(grdExport.TextMatrix(llPRow, ENDTIMEINDEX))
        tmCurrAFE(ilPass).iMaterialType = Val(grdExport.TextMatrix(llPRow, MATERIALINDEX))
        tmCurrAFE(ilPass).iAudioName = Val(grdExport.TextMatrix(llPRow, AUDIONAMEINDEX))
        tmCurrAFE(ilPass).iAudioItemID = Val(grdExport.TextMatrix(llPRow, AUDIOITEMIDINDEX))
        tmCurrAFE(ilPass).iAudioISCI = Val(grdExport.TextMatrix(llPRow, AUDIOISCIINDEX))
        tmCurrAFE(ilPass).iAudioControl = Val(grdExport.TextMatrix(llPRow, AUDIOCTRLINDEX))
        tmCurrAFE(ilPass).iBkupAudioName = Val(grdExport.TextMatrix(llPRow, BACKUPNAMEINDEX))
        tmCurrAFE(ilPass).iBkupAudioControl = Val(grdExport.TextMatrix(llPRow, BACKUPCTRLINDEX))
        tmCurrAFE(ilPass).iProtAudioName = Val(grdExport.TextMatrix(llPRow, PROTNAMEINDEX))
        tmCurrAFE(ilPass).iProtItemID = Val(grdExport.TextMatrix(llPRow, PROTITEMIDINDEX))
        tmCurrAFE(ilPass).iProtISCI = Val(grdExport.TextMatrix(llPRow, PROTISCIINDEX))
        tmCurrAFE(ilPass).iProtAudioControl = Val(grdExport.TextMatrix(llPRow, PROTCTRLINDEX))
        tmCurrAFE(ilPass).iRelay1 = Val(grdExport.TextMatrix(llPRow, RELAY1INDEX))
        tmCurrAFE(ilPass).iRelay2 = Val(grdExport.TextMatrix(llPRow, RELAY2INDEX))
        tmCurrAFE(ilPass).iFollow = Val(grdExport.TextMatrix(llPRow, FOLLOWINDEX))
        tmCurrAFE(ilPass).iSilenceTime = Val(grdExport.TextMatrix(llPRow, SILENCETIMEINDEX))
        tmCurrAFE(ilPass).iSilence1 = Val(grdExport.TextMatrix(llPRow, SILENCE1INDEX))
        tmCurrAFE(ilPass).iSilence2 = Val(grdExport.TextMatrix(llPRow, SILENCE2INDEX))
        tmCurrAFE(ilPass).iSilence3 = Val(grdExport.TextMatrix(llPRow, SILENCE3INDEX))
        tmCurrAFE(ilPass).iSilence4 = Val(grdExport.TextMatrix(llPRow, SILENCE4INDEX))
        tmCurrAFE(ilPass).iStartNetcue = Val(grdExport.TextMatrix(llPRow, NETCUE1INDEX))
        tmCurrAFE(ilPass).iStopNetcue = Val(grdExport.TextMatrix(llPRow, NETCUE2INDEX))
        tmCurrAFE(ilPass).iTitle1 = Val(grdExport.TextMatrix(llPRow, TITLE1INDEX))
        tmCurrAFE(ilPass).iTitle2 = Val(grdExport.TextMatrix(llPRow, TITLE2INDEX))
        tmCurrAFE(ilPass).iTitle2 = Val(grdExport.TextMatrix(llPRow, TITLE2INDEX))
        tmCurrAFE(ilPass).iEventID = Val(grdExport.TextMatrix(llPRow, EVENTIDINDEX))
        tmCurrAFE(ilPass).iDate = Val(grdExport.TextMatrix(llPRow, DATEINDEX))
        If sgClientFields = "A" Then
            tmCurrAFE(ilPass).iABCFormat = Val(grdExport.TextMatrix(llPRow, ABCFORMATINDEX))
            tmCurrAFE(ilPass).iABCPgmCode = Val(grdExport.TextMatrix(llPRow, ABCPGMCODEINDEX))
            tmCurrAFE(ilPass).iABCXDSMode = Val(grdExport.TextMatrix(llPRow, ABCXDSMODEINDEX))
            tmCurrAFE(ilPass).iABCRecordItem = Val(grdExport.TextMatrix(llPRow, ABCRECORDITEMINDEX))
        Else
            tmCurrAFE(ilPass).iABCFormat = 0
            tmCurrAFE(ilPass).iABCPgmCode = 0
            tmCurrAFE(ilPass).iABCXDSMode = 0
            tmCurrAFE(ilPass).iABCRecordItem = 0
        End If
        tmCurrAFE(ilPass).sUnused = ""
        tmCurrAFE(ilPass).iCode = Val(grdExport.TextMatrix(llPRow, CODEINDEX))
    Next ilPass
    
End Sub
Private Sub mMoveADERecToCtrls()
    Dim llPRow As Long
    Dim ilPass As Integer
    
    If imAeeCode <= 0 Then
        ReDim tmCurrADE(0 To 1) As ADE
    End If
    llPRow = 3
    ilPass = 0
    If imAeeCode <= 0 Then
        mInitADE tmCurrADE(ilPass)
    End If
    grdImport.TextMatrix(llPRow, ECHOSTARTINDEX) = tmCurrADE(ilPass).iScheduleData
    grdImport.TextMatrix(llPRow, DATESTARTINDEX) = tmCurrADE(ilPass).iDate
    grdImport.TextMatrix(llPRow, DATENOCHARINDEX) = tmCurrADE(ilPass).iDateNoChar
    grdImport.TextMatrix(llPRow, TIMESTARTINDEX) = tmCurrADE(ilPass).iTime
    grdImport.TextMatrix(llPRow, TIMENOCHARINDEX) = tmCurrADE(ilPass).iTimeNoChar
    grdImport.TextMatrix(llPRow, AUTOOFFINDEX) = tmCurrADE(ilPass).iAutoOff
    grdImport.TextMatrix(llPRow, DATAINDEX) = tmCurrADE(ilPass).iData
    grdImport.TextMatrix(llPRow, SCHEDULEINDEX) = tmCurrADE(ilPass).iSchedule
    grdImport.TextMatrix(llPRow, TRUETIMEINDEX) = tmCurrADE(ilPass).iTrueTime
    grdImport.TextMatrix(llPRow, SRCECONFLICTINDEX) = tmCurrADE(ilPass).iSourceConflict
    grdImport.TextMatrix(llPRow, SRCEUNAVAILINDEX) = tmCurrADE(ilPass).iSourceUnavail
    grdImport.TextMatrix(llPRow, SRCEITEMINDEX) = tmCurrADE(ilPass).iSourceItem
    grdImport.TextMatrix(llPRow, BKUPUNAVAILINDEX) = tmCurrADE(ilPass).iBkupSrceUnavail
    grdImport.TextMatrix(llPRow, BKUPITEMINDEX) = tmCurrADE(ilPass).iBkupSrceItem
    grdImport.TextMatrix(llPRow, PROTUNAVAILINDEX) = tmCurrADE(ilPass).iProtSrceUnavail
    grdImport.TextMatrix(llPRow, PROTITEMINDEX) = tmCurrADE(ilPass).iProtSrceItem
    grdImport.TextMatrix(llPRow, ADECODEINDEX) = tmCurrADE(ilPass).iCode
    
End Sub
Private Sub mMoveAFERecToCtrls()
    Dim ilPass As Long
    Dim llPRow As Long
    
    If imAeeCode <= 0 Then
        ReDim tmCurrAFE(0 To 2) As AFE
        tmCurrAFE(0).sType = "S"
        tmCurrAFE(0).sSubType = "S"
        mInitAFE tmCurrAFE(0), tmCurrAFE(0).sType, tmCurrAFE(0).sSubType
        tmCurrAFE(1).sType = "S"
        tmCurrAFE(1).sSubType = "N"
        mInitAFE tmCurrAFE(1), tmCurrAFE(1).sType, tmCurrAFE(1).sSubType
    End If
'    For ilPass = 0 To 3 Step 1
    For ilPass = 0 To 1 Step 1
        If ilPass = 0 Then
            llPRow = 3
        ElseIf ilPass = 1 Then
            llPRow = 4
'        ElseIf ilPass = 2 Then
'            llPRow = 6
'            If imAeeCode <= 0 Then
'                tmCurrAFE(ilPass).sType = "A"
'                tmCurrAFE(ilPass).sSubType = "S"
'                mInitAFE tmCurrAFE(ilPass), tmCurrAFE(ilPass).sType, tmCurrAFE(ilPass).sSubType
'            End If
'        Else
'            llPRow = 7
'            If imAeeCode <= 0 Then
'                tmCurrAFE(ilPass).sType = "A"
'                tmCurrAFE(ilPass).sSubType = "N"
'                mInitAFE tmCurrAFE(ilPass), tmCurrAFE(ilPass).sType, tmCurrAFE(ilPass).sSubType
'            End If
        End If
        grdExport.TextMatrix(llPRow, BUSNAMEINDEX) = tmCurrAFE(ilPass).iBus
        grdExport.TextMatrix(llPRow, BUSCTRLINDEX) = tmCurrAFE(ilPass).iBusControl
        grdExport.TextMatrix(llPRow, EVENTTYPEINDEX) = tmCurrAFE(ilPass).iEventType
        grdExport.TextMatrix(llPRow, TIMEINDEX) = tmCurrAFE(ilPass).iTime
        grdExport.TextMatrix(llPRow, STARTTYPEINDEX) = tmCurrAFE(ilPass).iStartType
        grdExport.TextMatrix(llPRow, FIXEDINDEX) = tmCurrAFE(ilPass).iFixedTime
        grdExport.TextMatrix(llPRow, ENDTYPEINDEX) = tmCurrAFE(ilPass).iEndType
        grdExport.TextMatrix(llPRow, DURATIONINDEX) = tmCurrAFE(ilPass).iDuration
        grdExport.TextMatrix(llPRow, ENDTIMEINDEX) = tmCurrAFE(ilPass).iEndTime
        grdExport.TextMatrix(llPRow, MATERIALINDEX) = tmCurrAFE(ilPass).iMaterialType
        grdExport.TextMatrix(llPRow, AUDIONAMEINDEX) = tmCurrAFE(ilPass).iAudioName
        grdExport.TextMatrix(llPRow, AUDIOITEMIDINDEX) = tmCurrAFE(ilPass).iAudioItemID
        grdExport.TextMatrix(llPRow, AUDIOISCIINDEX) = tmCurrAFE(ilPass).iAudioISCI
        grdExport.TextMatrix(llPRow, AUDIOCTRLINDEX) = tmCurrAFE(ilPass).iAudioControl
        grdExport.TextMatrix(llPRow, BACKUPNAMEINDEX) = tmCurrAFE(ilPass).iBkupAudioName
        grdExport.TextMatrix(llPRow, BACKUPCTRLINDEX) = tmCurrAFE(ilPass).iBkupAudioControl
        grdExport.TextMatrix(llPRow, PROTNAMEINDEX) = tmCurrAFE(ilPass).iProtAudioName
        grdExport.TextMatrix(llPRow, PROTITEMIDINDEX) = tmCurrAFE(ilPass).iProtItemID
        grdExport.TextMatrix(llPRow, PROTISCIINDEX) = tmCurrAFE(ilPass).iProtISCI
        grdExport.TextMatrix(llPRow, PROTCTRLINDEX) = tmCurrAFE(ilPass).iProtAudioControl
        grdExport.TextMatrix(llPRow, RELAY1INDEX) = tmCurrAFE(ilPass).iRelay1
        grdExport.TextMatrix(llPRow, RELAY2INDEX) = tmCurrAFE(ilPass).iRelay2
        grdExport.TextMatrix(llPRow, FOLLOWINDEX) = tmCurrAFE(ilPass).iFollow
        grdExport.TextMatrix(llPRow, SILENCETIMEINDEX) = tmCurrAFE(ilPass).iSilenceTime
        grdExport.TextMatrix(llPRow, SILENCE1INDEX) = tmCurrAFE(ilPass).iSilence1
        grdExport.TextMatrix(llPRow, SILENCE2INDEX) = tmCurrAFE(ilPass).iSilence2
        grdExport.TextMatrix(llPRow, SILENCE3INDEX) = tmCurrAFE(ilPass).iSilence3
        grdExport.TextMatrix(llPRow, SILENCE4INDEX) = tmCurrAFE(ilPass).iSilence4
        grdExport.TextMatrix(llPRow, NETCUE1INDEX) = tmCurrAFE(ilPass).iStartNetcue
        grdExport.TextMatrix(llPRow, NETCUE2INDEX) = tmCurrAFE(ilPass).iStopNetcue
        grdExport.TextMatrix(llPRow, TITLE1INDEX) = tmCurrAFE(ilPass).iTitle1
        grdExport.TextMatrix(llPRow, TITLE2INDEX) = tmCurrAFE(ilPass).iTitle2
        grdExport.TextMatrix(llPRow, TITLE2INDEX) = tmCurrAFE(ilPass).iTitle2
        grdExport.TextMatrix(llPRow, EVENTIDINDEX) = tmCurrAFE(ilPass).iEventID
        grdExport.TextMatrix(llPRow, DATEINDEX) = tmCurrAFE(ilPass).iDate
        If sgClientFields = "A" Then
            grdExport.TextMatrix(llPRow, ABCFORMATINDEX) = tmCurrAFE(ilPass).iABCFormat
            grdExport.TextMatrix(llPRow, ABCPGMCODEINDEX) = tmCurrAFE(ilPass).iABCPgmCode
            grdExport.TextMatrix(llPRow, ABCXDSMODEINDEX) = tmCurrAFE(ilPass).iABCXDSMode
            grdExport.TextMatrix(llPRow, ABCRECORDITEMINDEX) = tmCurrAFE(ilPass).iABCRecordItem
        Else
            grdExport.TextMatrix(llPRow, ABCFORMATINDEX) = 0
            grdExport.TextMatrix(llPRow, ABCPGMCODEINDEX) = 0
            grdExport.TextMatrix(llPRow, ABCXDSMODEINDEX) = 0
            grdExport.TextMatrix(llPRow, ABCRECORDITEMINDEX) = 0
        End If
        grdExport.TextMatrix(llPRow, CODEINDEX) = tmCurrAFE(ilPass).iCode
    Next ilPass
    grdExport.Refresh
End Sub


Private Sub mMoveRecToCtrls(ilAEEIndex As Integer)
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    
    edcName.text = Trim$(tgCurrAEE(ilAEEIndex).sName)
    edcDescription.text = Trim$(tgCurrAEE(ilAEEIndex).sDescription)
    edcManufacture.text = Trim$(tgCurrAEE(ilAEEIndex).sManufacture)
    edcFixedTimeChar.text = Trim$(tgCurrAEE(ilAEEIndex).sFixedTimeChar)
    edcDelay.text = gLongToLength(tgCurrAEE(ilAEEIndex).lAlertSchdDelay, True)
    ReDim tmCurrACE(0 To 2) As ACE
    For ilLoop = LBound(tgCurrACE) To UBound(tgCurrACE) - 1 Step 1
        If (tgCurrACE(ilLoop).sType = "P") And (tgCurrACE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrACE(0) = tgCurrACE(ilLoop)
            edcPriContactName.text = Trim$(tgCurrACE(ilLoop).sContact)
            edcPriPhone.text = Trim$(tgCurrACE(ilLoop).sPhone)
            edcPriFax.text = Trim$(tgCurrACE(ilLoop).sFax)
            edcPriEMail.text = Trim$(tgCurrACE(ilLoop).sEMail)
            Exit For
        End If
    Next ilLoop
    For ilLoop = LBound(tgCurrACE) To UBound(tgCurrACE) - 1 Step 1
        If (tgCurrACE(ilLoop).sType = "S") And (tgCurrACE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrACE(1) = tgCurrACE(ilLoop)
            edcSecContactName.text = Trim$(tgCurrACE(ilLoop).sContact)
            edcSecPhone.text = Trim$(tgCurrACE(ilLoop).sPhone)
            edcSecFax.text = Trim$(tgCurrACE(ilLoop).sFax)
            edcSecEMail.text = Trim$(tgCurrACE(ilLoop).sEMail)
            Exit For
        End If
    Next ilLoop
    
    ReDim tmCurrAFE(0 To 2) As AFE
    For ilLoop = LBound(tgCurrAFE) To UBound(tgCurrAFE) - 1 Step 1
        If (tgCurrAFE(ilLoop).sSubType = "S") And (tgCurrAFE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrAFE(0) = tgCurrAFE(ilLoop)
            Exit For
        End If
    Next ilLoop
    For ilLoop = LBound(tgCurrAFE) To UBound(tgCurrAFE) - 1 Step 1
        If (tgCurrAFE(ilLoop).sSubType = "N") And (tgCurrAFE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrAFE(1) = tgCurrAFE(ilLoop)
            Exit For
        End If
    Next ilLoop
    mMoveAFERecToCtrls
    
    ReDim tmCurrADE(0 To 1) As ADE
    For ilLoop = LBound(tgCurrADE) To UBound(tgCurrADE) - 1 Step 1
        If (tgCurrADE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrADE(0) = tgCurrADE(ilLoop)
            Exit For
        End If
    Next ilLoop
    mMoveADERecToCtrls
    
    ReDim tmCurrAPE(0 To 6) As APE
    For ilLoop = LBound(tgCurrAPE) To UBound(tgCurrAPE) - 1 Step 1
        If (tgCurrAPE(ilLoop).sType = "SE") And ((tgCurrAPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrAPE(ilLoop).sSubType) = "")) And (tgCurrAPE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrAPE(0) = tgCurrAPE(ilLoop)
            edcDateFormat.text = Trim$(tgCurrAPE(ilLoop).sDateFormat)
            edcTimeFormat.text = Trim$(tgCurrAPE(ilLoop).sTimeFormat)
            edcExportNewFileFormat.text = Trim$(tgCurrAPE(ilLoop).sNewFileName)
            edcExportChgFileFormat.text = Trim$(tgCurrAPE(ilLoop).sChgFileName)
            edcExportDelFileFormat.text = Trim$(tgCurrAPE(ilLoop).sDelFileName)
            edcExportExtNew.text = Trim$(tgCurrAPE(ilLoop).sNewFileExt)
            edcExportExtChg.text = Trim$(tgCurrAPE(ilLoop).sChgFileExt)
            edcExportExtDel.text = Trim$(tgCurrAPE(ilLoop).sDelFileExt)
            edcServerExportPath.text = Trim$(tgCurrAPE(ilLoop).sPath)
            Exit For
        End If
    Next ilLoop
    For ilLoop = LBound(tgCurrAPE) To UBound(tgCurrAPE) - 1 Step 1
        If (tgCurrAPE(ilLoop).sType = "SI") And (tgCurrAPE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrAPE(1) = tgCurrAPE(ilLoop)
            edcImportFileFormat.text = Trim$(tgCurrAPE(ilLoop).sNewFileName)
            edcImportExt.text = Trim$(tgCurrAPE(ilLoop).sNewFileExt)
            edcServerImportPath.text = Trim$(tgCurrAPE(ilLoop).sPath)
            Exit For
        End If
    Next ilLoop
    For ilLoop = LBound(tgCurrAPE) To UBound(tgCurrAPE) - 1 Step 1
        If (tgCurrAPE(ilLoop).sType = "CE") And ((tgCurrAPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrAPE(ilLoop).sSubType) = "")) And (tgCurrAPE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrAPE(2) = tgCurrAPE(ilLoop)
            edcClientExportPath.text = Trim$(tgCurrAPE(ilLoop).sPath)
            Exit For
        End If
    Next ilLoop
    For ilLoop = LBound(tgCurrAPE) To UBound(tgCurrAPE) - 1 Step 1
        If (tgCurrAPE(ilLoop).sType = "CI") And (tgCurrAPE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrAPE(3) = tgCurrAPE(ilLoop)
            edcClientImportPath.text = Trim$(tgCurrAPE(ilLoop).sPath)
            Exit For
        End If
    Next ilLoop
    For ilLoop = LBound(tgCurrAPE) To UBound(tgCurrAPE) - 1 Step 1
        If (tgCurrAPE(ilLoop).sType = "SE") And (tgCurrAPE(ilLoop).sSubType = "T") And (tgCurrAPE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrAPE(4) = tgCurrAPE(ilLoop)
            edcDateFormat.text = Trim$(tgCurrAPE(ilLoop).sDateFormat)
            edcTimeFormat.text = Trim$(tgCurrAPE(ilLoop).sTimeFormat)
            edcExportNewFileFormat.text = Trim$(tgCurrAPE(ilLoop).sNewFileName)
            edcExportChgFileFormat.text = Trim$(tgCurrAPE(ilLoop).sChgFileName)
            edcExportDelFileFormat.text = Trim$(tgCurrAPE(ilLoop).sDelFileName)
            edcExportExtNew.text = Trim$(tgCurrAPE(ilLoop).sNewFileExt)
            edcExportExtChg.text = Trim$(tgCurrAPE(ilLoop).sChgFileExt)
            edcExportExtDel.text = Trim$(tgCurrAPE(ilLoop).sDelFileExt)
            edcServerExportPathTest.text = Trim$(tgCurrAPE(ilLoop).sPath)
            Exit For
        End If
    Next ilLoop
    For ilLoop = LBound(tgCurrAPE) To UBound(tgCurrAPE) - 1 Step 1
        If (tgCurrAPE(ilLoop).sType = "CE") And (tgCurrAPE(ilLoop).sSubType = "T") And (tgCurrAPE(ilLoop).iAeeCode = imAeeCode) Then
            LSet tmCurrAPE(5) = tgCurrAPE(ilLoop)
            edcClientExportPathTest.text = Trim$(tgCurrAPE(ilLoop).sPath)
            Exit For
        End If
    Next ilLoop

    
    If tgCurrAEE(ilAEEIndex).sState = "D" Then
        rbcState(1).Value = True
    Else
        rbcState(0).Value = True
    End If
    imVersion = tgCurrAEE(ilAEEIndex).iVersion
    smUsedFlag = tgCurrAEE(ilAEEIndex).sUsedFlag
End Sub
Private Sub mSetCommands()
    Dim ilRet As Integer
    If imInChg Then
        Exit Sub
    End If
    If cmcDone.Enabled = False Then
        Exit Sub
    End If
    If imFieldChgd Then
        'Check that all mandatory answered
        ilRet = mCheckFields(False)
        If ilRet Then
            cmcSave.Enabled = True
            cbcSelect.Enabled = False
            cmcErase.Enabled = False
        Else
            cmcSave.Enabled = False
            cbcSelect.Enabled = False
            cmcErase.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
        cbcSelect.Enabled = True
        If (cbcSelect.ListCount <= 1) Or (imAeeCode <= 0) Or (smUsedFlag <> "N") Then
            cmcErase.Enabled = False
        Else
            cmcErase.Enabled = True
        End If
    End If
End Sub
Private Function mCheckFields(ilShowMsg As Integer) As Integer
    Dim slStr As String
    
    mCheckFields = True
    slStr = Trim$(edcName.text)
    If slStr = "" Then
        If ilShowMsg Then
            MsgBox "Automation Names must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            edcName.SetFocus
        End If
        mCheckFields = False
    End If
    If (rbcState(0).Value = False) And (rbcState(1).Value = False) Then
        If ilShowMsg Then
            MsgBox "Active or Dormant must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            rbcState(0).SetFocus
        End If
        mCheckFields = False
    End If
End Function

Private Sub mInitADE(tlADE As ADE)
    tlADE.iCode = 0
    tlADE.iAeeCode = imAeeCode
    tlADE.iScheduleData = 1
    tlADE.iDate = 6
    tlADE.iDateNoChar = 8
    tlADE.iTime = 23
    tlADE.iTimeNoChar = 10
    tlADE.iAutoOff = 257
    tlADE.iData = 258
    tlADE.iSchedule = 259
    tlADE.iTrueTime = 260
    tlADE.iSourceConflict = 261
    tlADE.iSourceUnavail = 262
    tlADE.iSourceItem = 263
    tlADE.iBkupSrceUnavail = 264
    tlADE.iBkupSrceItem = 265
    tlADE.iProtSrceUnavail = 266
    tlADE.iProtSrceItem = 267
End Sub

Private Function mNameOk() As Integer
    Dim slName As String
    Dim ilLoop As Integer
    
    slName = Trim$(edcName.text)
    For ilLoop = 0 To UBound(tgCurrAEE) - 1 Step 1
        If (StrComp(slName, Trim$(tgCurrAEE(ilLoop).sName), vbTextCompare) = 0) Then
            If imAeeCode <> tgCurrAEE(ilLoop).iCode Then
               MsgBox "Name previously used", vbOKOnly + vbExclamation, "Name Used"
               edcName.SetFocus
               mNameOk = False
            End If
        End If
    Next ilLoop
    mNameOk = True
    
End Function

Private Function mCompareAFE(ilCode As Integer) As Integer
    Dim ilAFENew As Integer
    Dim ilAFEOld As Integer
    
    If ilCode > 0 Then
        For ilAFENew = LBound(tmCurrAFE) To UBound(tmCurrAFE) - 1 Step 1
            If ilCode = tmCurrAFE(ilAFENew).iCode Then
                For ilAFEOld = LBound(tgCurrAFE) To UBound(tgCurrAFE) - 1 Step 1
                    If ilCode = tgCurrAFE(ilAFEOld).iCode Then
                        'Compare fields
                        If tmCurrAFE(ilAFENew).iBus <> tgCurrAFE(ilAFEOld).iBus Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iBusControl <> tgCurrAFE(ilAFEOld).iBusControl Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iDate <> tgCurrAFE(ilAFEOld).iDate Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iEventID <> tgCurrAFE(ilAFEOld).iEventID Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iEventType <> tgCurrAFE(ilAFEOld).iEventType Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iTime <> tgCurrAFE(ilAFEOld).iTime Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iStartType <> tgCurrAFE(ilAFEOld).iStartType Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iFixedTime <> tgCurrAFE(ilAFEOld).iFixedTime Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iEndType <> tgCurrAFE(ilAFEOld).iEndType Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iDuration <> tgCurrAFE(ilAFEOld).iDuration Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iEndTime <> tgCurrAFE(ilAFEOld).iEndTime Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iMaterialType <> tgCurrAFE(ilAFEOld).iMaterialType Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iAudioName <> tgCurrAFE(ilAFEOld).iAudioName Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iAudioItemID <> tgCurrAFE(ilAFEOld).iAudioItemID Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iAudioControl <> tgCurrAFE(ilAFEOld).iAudioControl Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iBkupAudioName <> tgCurrAFE(ilAFEOld).iBkupAudioName Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iBkupAudioControl <> tgCurrAFE(ilAFEOld).iBkupAudioControl Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iProtItemID <> tgCurrAFE(ilAFEOld).iProtItemID Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iProtAudioControl <> tgCurrAFE(ilAFEOld).iProtAudioControl Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iRelay1 <> tgCurrAFE(ilAFEOld).iRelay1 Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iRelay2 <> tgCurrAFE(ilAFEOld).iRelay2 Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iFollow <> tgCurrAFE(ilAFEOld).iFollow Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iSilenceTime <> tgCurrAFE(ilAFEOld).iSilenceTime Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iSilence1 <> tgCurrAFE(ilAFEOld).iSilence1 Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iSilence2 <> tgCurrAFE(ilAFEOld).iSilence2 Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iSilence3 <> tgCurrAFE(ilAFEOld).iSilence3 Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iSilence4 <> tgCurrAFE(ilAFEOld).iSilence4 Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iStartNetcue <> tgCurrAFE(ilAFEOld).iStartNetcue Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iStopNetcue <> tgCurrAFE(ilAFEOld).iStopNetcue Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iTitle1 <> tgCurrAFE(ilAFEOld).iTitle1 Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If tmCurrAFE(ilAFENew).iTitle2 <> tgCurrAFE(ilAFEOld).iTitle2 Then
                            mCompareAFE = False
                            Exit Function
                        End If
                        If sgClientFields = "A" Then
                            If tmCurrAFE(ilAFENew).iABCFormat <> tgCurrAFE(ilAFEOld).iABCFormat Then
                                mCompareAFE = False
                                Exit Function
                            End If
                            If tmCurrAFE(ilAFENew).iABCPgmCode <> tgCurrAFE(ilAFEOld).iABCPgmCode Then
                                mCompareAFE = False
                                Exit Function
                            End If
                            If tmCurrAFE(ilAFENew).iABCXDSMode <> tgCurrAFE(ilAFEOld).iABCXDSMode Then
                                mCompareAFE = False
                                Exit Function
                            End If
                            If tmCurrAFE(ilAFENew).iABCRecordItem <> tgCurrAFE(ilAFEOld).iABCRecordItem Then
                                mCompareAFE = False
                                Exit Function
                            End If
                        End If
                        mCompareAFE = True
                        Exit Function
                    End If
                Next ilAFEOld
                mCompareAFE = True
                Exit Function
            End If
        Next ilAFENew
    Else
        mCompareAFE = True
    End If
End Function

Private Function mCompareADE(ilCode As Integer) As Integer
    Dim ilADENew As Integer
    Dim ilADEOld As Integer
    
    If ilCode > 0 Then
        For ilADENew = LBound(tmCurrADE) To UBound(tmCurrADE) - 1 Step 1
            If ilCode = tmCurrADE(ilADENew).iCode Then
                For ilADEOld = LBound(tgCurrADE) To UBound(tgCurrADE) - 1 Step 1
                    If ilCode = tgCurrADE(ilADEOld).iCode Then
                        'Compare fields
                        If tmCurrADE(ilADENew).iScheduleData <> tgCurrADE(ilADEOld).iScheduleData Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iDate <> tgCurrADE(ilADEOld).iDate Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iDateNoChar <> tgCurrADE(ilADEOld).iDateNoChar Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iTime <> tgCurrADE(ilADEOld).iTime Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iTimeNoChar <> tgCurrADE(ilADEOld).iTimeNoChar Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iAutoOff <> tgCurrADE(ilADEOld).iAutoOff Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iData <> tgCurrADE(ilADEOld).iData Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iSchedule <> tgCurrADE(ilADEOld).iSchedule Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iTrueTime <> tgCurrADE(ilADEOld).iTrueTime Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iSourceConflict <> tgCurrADE(ilADEOld).iSourceConflict Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iSourceUnavail <> tgCurrADE(ilADEOld).iSourceUnavail Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iSourceItem <> tgCurrADE(ilADEOld).iSourceItem Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iBkupSrceUnavail <> tgCurrADE(ilADEOld).iBkupSrceUnavail Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iBkupSrceItem <> tgCurrADE(ilADEOld).iBkupSrceItem Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iProtSrceUnavail <> tgCurrADE(ilADEOld).iProtSrceUnavail Then
                            mCompareADE = False
                            Exit Function
                        End If
                        If tmCurrADE(ilADENew).iProtSrceItem <> tgCurrADE(ilADEOld).iProtSrceItem Then
                            mCompareADE = False
                            Exit Function
                        End If
                        mCompareADE = True
                        Exit Function
                    End If
                Next ilADEOld
                mCompareADE = True
                Exit Function
            End If
        Next ilADENew
    Else
        mCompareADE = True
    End If
End Function
Private Sub mEEnableBox()
    Dim llColPos As Long
    Dim ilCol As Integer
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(EVENTTYPELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdExport.Row >= grdExport.FixedRows) And (grdExport.Row < grdExport.Rows) And (grdExport.Col >= 0) And (grdExport.Col < grdExport.Cols - 1) Then
        mISetShow
        lmEEnableRow = grdExport.Row
        ilCol = grdExport.Col
        If grdExport.Col >= EVENTIDINDEX Then
            grdExport.LeftCol = grdExport.LeftCol + 5
            DoEvents
        End If
        lmEEnableRow = grdExport.Row
        grdExport.Col = ilCol
        lmEEnableCol = grdExport.Col
        llColPos = 0
        For ilCol = 0 To grdExport.Col - 1 Step 1
            If grdExport.ColIsVisible(ilCol) Then
                llColPos = llColPos + grdExport.ColWidth(ilCol)
            End If
        Next ilCol
        edcExport.Move grdExport.Left + llColPos + 30, grdExport.Top + grdExport.RowPos(grdExport.Row) + 15, grdExport.ColWidth(grdExport.Col) - 30, grdExport.RowHeight(grdExport.Row) - 15
        'edcExport.Move grdExport.Left + grdExport.ColPos(grdExport.Col) + 30, grdExport.Top + grdExport.RowPos(grdExport.Row) + 15, grdExport.ColWidth(grdExport.Col) - 30, grdExport.RowHeight(grdExport.Row) - 15
        If lmEEnableRow = 3 Then
            edcExport.MaxLength = 3
        Else
            If imMaxChar(lmEEnableCol) < 10 Then
                edcExport.MaxLength = 1
            ElseIf imMaxChar(lmEEnableCol) < 100 Then
                edcExport.MaxLength = 2
            Else
                edcExport.MaxLength = 3
            End If
        End If
        edcExport.text = grdExport.text
        edcExport.Visible = True
        edcExport.SetFocus
    End If
End Sub

Private Sub mIEnableBox()
    Dim llColPos As Integer
    Dim ilCol As Integer
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(EVENTTYPELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdImport.Row >= grdImport.FixedRows) And (grdImport.Row < grdImport.Rows) And (grdImport.Col >= 0) And (grdImport.Col < grdImport.Cols - 1) Then
        mESetShow
        lmIEnableRow = grdImport.Row
        lmIEnableCol = grdImport.Col
        llColPos = 0
        For ilCol = 0 To grdImport.Col - 1 Step 1
            llColPos = llColPos + grdImport.ColWidth(ilCol)
        Next ilCol
        edcImport.Move grdImport.Left + llColPos + 30, grdImport.Top + grdImport.RowPos(grdImport.Row) + 15, grdImport.ColWidth(grdImport.Col) - 30, grdImport.RowHeight(grdImport.Row) - 15
        'edcImport.Move grdImport.Left + grdImport.ColPos(grdImport.Col) + 30, grdImport.Top + grdImport.RowPos(grdImport.Row) + 15, grdImport.ColWidth(grdImport.Col) - 30, grdImport.RowHeight(grdImport.Row) - 15
        edcImport.MaxLength = 3
        edcImport.text = grdImport.text
        edcImport.Visible = True
        edcImport.SetFocus
    End If
End Sub
Private Sub mESetShow()
    edcExport.Visible = False
End Sub

Private Sub mISetShow()
    edcImport.Visible = False
End Sub

Private Sub mSetMaxChar()
    imMaxChar(BUSNAMEINDEX) = 8
    imMaxChar(BUSCTRLINDEX) = 1
    imMaxChar(EVENTTYPEINDEX) = 1
    imMaxChar(TIMEINDEX) = 10
    imMaxChar(STARTTYPEINDEX) = 3
    imMaxChar(FIXEDINDEX) = 1
    imMaxChar(ENDTYPEINDEX) = 3
    imMaxChar(DURATIONINDEX) = 10
    imMaxChar(ENDTIMEINDEX) = 10
    imMaxChar(MATERIALINDEX) = 3
    imMaxChar(AUDIONAMEINDEX) = 8
    imMaxChar(AUDIOITEMIDINDEX) = 32
    imMaxChar(AUDIOISCIINDEX) = 20
    imMaxChar(AUDIOCTRLINDEX) = 1
    imMaxChar(BACKUPNAMEINDEX) = 8
    imMaxChar(BACKUPCTRLINDEX) = 1
    imMaxChar(PROTNAMEINDEX) = 8
    imMaxChar(PROTITEMIDINDEX) = 32
    imMaxChar(PROTISCIINDEX) = 20
    imMaxChar(PROTCTRLINDEX) = 1
    imMaxChar(RELAY1INDEX) = 8
    imMaxChar(RELAY2INDEX) = 8
    imMaxChar(FOLLOWINDEX) = 19
    imMaxChar(SILENCETIMEINDEX) = 5
    imMaxChar(SILENCE1INDEX) = 1
    imMaxChar(SILENCE2INDEX) = 1
    imMaxChar(SILENCE3INDEX) = 1
    imMaxChar(SILENCE4INDEX) = 1
    imMaxChar(NETCUE1INDEX) = 3
    imMaxChar(NETCUE2INDEX) = 3
    imMaxChar(TITLE1INDEX) = 66
    imMaxChar(TITLE2INDEX) = 90
    imMaxChar(DATEINDEX) = 8
    imMaxChar(EVENTIDINDEX) = 20
    imMaxChar(ABCFORMATINDEX) = 1
    imMaxChar(ABCPGMCODEINDEX) = 25
    imMaxChar(ABCXDSMODEINDEX) = 2
    imMaxChar(ABCRECORDITEMINDEX) = 5
    
End Sub

Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim ilCompare As Integer
    Dim ilCode As Integer
    
    Screen.MousePointer = vbHourglass
    If Not mCheckFields(True) Then
        Screen.MousePointer = vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        Screen.MousePointer = vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mCheckExportColumns() Then
        Screen.MousePointer = vbDefault
        MsgBox "Export Column Definition in Conflict (overlap)", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    mMoveCtrlsToRec
    mMoveADECtrlsToRec
    mMoveAFECtrlsToRec
    If imAeeCode <= 0 Then
        ilRet = gPutInsert_AEE_AutoEquip(0, tmAEE, "User Option-mSave: AEE")
    Else
        ilRet = False
        For ilLoop = LBound(tgCurrAEE) To UBound(tgCurrAEE) - 1 Step 1
            If imAeeCode = tgCurrAEE(ilLoop).iCode Then
                ilRet = mCompare(tmAEE, tgCurrAEE(ilLoop))
                Exit For
            End If
        Next ilLoop
        If ilRet Then
            ilRet = gPutUpdate_AEE_AutoEquip(0, tmAEE, "User Option-mSave: AEE")
        Else
            ilRet = gPutUpdate_AEE_AutoEquip(1, tmAEE, "User Option-mSave: AEE")
        End If
    End If
    For ilLoop = 0 To UBound(tmCurrACE) - 1 Step 1
        If imAeeCode > 0 Then
            ilCode = tmCurrACE(ilLoop).iCode
            ilCompare = mCompareACE(ilCode)
        Else
            ilCompare = True
        End If
        tmCurrACE(ilLoop).iCode = 0
        tmCurrACE(ilLoop).iAeeCode = tmAEE.iCode
        ilRet = gPutInsert_ACE_AutoContact(tmCurrACE(ilLoop), "Automation- mSave: Insert ACE")
        If Not ilCompare Then
            ilRet = gUpdateAIE(1, tmAEE.iVersion, "ACE", CLng(ilCode), CLng(tmCurrACE(ilLoop).iCode), CLng(tmAEE.iOrigAeeCode), "Automation- mSave: Insert ACE:AIE")
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tmCurrADE) - 1 Step 1
        If imAeeCode > 0 Then
            ilCode = tmCurrADE(ilLoop).iCode
            ilCompare = mCompareADE(ilCode)
        Else
            ilCompare = True
        End If
        tmCurrADE(ilLoop).iCode = 0
        tmCurrADE(ilLoop).iAeeCode = tmAEE.iCode
        ilRet = gPutInsert_ADE_AutoDataFlags(tmCurrADE(ilLoop), "Automation- mSave: Insert ADE")
        If Not ilCompare Then
            ilRet = gUpdateAIE(1, tmAEE.iVersion, "ADE", CLng(ilCode), CLng(tmCurrADE(ilLoop).iCode), CLng(tmAEE.iOrigAeeCode), "Automation- mSave: Insert ADE:AIE")
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tmCurrAFE) - 1 Step 1
        If imAeeCode > 0 Then
            ilCode = tmCurrAFE(ilLoop).iCode
            ilCompare = mCompareAFE(ilCode)
        Else
            ilCompare = True
        End If
        tmCurrAFE(ilLoop).iCode = 0
        tmCurrAFE(ilLoop).iAeeCode = tmAEE.iCode
        ilRet = gPutInsert_AFE_AutoFormat(tmCurrAFE(ilLoop), "Automation- mSave: Insert AFE")
        If Not ilCompare Then
            ilRet = gUpdateAIE(1, tmAEE.iVersion, "AFE", CLng(ilCode), CLng(tmCurrAFE(ilLoop).iCode), CLng(tmAEE.iOrigAeeCode), "Automation- mSave: Insert AFE:AIE")
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tmCurrAPE) - 1 Step 1
        If imAeeCode > 0 Then
            ilCode = tmCurrAPE(ilLoop).iCode
            ilCompare = mCompareAPE(ilCode)
        Else
            ilCompare = True
        End If
        tmCurrAPE(ilLoop).iCode = 0
        tmCurrAPE(ilLoop).iAeeCode = tmAEE.iCode
        ilRet = gPutInsert_APE_AutoPath(tmCurrAPE(ilLoop), "Automation- mSave: Insert APE")
        If Not ilCompare Then
            ilRet = gUpdateAIE(1, tmAEE.iVersion, "APE", CLng(ilCode), CLng(tmCurrAPE(ilLoop).iCode), CLng(tmAEE.iOrigAeeCode), "Automation- mSave: Insert APE:AIE")
        End If
    Next ilLoop
    
    mSave = True
    Screen.MousePointer = vbDefault
End Function

Private Function mCheckExportColumns() As Integer
    Dim ilStartCol As Integer
    Dim ilEndCol As Integer
    Dim ilCol As Integer
    Dim ilTestStartCol As Integer
    Dim ilTestEndCol As Integer
    Dim ilTestCol As Integer
    
    mCheckExportColumns = True
    For ilCol = BUSNAMEINDEX To imMaxCols Step 1
        If (Val(grdExport.TextMatrix(3, ilCol)) > 0) And (Val(grdExport.TextMatrix(4, ilCol)) > 0) Then
            ilStartCol = Val(grdExport.TextMatrix(3, ilCol))
            ilEndCol = ilStartCol + Val(grdExport.TextMatrix(4, ilCol)) - 1
            For ilTestCol = ilCol + 1 To imMaxCols Step 1
                If (Val(grdExport.TextMatrix(3, ilTestCol)) > 0) And (Val(grdExport.TextMatrix(4, ilTestCol)) > 0) Then
                    ilTestStartCol = grdExport.TextMatrix(3, ilTestCol)
                    ilTestEndCol = ilTestStartCol + grdExport.TextMatrix(4, ilTestCol) - 1
                    If (ilTestEndCol >= ilStartCol) And (ilTestStartCol <= ilEndCol) Then
                        grdExport.Row = 4
                        grdExport.Col = ilTestCol
                        grdExport.CellForeColor = vbRed
                        mCheckExportColumns = False
                    End If
                End If
            Next ilTestCol
        End If
    Next ilCol
End Function
