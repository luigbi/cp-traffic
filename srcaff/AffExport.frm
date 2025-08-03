VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmExport 
   Caption         =   "Export Center"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   Icon            =   "AffExport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   9240
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3105
      Top             =   6645
   End
   Begin VB.PictureBox pbcArial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2685
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   31
      Top             =   6660
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox lbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      ItemData        =   "AffExport.frx":08CA
      Left            =   135
      List            =   "AffExport.frx":08CC
      TabIndex        =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Timer tmcStatus 
      Interval        =   30000
      Left            =   375
      Top             =   6630
   End
   Begin VB.PictureBox pbcStatusTab 
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
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   525
      Width           =   60
   End
   Begin VB.PictureBox pbcStatusSTab 
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
      TabIndex        =   5
      Top             =   0
      Width           =   15
   End
   Begin VB.TextBox edcStatusDropdown 
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
      Left            =   6165
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   195
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer tmcClock 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1830
      Top             =   6570
   End
   Begin VB.CommandButton cmcSpec 
      Caption         =   "Specifications"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   6210
      Width           =   1335
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
      Left            =   75
      ScaleHeight     =   90
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   6300
      Width           =   15
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
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   6480
      Width           =   60
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
      Left            =   4350
      MaxLength       =   20
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   165
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcGen 
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
      Left            =   3510
      ScaleHeight     =   210
      ScaleWidth      =   375
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pbcUndo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   975
      Picture         =   "AffExport.frx":08CE
      ScaleHeight     =   300
      ScaleWidth      =   285
      TabIndex        =   29
      Top             =   6630
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Station"
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   2
      Left            =   8085
      TabIndex        =   20
      Top             =   2865
      Visible         =   0   'False
      Width           =   8775
      Begin V81Affiliate.CSI_Calendar edcExportStartDate 
         Height          =   285
         Left            =   1380
         TabIndex        =   26
         Top             =   705
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         Text            =   "4/15/2020"
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
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.TextBox txUFile 
         Height          =   300
         Left            =   870
         TabIndex        =   23
         Top             =   210
         Width           =   6030
      End
      Begin VB.CommandButton cmcUBrowse 
         Caption         =   "&Browse..."
         Height          =   300
         Left            =   7245
         TabIndex        =   22
         Top             =   210
         Width           =   1170
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStations 
         Height          =   2550
         Left            =   1410
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1260
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   4498
         _Version        =   393216
         Rows            =   4
         Cols            =   5
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollBars      =   2
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
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin V81Affiliate.CSI_Calendar edcExportEndDate 
         Height          =   285
         Left            =   4530
         TabIndex        =   28
         Top             =   705
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         Text            =   "4/15/2020"
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
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.Label lacExportEndDate 
         Caption         =   "Export End Date"
         Height          =   315
         Left            =   3135
         TabIndex        =   27
         Top             =   750
         Width           =   1530
      End
      Begin VB.Label lacExportStartDate 
         Caption         =   "Export Start Date"
         Height          =   315
         Left            =   0
         TabIndex        =   25
         Top             =   750
         Width           =   1530
      End
      Begin VB.Label lbcUFile 
         Caption         =   "Import File"
         Height          =   315
         Left            =   0
         TabIndex        =   24
         Top             =   195
         Width           =   780
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   60
      ScaleHeight     =   90
      ScaleWidth      =   15
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6270
      Width           =   15
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Alerts"
      ForeColor       =   &H80000008&
      Height          =   5190
      Index           =   1
      Left            =   8190
      TabIndex        =   15
      Top             =   2220
      Visible         =   0   'False
      Width           =   8775
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAlerts 
         Height          =   4890
         Left            =   75
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   180
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   8625
         _Version        =   393216
         Rows            =   4
         Cols            =   9
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollBars      =   2
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Selection"
      ForeColor       =   &H80000008&
      Height          =   5145
      Index           =   0
      Left            =   4260
      TabIndex        =   8
      Top             =   480
      Width           =   8895
      Begin VB.PictureBox pbcArrow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   0
         Picture         =   "AffExport.frx":1198
         ScaleHeight     =   165
         ScaleWidth      =   90
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1485
         Visible         =   0   'False
         Width           =   90
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdExport 
         Height          =   2355
         Left            =   180
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   285
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   4154
         _Version        =   393216
         Rows            =   8
         Cols            =   17
         FixedRows       =   2
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollBars      =   2
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStatus 
         Height          =   690
         Left            =   150
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4275
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   1217
         _Version        =   393216
         Rows            =   4
         Cols            =   16
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollBars      =   2
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
         _Band(0).Cols   =   16
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1260
      Top             =   6600
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6750
      FormDesignWidth =   9240
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   1140
      TabIndex        =   11
      Top             =   6210
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   6210
      Width           =   1335
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5745
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   10134
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Options"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Alerts"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmcCustom 
      Caption         =   "Custom"
      Height          =   375
      Left            =   4140
      TabIndex        =   13
      Top             =   6210
      Width           =   1335
   End
   Begin VB.CommandButton cmcCheckCopy 
      Caption         =   "Check Copy"
      Height          =   375
      Left            =   7140
      TabIndex        =   32
      Top             =   6210
      Width           =   1335
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   135
      Picture         =   "AffExport.frx":14A2
      Top             =   6195
      Width           =   480
   End
   Begin VB.Label lacProcess 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   105
      TabIndex        =   16
      Top             =   5775
      Visible         =   0   'False
      Width           =   8775
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmExport - shows certificate of performance information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text
Private imIntegralSet As Integer
Private imTabIndex As Integer
Private imVefCode As Integer
Private imShfCode As Integer
Private imAllClick As Integer
Private imNextPriority As Integer
Private bmDateError As Boolean
'Private smCntrNo As String
'Private smChfType As String
'Private tmCmmlSum() As CMMLSUM
'Private imMaxDays As Integer
'Private chfrst As ADODB.Recordset
'Private smDate As String
'Private tmCPInfo() As CPINFO
'Private tmCPDat() As DAT
'Private tmAstInfo() As ASTINFO
'Private cprst As ADODB.Recordset
'Private lstrst As ADODB.Recordset
Private imFirstTime As Integer
Private bFormWasAlreadyResized As Boolean
Private hmAst As Integer
Private tmAufView() As AUFVIEW

Private tmSplitEhtInfo() As EHTINFO
Private tmSplitEvtInfo() As EVTINFO
Private tmSplitEctInfo() As ECTINFO
Private lmStandardEhtCode() As Long
Private tmEhtStdColor() As EHTSTDCOLOR


Private smGen As String
Private smStatusPrevTip As String

'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long
Private lmEnableCol As Long
Private imCtrlVisible As Integer

Private lmStatusTopRow As Long            'Top row when cell clicked or - 1
Private lmStatusEnableRow As Long
Private lmStatusEnableCol As Long
Private imStatusCtrlVisible As Integer
Private smUserChgPriority As String

Private lm1970 As Long

Private Type LOGANDCOPYCOLOR
    lEhtCode As Long
    sLogColor As String * 1
    sCopyColor As String * 1
End Type
Private tmLogAndCopyColor() As LOGANDCOPYCOLOR


Private rst_Eht As ADODB.Recordset
Private rst_Evt As ADODB.Recordset
Private rst_Ect As ADODB.Recordset
Private rst_Eqt As ADODB.Recordset
Private rst_Ust As ADODB.Recordset
Private rst_Lst As ADODB.Recordset

Private Const GENINDEX = 0
Private Const EXPORTTYPEINDEX = 1
Private Const EXPORTNAMEINDEX = 2
Private Const VEHICLEINDEX = 3
Private Const LOGSTATUSINDEX = 4
Private Const COPYSTATUSINDEX = 5
Private Const WORKDATEINDEX = 6
Private Const LASTDATEINDEX = 7
Private Const LEADTIMEINDEX = 8
Private Const CYCLEINDEX = 9
Private Const STARTDATEINDEX = 10
Private Const ENDDATEINDEX = 11
Private Const CLOSEINDEX = 12
Private Const EHTINFOINDEX = 13
Private Const EHTTYPECHARINDEX = 14
Private Const SORTINDEX = 15
Private Const EHTCODEINDEX = 16

Private Const SEXPORTTYPEINDEX = 0
Private Const SEXPORTNAMEINDEX = 1
Private Const SVEHICLEINDEX = 2
Private Const SUSERINDEX = 3
Private Const SPRIORITYINDEX = 4
Private Const SSTATUSINDEX = 5
Private Const SEXPORTINFOINDEX = 6
Private Const STIMEREQUESTINDEX = 7
Private Const STIMESTARTEDINDEX = 8
Private Const STIMEENDINDEX = 9
Private Const SCLOSEINDEX = 10
Private Const SSORTINDEX = 11
Private Const SUSTCODEINDEX = 12
Private Const SUSERNAMEINDEX = 13
Private Const SRESULTFILEINDEX = 14
Private Const SEQTCODEINDEX = 15

Private Const AEXPORTINDEX = 0
Private Const AACTIONINDEX = 1
Private Const AREASONINDEX = 2
Private Const ACREATIONDATEINDEX = 3
Private Const ACREATIONTIMEINDEX = 4
Private Const AVEHICLEINDEX = 5
Private Const ADATEINDEX = 6
Private Const ADELETEINDEX = 7
Private Const AAUFCODEINDEX = 8

Private Const IEXPORTINDEX = 0
Private Const ISTATIONINDEX = 1
Private Const ISTARTDATEINDEX = 2
Private Const IENDDATEINDEX = 3
Private Const ISHTTCODEINDEX = 4

Private Sub mClearGrid(grdCtrl As MSHFlexGrid)
    
    gGrid_Clear grdCtrl, True
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdExport.ColWidth(EHTCODEINDEX) = 0
    grdExport.ColWidth(EHTINFOINDEX) = 0
    grdExport.ColWidth(SORTINDEX) = 0
    grdExport.ColWidth(EHTTYPECHARINDEX) = 0
    grdExport.ColWidth(CLOSEINDEX) = grdExport.Width * 0.02
    grdExport.ColWidth(GENINDEX) = grdExport.Width * 0.04
    grdExport.ColWidth(EXPORTTYPEINDEX) = grdExport.Width * 0.09    '0.15
    grdExport.ColWidth(VEHICLEINDEX) = grdExport.Width * 0.09    '0.15
    grdExport.ColWidth(LOGSTATUSINDEX) = grdExport.Width * 0.02  '0.13
    grdExport.ColWidth(COPYSTATUSINDEX) = grdExport.Width * 0.02  '0.13
    grdExport.ColWidth(WORKDATEINDEX) = grdExport.Width * 0.09  '0.13
    grdExport.ColWidth(LASTDATEINDEX) = grdExport.Width * 0.09
    grdExport.ColWidth(LEADTIMEINDEX) = grdExport.Width * 0.06  '0.11
    grdExport.ColWidth(CYCLEINDEX) = grdExport.Width * 0.06
    grdExport.ColWidth(STARTDATEINDEX) = grdExport.Width * 0.09
    grdExport.ColWidth(ENDDATEINDEX) = grdExport.Width * 0.09

    grdExport.ColWidth(EXPORTNAMEINDEX) = grdExport.Width - GRIDSCROLLWIDTH - 15
    For ilCol = GENINDEX To CLOSEINDEX Step 1
        If ilCol <> EXPORTNAMEINDEX Then
            grdExport.ColWidth(EXPORTNAMEINDEX) = grdExport.ColWidth(EXPORTNAMEINDEX) - grdExport.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdExport
    
    grdStatus.ColWidth(SSORTINDEX) = 0
    grdStatus.ColWidth(SUSTCODEINDEX) = 0
    grdStatus.ColWidth(SUSERNAMEINDEX) = 0
    grdStatus.ColWidth(SRESULTFILEINDEX) = 0
    grdStatus.ColWidth(SEQTCODEINDEX) = 0
    grdStatus.ColWidth(SEXPORTTYPEINDEX) = grdStatus.Width * 0.09
    grdStatus.ColWidth(SVEHICLEINDEX) = grdStatus.Width * 0.09    '0.15
    grdStatus.ColWidth(SUSERINDEX) = grdStatus.Width * 0.06
    grdStatus.ColWidth(STIMEREQUESTINDEX) = grdStatus.Width * 0.11
    grdStatus.ColWidth(SPRIORITYINDEX) = grdStatus.Width * 0.07
    grdStatus.ColWidth(SSTATUSINDEX) = grdStatus.Width * 0.09
    grdStatus.ColWidth(SEXPORTINFOINDEX) = grdStatus.Width * 0.09
    grdStatus.ColWidth(STIMESTARTEDINDEX) = grdStatus.Width * 0.09
    grdStatus.ColWidth(STIMEENDINDEX) = grdStatus.Width * 0.11
    grdStatus.ColWidth(SCLOSEINDEX) = grdStatus.Width * 0.02

    grdStatus.ColWidth(SEXPORTNAMEINDEX) = grdStatus.Width - GRIDSCROLLWIDTH - 15
    For ilCol = SEXPORTTYPEINDEX To SCLOSEINDEX Step 1
        If ilCol <> SEXPORTNAMEINDEX Then
            grdStatus.ColWidth(SEXPORTNAMEINDEX) = grdStatus.ColWidth(SEXPORTNAMEINDEX) - grdStatus.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdStatus

    grdAlerts.ColWidth(AAUFCODEINDEX) = 0
    grdAlerts.ColWidth(AEXPORTINDEX) = 0   'grdAlerts.Width * 0.06
    grdAlerts.ColWidth(AACTIONINDEX) = grdAlerts.Width * 0.12
    grdAlerts.ColWidth(ACREATIONDATEINDEX) = grdAlerts.Width * 0.11
    grdAlerts.ColWidth(ACREATIONTIMEINDEX) = grdAlerts.Width * 0.12
    grdAlerts.ColWidth(ADATEINDEX) = grdAlerts.Width * 0.09
    grdAlerts.ColWidth(AREASONINDEX) = grdAlerts.Width * 0.16
    grdAlerts.ColWidth(ADELETEINDEX) = grdAlerts.Width * 0.07

    grdAlerts.ColWidth(AVEHICLEINDEX) = grdAlerts.Width - GRIDSCROLLWIDTH - 15
    For ilCol = AEXPORTINDEX To ADELETEINDEX Step 1
        If ilCol <> AVEHICLEINDEX Then
            grdAlerts.ColWidth(AVEHICLEINDEX) = grdAlerts.ColWidth(AVEHICLEINDEX) - grdAlerts.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdAlerts

    grdStations.ColWidth(ISHTTCODEINDEX) = 0
    grdStations.ColWidth(IEXPORTINDEX) = grdStations.Width * 0.12
    grdStations.ColWidth(ISTARTDATEINDEX) = grdStations.Width * 0.2
    grdStations.ColWidth(IENDDATEINDEX) = grdStations.Width * 0.2

    grdStations.ColWidth(ISTATIONINDEX) = grdStations.Width - GRIDSCROLLWIDTH - 15
    For ilCol = IEXPORTINDEX To IENDDATEINDEX Step 1
        If ilCol <> ISTATIONINDEX Then
            grdStations.ColWidth(ISTATIONINDEX) = grdStations.ColWidth(ISTATIONINDEX) - grdStations.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdStations

End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdExport.TextMatrix(0, GENINDEX) = "Gen"
    grdExport.TextMatrix(1, GENINDEX) = ""
    grdExport.TextMatrix(0, EXPORTTYPEINDEX) = "Export"
    grdExport.TextMatrix(1, EXPORTTYPEINDEX) = "Type"
    grdExport.TextMatrix(0, EXPORTNAMEINDEX) = "Export"
    grdExport.TextMatrix(1, EXPORTNAMEINDEX) = "Name"
    grdExport.TextMatrix(0, VEHICLEINDEX) = "Vehicle"
    grdExport.TextMatrix(1, VEHICLEINDEX) = "List"
    grdExport.TextMatrix(0, LOGSTATUSINDEX) = "L"
    grdExport.TextMatrix(0, COPYSTATUSINDEX) = "C"
    grdExport.TextMatrix(0, WORKDATEINDEX) = "Working"
    grdExport.TextMatrix(1, WORKDATEINDEX) = "Date"
    grdExport.TextMatrix(0, LASTDATEINDEX) = "Last Date"
    grdExport.TextMatrix(1, LASTDATEINDEX) = "Exported"
    grdExport.TextMatrix(0, LEADTIMEINDEX) = "Lead"
    grdExport.TextMatrix(1, LEADTIMEINDEX) = "Time"
    grdExport.TextMatrix(0, CYCLEINDEX) = "Cycle"
    grdExport.TextMatrix(0, STARTDATEINDEX) = "Export"
    grdExport.TextMatrix(1, STARTDATEINDEX) = "Start Date"
    grdExport.TextMatrix(0, ENDDATEINDEX) = "Export"
    grdExport.TextMatrix(1, ENDDATEINDEX) = "End Date"

    grdStatus.TextMatrix(0, SEXPORTTYPEINDEX) = "Type"
    grdStatus.TextMatrix(0, SEXPORTNAMEINDEX) = "Name"
    grdStatus.TextMatrix(0, SVEHICLEINDEX) = "Vehicle"
    grdStatus.TextMatrix(0, SUSERINDEX) = "User"
    grdStatus.TextMatrix(0, STIMEREQUESTINDEX) = "Requested"
    grdStatus.TextMatrix(0, SPRIORITYINDEX) = "Priority"
    grdStatus.TextMatrix(0, SSTATUSINDEX) = "Status"
    grdStatus.TextMatrix(0, SEXPORTINFOINDEX) = "Export"
    grdStatus.TextMatrix(0, STIMESTARTEDINDEX) = "Start"
    grdStatus.TextMatrix(0, STIMEENDINDEX) = "Completed"

    grdAlerts.TextMatrix(0, AEXPORTINDEX) = "Export"
    grdAlerts.TextMatrix(0, AACTIONINDEX) = "Action Req"
    grdAlerts.TextMatrix(0, ACREATIONDATEINDEX) = "Date Created"
    grdAlerts.TextMatrix(0, ACREATIONTIMEINDEX) = "Time Created"
    grdAlerts.TextMatrix(0, AVEHICLEINDEX) = "Vehicle Name"
    grdAlerts.TextMatrix(0, ADATEINDEX) = "Date"
    grdAlerts.TextMatrix(0, AREASONINDEX) = "Reason"
    grdAlerts.TextMatrix(0, ADELETEINDEX) = "Delete"

    grdStations.TextMatrix(0, IEXPORTINDEX) = "Export"
    grdStations.TextMatrix(0, ISTATIONINDEX) = "Affiliate"
    grdStations.TextMatrix(0, ISTARTDATEINDEX) = "Start Date"
    grdStations.TextMatrix(0, IENDDATEINDEX) = "End Date"

End Sub

Private Sub cmcCheckCopy_Click()
    mCheckCopy
End Sub

Private Sub cmcCheckCopy_GotFocus()
    mSetShow
    mStatusSetShow
End Sub

Private Sub cmcCustom_Click()
    igModelType = 2
    frmModel.Show vbModal
    If igModelReturn Then
        igExportSource = 1
        igExportReturn = 0
        Select Case Chr(lgModelFromCode)
            Case "1"    'Marketron
                igExportTypeNumber = 1
                FrmExportMarketron.Show vbModal
            Case "2"    'Univision
                igExportTypeNumber = 2
                frmExportSchdSpot.Show vbModal
            Case "3"    'Web
                igExportTypeNumber = 3
                sgWebExport = "B"
                frmWebExportSchdSpot.Show vbModal
            Case "C"    'Clearance and Compensation
                igExportTypeNumber = 6
                frmExportCnCSpots.Show vbModal
            Case "D"    'IDC
                igExportTypeNumber = 7
                FrmExportIDC.Show vbModal
            Case "I"    'ISCI
                igExportTypeNumber = 8
                frmExportISCI.Show vbModal
            Case "R"    'ISCI Cross Reference
                igExportTypeNumber = 9
                frmExportISCIXRef.Show vbModal
            Case "4"    'RCS 4
                igExportTypeNumber = 4
                igRCSExportBy = 4
                frmExportRCS.Show vbModal
            Case "5"    'RCS 5
                igExportTypeNumber = 5
                igRCSExportBy = 5
                frmExportRCS.Show vbModal
            Case "S"    'StarGuide
                igExportTypeNumber = 10
                frmExportStarGuide.Show vbModal
            Case "W"    'Wegener
                igExportTypeNumber = 11
                FrmExportWegener.Show vbModal
            Case "X"    'X-Digital
                igExportTypeNumber = 12
                FrmExportXDigital.Show vbModal
            Case "P"
                igExportTypeNumber = 13
                FrmExportiPump.Show vbModal
        End Select
        grdStatus.Redraw = False
        gSetMousePointer grdExport, grdStatus, vbHourglass
        gSetMousePointer grdAlerts, grdStations, vbHourglass
        mStatusPopulate
        mSetStatusGridColor
        gSetMousePointer grdExport, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        grdStatus.Redraw = True
    End If
End Sub

Private Sub cmcCustom_GotFocus()
    mSetShow
    mStatusSetShow
End Sub

Private Sub cmcSpec_Click()
    frmExportSpec.Show vbModal
    gSetMousePointer grdExport, grdStatus, vbHourglass
    gSetMousePointer grdAlerts, grdStations, vbHourglass
    grdExport.Redraw = False
    grdStatus.Redraw = False
    mExportPopulate
    mCheckLogs
    mStatusPopulate
    mSetExportGridColors
    mSetStatusGridColor
    gSetMousePointer grdExport, grdStatus, vbDefault
    gSetMousePointer grdAlerts, grdStations, vbDefault
    grdExport.Redraw = True
    grdStatus.Redraw = True
    If bmDateError Then
        MsgBox "Red in Gen field indicates the date range crosses Sunday and must be fixed in Specifications", vbCritical + vbOKOnly
    End If
End Sub

Private Sub cmcSpec_GotFocus()
    mSetShow
    mStatusSetShow
End Sub

Private Sub cmdCancel_Click()
    Unload frmExport
End Sub

Private Sub cmdCancel_GotFocus()
    mSetShow
    mStatusSetShow
End Sub

Private Sub cmdGenerate_Click()
    Dim llRow As Long
    Dim iRet As Integer
    Dim ilPriority As Integer
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llEqtCode As Long
    Dim ilNextPriority As Integer
    Dim ilMinPriority As Integer
    Dim ilMaxPriority As Integer
    Dim llIndex As Long
    Dim llEhtCode As Long
    Dim llEct As Long
    Dim llEctCode As Long
    Dim llEvt As Long
    Dim llEvtCode As Long
    Dim slLogColor As String
    Dim slCopyColor As String
    
    On Error GoTo ErrHand
    gSetMousePointer grdExport, grdStatus, vbHourglass
    gSetMousePointer grdAlerts, grdStations, vbHourglass
    
    ilMinPriority = -1
    ilMaxPriority = -1
    For llRow = grdExport.FixedRows To grdExport.Rows - 1 Step 1
        If grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "" Then
            grdExport.Row = llRow
            grdExport.Col = GENINDEX
            If grdExport.CellFontName = "Arial" Then
                If ilMinPriority = -1 Then
                    ilMinPriority = Val(grdExport.TextMatrix(llRow, GENINDEX))
                    ilMaxPriority = Val(grdExport.TextMatrix(llRow, GENINDEX))
                Else
                    If Val(grdExport.TextMatrix(llRow, GENINDEX)) < ilMinPriority Then
                        ilMinPriority = Val(grdExport.TextMatrix(llRow, GENINDEX))
                    End If
                    If Val(grdExport.TextMatrix(llRow, GENINDEX)) > ilMaxPriority Then
                        ilMaxPriority = Val(grdExport.TextMatrix(llRow, GENINDEX))
                    End If
                End If
            End If
        End If
    Next llRow
    
    If ilMinPriority = -1 Then
        gSetMousePointer grdExport, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        MsgBox "Priority must be entered into the 'Gen' column prior to pressing the Generate Button", vbExclamation + vbOKOnly, "Warning"
        Exit Sub
    End If
    'Save splits or update splits
    For llRow = grdExport.FixedRows To grdExport.Rows - 1 Step 1
        If grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "" Then
            If grdExport.TextMatrix(llRow, EHTINFOINDEX) <> "" Then
                llIndex = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
                If tmSplitEhtInfo(llIndex).lEhtCode > 0 Then
                    'Remove
                    mRemoveEht tmSplitEhtInfo(llIndex).lEhtCode
                End If
                If Not tmSplitEhtInfo(llIndex).blRemoved Then
                    'Add
                    llEhtCode = mAddEht(llIndex)
                    tmSplitEhtInfo(llIndex).lEhtCode = llEhtCode
                    grdExport.Row = llRow
                    grdExport.Col = LOGSTATUSINDEX
                    slLogColor = "N"
                    If grdExport.CellBackColor = MIDGREENCOLOR Then
                        slLogColor = "G"
                    ElseIf grdExport.CellBackColor = vbRed Then
                        slLogColor = "R"
                    End If
                    grdExport.Col = COPYSTATUSINDEX
                    slCopyColor = "N"
                    If grdExport.CellBackColor = MIDGREENCOLOR Then
                        slCopyColor = "G"
                    ElseIf grdExport.CellBackColor = vbRed Then
                        slCopyColor = "R"
                    End If
                    tmLogAndCopyColor(UBound(tmLogAndCopyColor)).lEhtCode = llEhtCode
                    tmLogAndCopyColor(UBound(tmLogAndCopyColor)).sLogColor = slLogColor
                    tmLogAndCopyColor(UBound(tmLogAndCopyColor)).sCopyColor = slCopyColor
                    ReDim Preserve tmLogAndCopyColor(0 To UBound(tmLogAndCopyColor) + 1) As LOGANDCOPYCOLOR
                    
                    If tmSplitEhtInfo(llIndex).lEhtCode > 0 Then
                        llEct = tmSplitEhtInfo(llIndex).lFirstEct
                        Do While llEct <> -1
                            llEctCode = mAddEct(llEhtCode, tmSplitEctInfo(llEct).sLogType, tmSplitEctInfo(llEct).sFieldType, tmSplitEctInfo(llEct).sFieldName, tmSplitEctInfo(llEct).lFieldValue, tmSplitEctInfo(llEct).sFieldString)
                            llEct = tmSplitEctInfo(llEct).lNextEct
                        Loop
                        llEvt = tmSplitEhtInfo(llIndex).lFirstEvt
                        Do While llEvt <> -1
                            llEvtCode = mAddEvt(llEhtCode, tmSplitEvtInfo(llEvt).iVefCode)
                            llEvt = tmSplitEvtInfo(llEvt).lNextEvt
                        Loop
                    End If
                Else
                End If
            End If
        End If
    Next llRow
    
    SQLQuery = "SELECT max(eqtPriority) FROM eqt_Export_Queue WHERE eqtStatus = 'P' or eqtStatus = 'R'"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ilNextPriority = 1
    Else
        If Not rst.EOF Then
            ilNextPriority = rst(0).Value + 1
        Else
            ilNextPriority = 1
        End If
    End If
    'ilNextPriority = ilNextPriority + ilMinPriority - 1
    For ilPriority = ilMinPriority To ilMaxPriority Step 1
        For llRow = grdExport.FixedRows To grdExport.Rows - 1 Step 1
            If grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "" Then
                If Val(grdExport.TextMatrix(llRow, GENINDEX)) = ilPriority Then
                    grdExport.Row = llRow
                    grdExport.Col = GENINDEX
                    If grdExport.CellFontName = "Arial" Then
                        llEhtCode = Val(grdExport.TextMatrix(llRow, EHTCODEINDEX))
                        If grdExport.TextMatrix(llRow, EHTINFOINDEX) <> "" Then
                            llIndex = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
                            If tmSplitEhtInfo(llIndex).lEhtCode > 0 Then
                                llEhtCode = tmSplitEhtInfo(llIndex).lEhtCode
                            End If
                        End If
                        slDateTime = gNow()
                        slNowDate = Format$(slDateTime, "m/d/yy")
                        slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
                        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & llEhtCode
                        Set rst_Eht = gSQLSelectCall(SQLQuery)
                        If Not rst_Eht.EOF Then
                            SQLQuery = "SELECT Count(evtCode) FROM evt_Export_Vehicles WHERE evtEhtCode = " & rst_Eht!ehtCode
                            Set rst_Evt = gSQLSelectCall(SQLQuery)
                            ilPriority = Val(grdExport.TextMatrix(llRow, GENINDEX))
                            SQLQuery = "Insert Into eqt_Export_Queue ( "
                            SQLQuery = SQLQuery & "eqtCode, "
                            SQLQuery = SQLQuery & "eqtEhtCode, "
                            SQLQuery = SQLQuery & "eqtPriority, "
                            SQLQuery = SQLQuery & "eqtDateEntered, "
                            SQLQuery = SQLQuery & "eqtTimeEntered, "
                            SQLQuery = SQLQuery & "eqtStatus, "
                            SQLQuery = SQLQuery & "eqtDateStarted, "
                            SQLQuery = SQLQuery & "eqtTimeStarted, "
                            SQLQuery = SQLQuery & "eqtDateCompleted, "
                            SQLQuery = SQLQuery & "eqtTimeCompleted, "
                            SQLQuery = SQLQuery & "eqtUstCode, "
                            SQLQuery = SQLQuery & "eqtResultFile, "
                            SQLQuery = SQLQuery & "eqtType, "
                            SQLQuery = SQLQuery & "eqtStartDate, "
                            SQLQuery = SQLQuery & "eqtNumberDays, "
                            SQLQuery = SQLQuery & "eqtEndDate, "
                            SQLQuery = SQLQuery & "eqtProcesingVefCode, "
                            SQLQuery = SQLQuery & "eqtToBeProcessed, "
                            SQLQuery = SQLQuery & "eqtBeenProcessed, "
                            SQLQuery = SQLQuery & "eqtUnused "
                            SQLQuery = SQLQuery & ") "
                            SQLQuery = SQLQuery & "Values ( "
                            SQLQuery = SQLQuery & "Replace" & ", "
                            SQLQuery = SQLQuery & rst_Eht!ehtCode & ", "
                            SQLQuery = SQLQuery & ilNextPriority & ", "
                            SQLQuery = SQLQuery & "'" & Format$(slNowDate, sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "'" & Format$(slNowTime, sgSQLTimeForm) & "', "
                            SQLQuery = SQLQuery & "'" & "R" & "', "
                            SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
                            SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
                            SQLQuery = SQLQuery & igUstCode & ", "
                            SQLQuery = SQLQuery & "'" & "" & "', "
                            SQLQuery = SQLQuery & "'" & rst_Eht!ehtExportType & "', "
                            SQLQuery = SQLQuery & "'" & Format$(grdExport.TextMatrix(llRow, STARTDATEINDEX), sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & Val(grdExport.TextMatrix(llRow, CYCLEINDEX)) & ", "
                            SQLQuery = SQLQuery & "'" & Format$(grdExport.TextMatrix(llRow, ENDDATEINDEX), sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & 0 & ", "
                            If Not rst_Evt.EOF Then
                                SQLQuery = SQLQuery & rst_Evt(0).Value & ", "
                            Else
                                SQLQuery = SQLQuery & 0 & ", "
                            End If
                            SQLQuery = SQLQuery & 0 & ", "
                            SQLQuery = SQLQuery & "'" & "" & "' "
                            SQLQuery = SQLQuery & ") "
                            llEqtCode = gInsertAndReturnCode(SQLQuery, "eqt_export_queue", "eqtCode", "Replace")
                            If llEqtCode <= 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand1:
                                gSetMousePointer grdExport, grdStatus, vbDefault
                                gSetMousePointer grdAlerts, grdStations, vbDefault
                                gHandleError "AffErrorLog.txt", "Export-cmdGenerate_Click"
                                Exit Sub
                            End If
                            On Error GoTo ErrHand
                            ilNextPriority = ilNextPriority + 1
                        End If
                    End If
                End If
            End If
        Next llRow
    Next ilPriority
    
    'Remove Export Queue that are a week old
    gSetMousePointer grdExport, grdStatus, vbHourglass
    gSetMousePointer grdAlerts, grdStations, vbHourglass
    grdExport.Redraw = False
    grdStatus.Redraw = False
    mRemoveOldEqt
    mGetAltered
    mExportPopulate
    mCheckLogs
    mStatusPopulate
    mSetExportGridColors
    mSetStatusGridColor
    lacProcess.Caption = ""
    cmdCancel.Caption = "&Done"
    gSetMousePointer grdExport, grdStatus, vbDefault
    gSetMousePointer grdAlerts, grdStations, vbDefault
    grdExport.Redraw = True
    grdStatus.Redraw = True
    If bmDateError Then
        MsgBox "Red in Gen field indicates the date range crosses Sunday and must be fixed in Specifications", vbCritical + vbOKOnly
    End If
    Exit Sub
ErrHand:
    gSetMousePointer grdExport, grdStatus, vbDefault
    gSetMousePointer grdAlerts, grdStations, vbDefault
    gHandleError "AffErrorLog.txt", "Export-mGenerate"
    Exit Sub
'ErrHand1:
'    gSetMousePointer grdExport, grdStatus, vbDefault
'    gSetMousePointer grdAlerts, grdStations, vbDefault
'    gHandleError "AffErrorLog.txt", "Export-mGenerate"
'    Return
End Sub

Private Sub cmdGenerate_GotFocus()
    mSetShow
    mStatusSetShow
End Sub

Private Sub edcDropdown_Change()
    Select Case lmEnableCol
        Case GENINDEX
        Case LEADTIMEINDEX
        Case CYCLEINDEX
        Case STARTDATEINDEX
        Case ENDDATEINDEX
    End Select
End Sub

Private Sub edcDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropdown_KeyPress(KeyAscii As Integer)
    Select Case lmEnableCol
        Case GENINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
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
        Case STARTDATEINDEX
        Case ENDDATEINDEX
    End Select
End Sub

Private Sub edcStatusDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        bgExportVisible = True
        gSetMousePointer grdExport, grdStatus, vbHourglass
        gSetMousePointer grdAlerts, grdStations, vbHourglass
        grdExport.Redraw = False
        grdStatus.Redraw = False
        pbcClickFocus.Top = -pbcClickFocus.Height - 120
        pbcTab.Top = pbcClickFocus.Top
        pbcSTab.Top = pbcClickFocus.Top
        pbcStatusTab.Top = pbcClickFocus.Top
        pbcStatusSTab.Top = pbcClickFocus.Top
        'frmExport.BackColor = vbYellow
        frcTab(0).Move TabStrip1.Left + TabStrip1.ClientLeft, TabStrip1.ClientTop + 120, TabStrip1.ClientWidth - TabStrip1.Left, TabStrip1.ClientHeight - 120
        frcTab(1).Move TabStrip1.Left + TabStrip1.ClientLeft, TabStrip1.ClientTop + 120, TabStrip1.ClientWidth - TabStrip1.Left, TabStrip1.ClientHeight - 120
        frcTab(2).Move TabStrip1.Left + TabStrip1.ClientLeft, TabStrip1.ClientTop + 120, TabStrip1.ClientWidth - TabStrip1.Left, TabStrip1.ClientHeight - 120
        mSetGridColumns
        mSetGridTitles

        mGetAltered
        'gGrid_IntegralHeight grdExport
        grdExport.Move 0, 0
        grdExport.Height = frcTab(0).Height / 2
        gGrid_IntegralHeight grdExport
        gGrid_FillWithRows grdExport
        grdExport.Height = grdExport.Height + 30
        mExportPopulate
        mCheckLogs
        mSetExportGridColors
        grdStatus.Move grdExport.Left, grdExport.Top + grdExport.Height + 120, grdStatus.Width, frcTab(0).Height - (grdExport.Top + grdExport.Height + 120)
        gGrid_IntegralHeight grdStatus
        gGrid_FillWithRows grdStatus
        grdStatus.Height = grdStatus.Height + 30
        mStatusPopulate
        mSetStatusGridColor
        grdExport.Redraw = True
        grdStatus.Redraw = True
        If bmDateError Then
            MsgBox "Red in Gen field indicates the date range crosses Sunday and must be fixed in Specifications", vbCritical + vbOKOnly
        End If

        grdAlerts.Move 0, 0, grdAlerts.Width, frcTab(1).Height - 120
        gGrid_IntegralHeight grdAlerts
        gGrid_FillWithRows grdAlerts
        grdAlerts.Height = grdAlerts.Height + 30
        grdAlerts.Redraw = False
        mAlertPopulate
        
        gGrid_IntegralHeight grdStations
        gGrid_FillWithRows grdStations
        pbcGen.Font = "Monotype Sorts"
        gSetMousePointer grdExport, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        grdAlerts.Redraw = True
        imFirstTime = False
    End If
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Visible = False
    Me.Width = Screen.Width / 1.15
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmExport
    gCenterForm frmExport
End Sub

Private Sub Form_Load()
    Dim ilRet As Integer
   
    On Error GoTo ErrHand
   
    imAllClick = False
    bFormWasAlreadyResized = False
    
    imIntegralSet = False
    imTabIndex = 1
    imFirstTime = True
    
    imCtrlVisible = False
    imFromArrow = False
    lmTopRow = -1
    lmEnableRow = -1
    lm1970 = gDateValue("1/1/1970")
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    ReDim tmSplitEhtInfo(0 To 0) As EHTINFO
    ReDim tmSplitEvtInfo(0 To 0) As EVTINFO
    ReDim tmSplitEctInfo(0 To 0) As ECTINFO
    ReDim tmEhtStdColor(0 To 0) As EHTSTDCOLOR
    ReDim tmLogAndCopyColor(0 To 0) As LOGANDCOPYCOLOR

    'smUserChgPriority = "N"
    'SQLQuery = "SELECT ustChgExptPriority FROM Ust Where ustCode = " & igUstCode
    'Set rst_Ust = gSQLSelectCall(SQLQuery)
    'If Not rst_Ust.EOF Then
    '    smUserChgPriority = Trim$(rst_Ust!ustChgExptPriority)
    'End If
    smUserChgPriority = sgChgExptPriority
    If sgUstWin(14) = "V" Then
        cmdGenerate.Enabled = False
        cmcCustom.Enabled = False
        cmcSpec.Enabled = False
    End If
    
    tmcStart.Enabled = True
    Exit Sub

ErrHand:
    gSetMousePointer grdExport, grdStatus, vbDefault
    gSetMousePointer grdAlerts, grdStations, vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in CP-Form Load: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
End Sub


Private Sub Form_Resize()
    If bFormWasAlreadyResized Then
        Exit Sub
    End If
    bFormWasAlreadyResized = True
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    'TabStrip1.Left = frcDest.Left
    'TabStrip1.Height = TabStrip1.ClientTop - TabStrip1.Top + (10 * frcTab(1).Height) / 9
    'TabStrip1.Width = frcDest.Width
    'frcTab(0).Move TabStrip1.ClientLeft, TabStrip1.ClientTop
    'frcTab(1).Move TabStrip1.ClientLeft, TabStrip1.ClientTop
    frcTab(0).BorderStyle = 0
    frcTab(1).BorderStyle = 0
    frcTab(2).BorderStyle = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    bgExportVisible = False
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    rst_Eht.Close
    rst_Evt.Close
    rst_Ect.Close
    rst_Eqt.Close
    rst_Ust.Close
    rst_Lst.Close
    
    Erase tmAufView
    Erase tmSplitEhtInfo
    Erase tmSplitEvtInfo
    Erase tmSplitEctInfo
    Erase lmStandardEhtCode
    Erase tmEhtStdColor
    Erase tmLogAndCopyColor
    
    Set frmExport = Nothing
End Sub

Private Sub grdExport_EnterCell()
    mSetShow
    mStatusSetShow
End Sub

Private Sub grdExport_GotFocus()
    If grdExport.Col >= grdExport.Cols Then
        Exit Sub
    End If
    'grdExport_Click
End Sub

Private Sub grdExport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdExport.TopRow
    grdExport.Redraw = False
End Sub

Private Sub grdExport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    Dim llEhtInfo As Long
    Dim llEvtInfo As Long
    Dim llEctInfo As Long
    Dim llNext As Long
    Dim ilCount As Integer
    Dim llSvNext As Long
    Dim llVef As Long
    Dim llCopyEct As Long
    Dim llEht As Long
    Dim llStdEhtCode As Long
    Dim llEvtNext As Long
    Dim llEctNext As Long
    Dim ilPass As Integer
    Dim ilSave As Integer
    Dim slLogColor As String
    Dim slCopyColor As String
    Dim slGenFont As String
    Dim slGen As String
    
    On Error GoTo ErrHand
    If sgUstWin(14) = "V" Then
        grdExport.Redraw = True
        Exit Sub
    End If
    
    If (grdExport.MouseRow < grdExport.FixedRows) Then
        'sort
        grdExport.Redraw = True
        Exit Sub
    End If
    If (grdExport.MouseRow < grdExport.FixedRows) Or (grdExport.MouseRow >= grdExport.Rows) Or (grdExport.MouseCol < grdExport.FixedCols) Or (grdExport.MouseCol > CLOSEINDEX) Then
        grdExport.Redraw = True
        Exit Sub
    End If
    
    If (grdExport.MouseCol = VEHICLEINDEX) And (grdExport.TextMatrix(grdExport.MouseRow, VEHICLEINDEX) <> "") Then
        grdExport.Row = grdExport.MouseRow
        grdExport.Col = grdExport.MouseCol
        llRow = grdExport.Row
        If grdExport.CellBackColor = LIGHTYELLOW Then
            grdExport.Redraw = True
            Exit Sub
        End If
'        ReDim tgEhtInfo(0 To 0) As EHTINFO
'        ReDim tgEvtInfo(0 To 0) As EVTINFO
'        ReDim tgEctInfo(0 To 0) As ECTINFO
'        sgExportTypeChar = grdExport.TextMatrix(llRow, EHTTYPECHARINDEX)
'        sgExportName = grdExport.TextMatrix(llRow, EXPORTNAMEINDEX)
'        lgExportEhtCode = Val(grdExport.TextMatrix(llRow, EHTCODEINDEX))
'        If grdExport.TextMatrix(llRow, EHTINFOINDEX) = "" Then
'            ilFound = False
'            lgExportEhtInfoIndex = UBound(tgEhtInfo)
'            llEhtInfo = lgExportEhtInfoIndex
'            tgEhtInfo(llEhtInfo).lEhtCode = 0
'            tgEhtInfo(llEhtInfo).lFirstEvt = -1
'            tgEhtInfo(llEhtInfo).lFirstEct = -1
'            tgEhtInfo(llEhtInfo).lStdEhtCode = lgExportEhtCode
'            SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & lgExportEhtCode
'            Set rst_Evt = gSQLSelectCall(SQLQuery)
'            Do While Not rst_Evt.EOF
'                If tgEhtInfo(llEhtInfo).lFirstEvt = -1 Then
'                    llNext = -1
'                Else
'                    llNext = tgEhtInfo(llEhtInfo).lFirstEvt
'                End If
'                llEvtInfo = UBound(tgEvtInfo)
'                tgEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
'                tgEvtInfo(llEvtInfo).iVefCode = rst_Evt!evtVefCode
'                tgEvtInfo(llEvtInfo).lNextEvt = llNext
'                ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
'                rst_Evt.MoveNext
'            Loop
'
'            SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & lgExportEhtCode
'            Set rst_Ect = gSQLSelectCall(SQLQuery)
'            Do While Not rst_Ect.EOF
'                If tgEhtInfo(llEhtInfo).lFirstEct = -1 Then
'                    llNext = -1
'                Else
'                    llNext = tgEhtInfo(llEhtInfo).lFirstEct
'                End If
'                llEctInfo = UBound(tgEctInfo)
'                tgEhtInfo(llEhtInfo).lFirstEct = llEctInfo
'                tgEctInfo(llEctInfo).sLogType = rst_Ect!ectLogType
'                tgEctInfo(llEctInfo).sFieldType = rst_Ect!ectFieldType
'                tgEctInfo(llEctInfo).sFieldName = rst_Ect!ectFieldName
'                tgEctInfo(llEctInfo).lFieldValue = rst_Ect!ectFieldValue
'                tgEctInfo(llEctInfo).sFieldString = rst_Ect!ectFieldString
'                tgEctInfo(llEctInfo).lNextEct = llNext
'                ReDim Preserve tgEctInfo(0 To UBound(tgEctInfo) + 1) As ECTINFO
'                rst_Ect.MoveNext
'            Loop
'            ReDim Preserve tgEhtInfo(0 To UBound(tgEhtInfo) + 1) As EHTINFO
'        Else
'            ilFound = True
'            lgExportEhtInfoIndex = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
'
'            llEhtInfo = UBound(tgEhtInfo)
'            tgEhtInfo(llEhtInfo).lEhtCode = tmSplitEhtInfo(lgExportEhtInfoIndex).lEhtCode
'            tgEhtInfo(llEhtInfo).lFirstEvt = -1
'            tgEhtInfo(llEhtInfo).lFirstEct = -1
'            tgEhtInfo(llEhtInfo).lStdEhtCode = lgExportEhtCode
'            llEvtNext = tmSplitEhtInfo(lgExportEhtInfoIndex).lFirstEvt
'            Do While llEvtNext <> -1
'                If tgEhtInfo(llEhtInfo).lFirstEvt = -1 Then
'                    llNext = -1
'                Else
'                    llNext = tgEhtInfo(llEhtInfo).lFirstEvt
'                End If
'                llEvtInfo = UBound(tgEvtInfo)
'                tgEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
'                tgEvtInfo(llEvtInfo).iVefCode = tmSplitEvtInfo(llEvtNext).iVefCode
'                tgEvtInfo(llEvtInfo).lNextEvt = llNext
'                ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
'                llEvtNext = tmSplitEvtInfo(llEvtNext).lNextEvt
'            Loop
'            llEctNext = tmSplitEhtInfo(lgExportEhtInfoIndex).lFirstEct
'            Do While llEctNext <> -1
'                If tgEhtInfo(llEhtInfo).lFirstEct = -1 Then
'                    llNext = -1
'                Else
'                    llNext = tgEhtInfo(llEhtInfo).lFirstEct
'                End If
'                llEctInfo = UBound(tgEctInfo)
'                tgEhtInfo(llEhtInfo).lFirstEct = llEctInfo
'                tgEctInfo(llEctInfo).sLogType = tmSplitEctInfo(llEctNext).sLogType
'                tgEctInfo(llEctInfo).sFieldType = tmSplitEctInfo(llEctNext).sFieldType
'                tgEctInfo(llEctInfo).sFieldName = tmSplitEctInfo(llEctNext).sFieldName
'                tgEctInfo(llEctInfo).lFieldValue = tmSplitEctInfo(llEctNext).lFieldValue
'                tgEctInfo(llEctInfo).sFieldString = tmSplitEctInfo(llEctNext).sFieldString
'                tgEctInfo(llEctInfo).lNextEct = llNext
'                ReDim Preserve tgEctInfo(0 To UBound(tgEctInfo) + 1) As ECTINFO
'                llEctNext = tmSplitEctInfo(llEctNext).lNextEct
'            Loop
'            ReDim Preserve tgEhtInfo(0 To UBound(tgEhtInfo) + 1) As EHTINFO
'        End If
        grdExport.Col = LOGSTATUSINDEX
        slLogColor = "N"
        If grdExport.CellBackColor = MIDGREENCOLOR Then
            slLogColor = "G"
        ElseIf grdExport.CellBackColor = vbRed Then
            slLogColor = "R"
        End If
        grdExport.Col = COPYSTATUSINDEX
        slCopyColor = "N"
        If grdExport.CellBackColor = MIDGREENCOLOR Then
            slCopyColor = "G"
        ElseIf grdExport.CellBackColor = vbRed Then
            slCopyColor = "R"
        End If
        grdExport.Col = GENINDEX
        slGenFont = Left$(grdExport.CellFontName, 1)
        slGen = grdExport.Text
        mPreSplit llRow
        
        lgExportEhtInfoIndex = 0
        ReDim ilVefCode(0 To 0) As Integer
        llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
        Do While llNext <> -1
            ilVefCode(UBound(ilVefCode)) = tgEvtInfo(llNext).iVefCode
            ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
            llNext = tgEvtInfo(llNext).lNextEvt
        Loop
        
        frmVehicleSelection.Show vbModal
        If igExportReturn = 0 Then
            'If Not ilFound Then
            '    llNext = tmSplitEhtInfo(lgExportEhtInfoIndex).lFirstEvt
            '    Do While llNext <> -1
            '        llSvNext = tmSplitEvtInfo(llNext).lNextEvt
            '        tmSplitEvtInfo(llNext).lNextEvt = -9999
            '        llNext = llSvNext
            '    Loop
            '    llNext = tmSplitEhtInfo(lgExportEhtInfoIndex).lFirstEct
            '    Do While llNext <> -1
            '        llSvNext = tgEctInfo(llNext).lNextEct
            '        tgEctInfo(llNext).lNextEct = -9999
            '        llNext = llSvNext
            '    Loop
            '    ReDim Preserve tmSplitEhtInfo(0 To UBound(tmSplitEhtInfo) - 1) As EHTINFO
            'End If
            grdExport.Redraw = False
            grdExport.Redraw = True
            Exit Sub
        End If
        grdExport.Redraw = False
        gSetMousePointer grdExport, grdStatus, vbHourglass
        gSetMousePointer grdAlerts, grdStations, vbHourglass
        
        ilCount = 0
        llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
        Do While llNext <> -1
            ilCount = ilCount + 1
            llNext = tgEvtInfo(llNext).lNextEvt
        Loop
        If ilCount <> UBound(ilVefCode) Then
        
'            llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
'            Do While llNext <> -1
'                For llVef = 0 To UBound(ilVefCode) - 1 Step 1
'                    If ilVefCode(llVef) = tgEvtInfo(llNext).iVefCode Then
'                        ilVefCode(llVef) = -1
'                    End If
'                Next llVef
'                llNext = tgEvtInfo(llNext).lNextEvt
'            Loop
'            llEhtInfo = UBound(tgEhtInfo)
'            tgEhtInfo(llEhtInfo).iRefRowNo = 0
'            tgEhtInfo(llEhtInfo).lFirstEvt = -1
'            tgEhtInfo(llEhtInfo).lFirstEct = -1
'            tgEhtInfo(llEhtInfo).lStdEhtCode = lgExportEhtCode
'            For llVef = 0 To UBound(ilVefCode) - 1 Step 1
'                If ilVefCode(llVef) > 0 Then
'                    If tgEhtInfo(llEhtInfo).lFirstEvt = -1 Then
'                        llNext = -1
'                    Else
'                        llNext = tgEhtInfo(llEhtInfo).lFirstEvt
'                    End If
'                    llEvtInfo = UBound(tgEvtInfo)
'                    tgEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
'                    tgEvtInfo(llEvtInfo).iVefCode = ilVefCode(llVef)
'                    tgEvtInfo(llEvtInfo).lNextEvt = llNext
'                    ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
'                End If
'            Next llVef
'
'            llCopyEct = tgEctInfo(lgExportEhtInfoIndex).lNextEct
'            Do While llCopyEct <> -1
'                If tgEhtInfo(llEhtInfo).lFirstEct = -1 Then
'                    llNext = -1
'                Else
'                    llNext = tgEhtInfo(llEhtInfo).lFirstEct
'                End If
'                llEctInfo = UBound(tgEctInfo)
'                tgEhtInfo(llEhtInfo).lFirstEct = llEctInfo
'                tgEctInfo(llEctInfo).sLogType = tgEctInfo(llCopyEct).sLogType
'                tgEctInfo(llEctInfo).sFieldType = tgEctInfo(llCopyEct).sFieldType
'                tgEctInfo(llEctInfo).sFieldName = tgEctInfo(llCopyEct).sFieldName
'                tgEctInfo(llEctInfo).lFieldValue = tgEctInfo(llCopyEct).lFieldValue
'                tgEctInfo(llEctInfo).sFieldString = tgEctInfo(llCopyEct).sFieldString
'                tgEctInfo(llEctInfo).lNextEct = llNext
'                ReDim Preserve tgEctInfo(0 To UBound(tgEctInfo) + 1) As ECTINFO
'                llCopyEct = tgEctInfo(llCopyEct).lNextEct
'            Loop
'            ReDim Preserve tgEhtInfo(0 To UBound(tgEhtInfo) + 1) As EHTINFO
        
            'Create two images, one from tgEhtInfo and the other from ilVefCode
            'Step 1: The ones that will be created from tgEctInfo need to be removed from ilVefCode
            llNext = tgEhtInfo(0).lFirstEvt
            Do While llNext <> -1
                For llVef = 0 To UBound(ilVefCode) - 1 Step 1
                    If ilVefCode(llVef) = tgEvtInfo(llNext).iVefCode Then
                        ilVefCode(llVef) = -1
                    End If
                Next llVef
                llNext = tgEvtInfo(llNext).lNextEvt
            Loop
            
'            llEvtNext = tgEhtInfo(0).lFirstEvt
'            llVef = 0
'            For ilPass = 0 To 1 Step 1
'                If (ilPass = 0) And ilFound Then
'                    llEhtInfo = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
'                Else
'                    llEhtInfo = UBound(tmSplitEhtInfo)
'                    tmSplitEhtInfo(llEhtInfo).lEhtCode = 0
'                End If
'                tmSplitEhtInfo(llEhtInfo).lFirstEvt = -1
'                tmSplitEhtInfo(llEhtInfo).lFirstEct = -1
'                tmSplitEhtInfo(llEhtInfo).lStdEhtCode = lgExportEhtCode
'                Do
'                    If tmSplitEhtInfo(llEhtInfo).lFirstEvt = -1 Then
'                        llNext = -1
'                    Else
'                        llNext = tmSplitEhtInfo(llEhtInfo).lFirstEvt
'                    End If
'                    ilSave = False
'                    llEvtInfo = UBound(tmSplitEvtInfo)
'                    If ilPass = 0 Then
'                        If llEvtNext <> -1 Then
'                            tmSplitEvtInfo(llEvtInfo).iVefCode = tgEvtInfo(llEvtNext).iVefCode
'                            llEvtNext = tgEvtInfo(llEvtNext).lNextEvt
'                            ilSave = True
'                        End If
'                    Else
'                        Do While llVef < UBound(ilVefCode)
'                            If ilVefCode(llVef) > 0 Then
'                                tmSplitEvtInfo(llEvtInfo).iVefCode = ilVefCode(llVef)
'                                llVef = llVef + 1
'                                ilSave = True
'                                Exit Do
'                            End If
'                            llVef = llVef + 1
'                        Loop
'                    End If
'                    If ilSave Then
'                        tmSplitEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
'                        tmSplitEvtInfo(llEvtInfo).lNextEvt = llNext
'                        ReDim Preserve tmSplitEvtInfo(0 To UBound(tmSplitEvtInfo) + 1) As EVTINFO
'                    Else
'                        Exit Do
'                    End If
'                Loop
'                llCopyEct = tgEhtInfo(0).lFirstEct
'                Do While llCopyEct <> -1
'                    If tmSplitEhtInfo(llEhtInfo).lFirstEct = -1 Then
'                        llNext = -1
'                    Else
'                        llNext = tmSplitEhtInfo(llEhtInfo).lFirstEct
'                    End If
'                    llEctInfo = UBound(tmSplitEctInfo)
'                    tmSplitEhtInfo(llEhtInfo).lFirstEct = llEctInfo
'                    tmSplitEctInfo(llEctInfo).sLogType = tgEctInfo(llCopyEct).sLogType
'                    tmSplitEctInfo(llEctInfo).sFieldType = tgEctInfo(llCopyEct).sFieldType
'                    tmSplitEctInfo(llEctInfo).sFieldName = tgEctInfo(llCopyEct).sFieldName
'                    tmSplitEctInfo(llEctInfo).lFieldValue = tgEctInfo(llCopyEct).lFieldValue
'                    tmSplitEctInfo(llEctInfo).sFieldString = tgEctInfo(llCopyEct).sFieldString
'                    tmSplitEctInfo(llEctInfo).lNextEct = llNext
'                    ReDim Preserve tmSplitEctInfo(0 To UBound(tmSplitEctInfo) + 1) As ECTINFO
'                    llCopyEct = tgEctInfo(llCopyEct).lNextEct
'                Loop
'                If (ilPass = 1) Or (Not ilFound) Then
'                    tmSplitEhtInfo(llEhtInfo).blRemoved = False
'                    ReDim Preserve tmSplitEhtInfo(0 To UBound(tmSplitEhtInfo) + 1) As EHTINFO
'                End If
'            Next ilPass
            mPostSplit llRow, ilVefCode(), slGenFont, slGen, LOGSTATUSINDEX, slLogColor, slLogColor, COPYSTATUSINDEX, slCopyColor, slCopyColor
        End If
        gSetMousePointer grdExport, grdStatus, vbHourglass
        gSetMousePointer grdAlerts, grdStations, vbHourglass
        grdExport.Redraw = False
        mExportPopulate
        mSetExportGridColors
        gSetMousePointer grdExport, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        grdExport.Redraw = True
        If bmDateError Then
            MsgBox "Red in Gen field indicates the date range crosses Sunday and must be fixed in Specifications", vbCritical + vbOKOnly
        End If
    ElseIf (grdExport.MouseCol = CLOSEINDEX) And (grdExport.TextMatrix(grdExport.MouseRow, CLOSEINDEX) = "X") Then
        grdExport.Redraw = False
        gSetMousePointer grdExport, grdStatus, vbHourglass
        gSetMousePointer grdAlerts, grdStations, vbHourglass
        llRow = grdExport.MouseRow
        llStdEhtCode = tmSplitEhtInfo(Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))).lStdEhtCode
        slLogColor = ""
        slCopyColor = ""
        'Hierachy: N, R finally G
        For llEht = 0 To UBound(tmSplitEhtInfo) - 1 Step 1
            If Not tmSplitEhtInfo(llEht).blRemoved Then
                If llStdEhtCode = tmSplitEhtInfo(llEht).lStdEhtCode Then
                    If slLogColor = "" Then
                        slLogColor = tmSplitEhtInfo(llEht).sLogStatus
                    Else
                        If slLogColor <> "N" Then
                            If tmSplitEhtInfo(llEht).sLogStatus = "N" Then
                                slLogColor = "N"
                            ElseIf slLogColor = "G" Then
                                If tmSplitEhtInfo(llEht).sLogStatus <> "G" Then
                                    slLogColor = "R"
                                End If
                            End If
                        End If
                    End If
                    If slCopyColor = "" Then
                        slCopyColor = tmSplitEhtInfo(llEht).sCopyStatus
                    Else
                        If slCopyColor <> "N" Then
                            If tmSplitEhtInfo(llEht).sCopyStatus = "N" Then
                                slCopyColor = "N"
                            ElseIf slLogColor = "G" Then
                                If tmSplitEhtInfo(llEht).sCopyStatus <> "G" Then
                                    slCopyColor = "R"
                                End If
                            End If
                        End If
                    End If
                    If tmSplitEhtInfo(llEht).lEhtCode > 0 Then
                        mRemoveEht tmSplitEhtInfo(llEht).lEhtCode
                    End If
                    tmSplitEhtInfo(llEht).blRemoved = True
                End If
            End If
        Next llEht
        For llEht = 0 To UBound(tmEhtStdColor) - 1 Step 1
            If tmEhtStdColor(llEht).lEhtCode = llStdEhtCode Then
                tmEhtStdColor(llEht).sLogStatus = slLogColor
                tmEhtStdColor(llEht).sCopyStatus = slCopyColor
                Exit For
            End If
        Next llEht
        grdExport.TextMatrix(llRow, EHTINFOINDEX) = ""
        mExportPopulate
        mSetExportGridColors
        gSetMousePointer grdExport, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        grdExport.Redraw = True
        If bmDateError Then
            MsgBox "Red in Gen field indicates the date range crosses Sunday and must be fixed in Specifications", vbCritical + vbOKOnly
        End If
    Else
        grdExport.Row = grdExport.MouseRow
        grdExport.Col = grdExport.MouseCol
        If Not mColOk() Then
            grdExport.Redraw = True
            On Error Resume Next
            cmdCancel.SetFocus
            Exit Sub
        End If
        grdExport.Redraw = True
        mEnableBox
    End If
    Exit Sub
ErrHand:
    gSetMousePointer grdExport, grdStatus, vbDefault
    gSetMousePointer grdAlerts, grdStations, vbDefault
    gHandleError "AffErrorLog.txt", "Export-grdExport_MouseUp"
End Sub

Private Sub grdExport_Scroll()
    pbcClickFocus.SetFocus
    'mSetShow
    'mStatusSetShow
End Sub

Private Sub grdStatus_EnterCell()
    mSetShow
    mStatusSetShow
End Sub

Private Sub grdStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmStatusTopRow = grdExport.TopRow
    grdStatus.Redraw = False
End Sub

Private Sub grdStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim slStr As String
    slStr = ""
    If Not imCtrlVisible Then
        If (grdStatus.MouseRow >= grdStatus.FixedRows) And (grdStatus.TextMatrix(grdStatus.MouseRow, grdStatus.MouseCol)) <> "" Then
            If (grdStatus.MouseCol <> SUSERINDEX) And (grdStatus.MouseCol <> SSTATUSINDEX) And (grdStatus.MouseCol <> SCLOSEINDEX) Then
                slStr = Trim$(grdStatus.TextMatrix(grdStatus.MouseRow, grdStatus.MouseCol))
            ElseIf grdStatus.MouseCol = SUSERINDEX Then
                slStr = Trim$(grdStatus.TextMatrix(grdStatus.MouseRow, SUSERNAMEINDEX))
            ElseIf grdStatus.MouseCol = SSTATUSINDEX Then
                slStr = Trim$(grdStatus.TextMatrix(grdStatus.MouseRow, SRESULTFILEINDEX))
            End If
        End If
    End If
    If smStatusPrevTip <> slStr Then
        grdStatus.ToolTipText = slStr
    End If
    smStatusPrevTip = slStr

End Sub

Private Sub grdStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    Dim llEhtInfo As Long
    Dim llEvtInfo As Long
    Dim llEctInfo As Long
    Dim llNext As Long
    Dim ilCount As Integer
    Dim llSvNext As Long
    Dim llVef As Long
    Dim llCopyEct As Long
    Dim llEht As Long
    Dim llStdEhtCode As Long
    Dim llEqtCode As Long
    Dim llEhtCode As Long
    Dim ilPriority As Integer
    
    On Error GoTo ErrHand
    If sgUstWin(14) = "V" Then
        grdStatus.Redraw = True
        Exit Sub
    End If
    If (grdStatus.MouseRow < grdStatus.FixedRows) Then
        'sort
        grdStatus.Redraw = True
        Exit Sub
    End If
    If (grdStatus.MouseRow < grdStatus.FixedRows) Or (grdStatus.MouseRow >= grdStatus.Rows) Or (grdStatus.MouseCol < grdStatus.FixedCols) Or (grdStatus.MouseCol > CLOSEINDEX) Then
        grdStatus.Redraw = True
        Exit Sub
    End If
    If grdStatus.TextMatrix(grdStatus.MouseRow, SEXPORTTYPEINDEX) = "" Then
        grdStatus.Redraw = True
        Exit Sub
    End If
    
    If (grdStatus.MouseCol = SPRIORITYINDEX) Then
        grdStatus.Row = grdStatus.MouseRow
        grdStatus.Col = grdStatus.MouseCol
        If grdStatus.CellBackColor = LIGHTYELLOW Then
            grdStatus.Redraw = True
            Exit Sub
        End If
        grdStatus.Redraw = True
        mStatusEnableBox
    ElseIf (grdStatus.MouseCol = SCLOSEINDEX) And (grdStatus.TextMatrix(grdStatus.MouseRow, SCLOSEINDEX) = "X") Then
        grdStatus.Redraw = False
        gSetMousePointer grdStatus, grdStatus, vbHourglass
        gSetMousePointer grdAlerts, grdStations, vbHourglass
        llEqtCode = Val(grdStatus.TextMatrix(grdStatus.MouseRow, SEQTCODEINDEX))
        SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtCode = " & llEqtCode
        Set rst_Eqt = gSQLSelectCall(SQLQuery)
        If rst_Eqt.EOF Then
            grdStatus.Redraw = True
            Exit Sub
        End If
        ilPriority = rst_Eqt!eqtPriority
        'SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & rst_Eqt!eqtEhtCode
        'Set rst_Eht = gSQLSelectCall(SQLQuery)
        'If Not rst_Eht.EOF Then
        '    If rst_Eht!ehtStandardEhtCode > 0 Then
        '        mRemoveEht rst_Eht!ehtCode
        '    End If
        'End If
        SQLQuery = "DELETE FROM eqt_Export_Queue WHERE eqtCode = " & llEqtCode
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            gSetMousePointer grdStatus, grdStatus, vbDefault
            gSetMousePointer grdAlerts, grdStations, vbDefault
            gHandleError "AffErrorLog.txt", "Export-grdStatus_MouseUp"
            Exit Sub
        End If
        mAdjustPriority ilPriority
        mExportPopulate
        mSetExportGridColors
        mStatusPopulate
        mSetStatusGridColor
        gSetMousePointer grdStatus, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        grdExport.Redraw = True
        grdStatus.Redraw = True
        If bmDateError Then
            MsgBox "Red in Gen field indicates the date range crosses Sunday and must be fixed in Specifications", vbCritical + vbOKOnly
        End If
    ElseIf (grdStatus.MouseCol = SSTATUSINDEX) Then
        grdStatus.Row = grdStatus.MouseRow
        grdStatus.Col = grdStatus.MouseCol
        If grdStatus.CellBackColor = LIGHTGREENCOLOR Then
            sgResultFileName = grdStatus.TextMatrix(grdStatus.Row, SRESULTFILEINDEX)
            igModelType = 3
            frmModel.Show vbModal
        End If
        grdStatus.Redraw = True
    Else
        grdStatus.Redraw = True
    End If
    Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gSetMousePointer grdAlerts, grdStations, vbDefault
    gHandleError "AffErrorLog.txt", "Export-grdStatus_MouseUp"
End Sub

Private Sub grdStatus_Scroll()
    pbcClickFocus.SetFocus
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbcKey.Visible = True
    lbcKey.ZOrder
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbcKey.Visible = False
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
    mStatusSetShow
End Sub

Private Sub pbcGen_KeyPress(KeyAscii As Integer)
    If lmEnableCol = GENINDEX Then
        If KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            smGen = "N"
            pbcGen_Paint
        ElseIf KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            smGen = "Y"
            pbcGen_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If smGen = "N" Then
                smGen = "Y"
                pbcGen_Paint
            ElseIf smGen = "Y" Then
                smGen = "N"
                pbcGen_Paint
            End If
        End If
    End If
End Sub

Private Sub pbcGen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lmEnableCol = GENINDEX Then
        If smGen = "N" Then
            smGen = "Y"
            pbcGen_Paint
        ElseIf smGen = "Y" Then
            smGen = "N"
            pbcGen_Paint
        End If
    End If
End Sub

Private Sub pbcGen_Paint()
    pbcGen.Cls
    pbcGen.CurrentX = 30
    pbcGen.CurrentY = 30
    If lmEnableCol = GENINDEX Then
        If smGen = "N" Then
            pbcGen.Print "  "
        ElseIf smGen = "Y" Then
            pbcGen.Print "4"
        Else
            pbcGen.Print "   "
        End If
    End If
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        'Branch
        Do
            ilNext = False
            Select Case lmEnableCol
                Case GENINDEX
                    If grdExport.Row = grdExport.FixedRows Then
                        mSetShow
                        cmdGenerate.SetFocus
                        Exit Sub
                    End If
                    lmTopRow = -1
                    grdExport.Row = grdExport.Row - 1
                    If Not grdExport.RowIsVisible(grdExport.Row) Then
                        grdExport.TopRow = grdExport.TopRow - 1
                    End If
                    grdExport.Col = ENDDATEINDEX
                Case Else
                    grdExport.Col = grdExport.Col - 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        lmTopRow = -1
        grdExport.TopRow = grdExport.FixedRows
        grdExport.Row = grdExport.FixedRows
        grdExport.Col = GENINDEX
        Do
            If mColOk() Then
                Exit Do
            End If
            If grdExport.Row + 1 >= grdExport.Rows Then
                cmdCancel.SetFocus
                Exit Sub
            End If
            grdExport.Row = grdExport.Row + 1
            Do
                If Not grdExport.RowIsVisible(grdExport.Row) Then
                    grdExport.TopRow = grdExport.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
        Loop
    End If
    mEnableBox
End Sub

Private Sub pbcStatusSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcStatusSTab.hwnd Then
        Exit Sub
    End If
    If imStatusCtrlVisible Then
        'Branch
        Do
            ilNext = False
            Select Case lmStatusEnableCol
                Case SPRIORITYINDEX
                    mStatusSetShow
            End Select
        Loop While ilNext
    End If
    cmdCancel.SetFocus
End Sub

Private Sub pbcStatusTab_GotFocus()
    Dim slStr As String
    Dim ilNext As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long

    If GetFocus() <> pbcStatusTab.hwnd Then
        Exit Sub
    End If
    If imStatusCtrlVisible Then
        'Branch
        Do
            ilNext = False
            Select Case lmStatusEnableCol
                Case SPRIORITYINDEX
                    mSetShow
            End Select
        Loop While ilNext
    End If
    cmdCancel.SetFocus
End Sub

Private Sub pbcTab_GotFocus()
    Dim slStr As String
    Dim ilNext As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        'Branch
        Do
            ilNext = False
            Select Case lmEnableCol
                Case ENDDATEINDEX
                    llEnableRow = lmEnableRow
                    llEnableCol = lmEnableCol
                    mSetShow
                    lmEnableRow = llEnableRow
                    lmEnableCol = llEnableCol
                    If (grdExport.Row + 1 >= grdExport.Rows) Then
                        cmdGenerate.SetFocus
                        Exit Sub
                    End If
                    If (grdExport.Row + 1 < grdExport.Rows) Then
                        If (Trim$(grdExport.TextMatrix(grdExport.Row + 1, EXPORTTYPEINDEX)) = "") Then
                            cmdGenerate.SetFocus
                            Exit Sub
                        End If
                    End If
                    grdExport.Row = grdExport.Row + 1
                    grdExport.Col = GENINDEX
                    If Not grdExport.RowIsVisible(grdExport.Row) Then
                        grdExport.TopRow = grdExport.TopRow + 1
                    End If
                Case Else
                    grdExport.Col = grdExport.Col + 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        grdExport.TopRow = grdExport.FixedRows
        grdExport.Col = GENINDEX
        Do
            If grdExport.Row <= grdExport.FixedRows Then
                cmdCancel.SetFocus
                Exit Sub
            End If
            grdExport.Row = grdExport.Rows - 1
            Do
                If Not grdExport.RowIsVisible(grdExport.Row) Then
                    grdExport.TopRow = grdExport.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
            If mColOk() Then
                Exit Do
            End If
        Loop
    End If
    mEnableBox

End Sub

Private Sub TabStrip1_Click()
    Dim iLoop As Integer
    Dim iZone As Integer
    
    
    If imTabIndex = TabStrip1.SelectedItem.Index Then
        Exit Sub
    End If
    'Selection
    If TabStrip1.SelectedItem.Index = 1 Then
        frcTab(0).Visible = True
        frcTab(1).Visible = False
        frcTab(2).Visible = False
    'Alert
    ElseIf TabStrip1.SelectedItem.Index = 2 Then
        frcTab(1).Visible = True
        frcTab(0).Visible = False
        frcTab(2).Visible = False
        cmdGenerate.Enabled = False
        cmcCustom.Enabled = False
        cmcSpec.Enabled = False
        '1/11/17 Dan
        cmcCheckCopy.Enabled = False
    ElseIf TabStrip1.SelectedItem.Index = 3 Then
        frcTab(2).Visible = True
        frcTab(1).Visible = False
        frcTab(0).Visible = False
    End If
    If sgUstWin(14) = "V" Then
        cmdGenerate.Enabled = False
        cmcCustom.Enabled = False
        cmcSpec.Enabled = False
        '1/11/17 Dan
        cmcCheckCopy.Enabled = False
    Else
        If TabStrip1.SelectedItem.Index = 1 Then
            cmdGenerate.Enabled = True
            cmcCustom.Enabled = True
            cmcSpec.Enabled = True
            '1/11/17 Dan
            cmcCheckCopy.Enabled = True
        ElseIf TabStrip1.SelectedItem.Index = 2 Then
            cmdGenerate.Enabled = False
            cmcCustom.Enabled = False
            cmcSpec.Enabled = False
            '1/11/17 Dan
            cmcCheckCopy.Enabled = False
        End If
    End If
    imTabIndex = TabStrip1.SelectedItem.Index
End Sub

Private Sub mExportPopulate()
    Dim llRow As Long
    Dim ilAdj As Integer
    Dim slWorkDate As String
    Dim slNowDate As String
    Dim llEht As Long
    Dim llEhtIndex As Long
    Dim ilCount As Integer
    Dim llNext As Long
    Dim blFound As Boolean
    Dim blFirstStep As Boolean
    Dim slDate As String
    Dim ilLastSColSorted As Integer
    Dim ilLastSSort As Integer
    Dim blDisplayEht As Boolean
    Dim ilFound As Integer
    Dim ilLogAndCopyColor As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilLoop As Integer
    
    On Error GoTo ErrHand
    mClearGrid grdExport
    bmDateError = False
    llRow = grdExport.FixedRows
    'Only get the standard as the Alter placed into tmSplit---Info
    SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtSubType = 'S'"
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        blFirstStep = False
        If gDateValue(Format(rst_Eht!ehtLDE, sgShowDateForm)) <> gDateValue("1/1/1970") Then
            ilFound = False
            For llEht = 0 To UBound(tmEhtStdColor) - 1 Step 1
                If tmEhtStdColor(llEht).lEhtCode = rst_Eht!ehtCode Then
                    ilFound = True
                    llEhtIndex = llEht
                    Exit For
                End If
            Next llEht
            If Not ilFound Then
                llEhtIndex = UBound(tmEhtStdColor)
                tmEhtStdColor(llEhtIndex).lEhtCode = rst_Eht!ehtCode
                tmEhtStdColor(llEhtIndex).sLogStatus = "N"
                tmEhtStdColor(llEhtIndex).sCopyStatus = "N"
                tmEhtStdColor(llEhtIndex).sGenFont = "N"
                tmEhtStdColor(llEhtIndex).sGen = ""
                ReDim Preserve tmEhtStdColor(0 To llEhtIndex + 1) As EHTSTDCOLOR
            End If
            blDisplayEht = False
            'Check that not in status area
            SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtEhtCode = " & rst_Eht!ehtCode & " AND (eqtStatus = 'P' or eqtStatus = 'R')"
            Set rst_Eqt = gSQLSelectCall(SQLQuery)
            If rst_Eqt.EOF Then
                For llEht = 0 To UBound(tmSplitEhtInfo) - 1 Step 1
                    If Not tmSplitEhtInfo(llEht).blRemoved Then
                        If rst_Eht!ehtCode = tmSplitEhtInfo(llEht).lStdEhtCode Then
                            blDisplayEht = True
                            Exit For
                        End If
                    End If
                Next llEht
                If Not blDisplayEht Then
                    blDisplayEht = True
                    For llEht = 0 To UBound(lmStandardEhtCode) - 1 Step 1
                        If rst_Eht!ehtCode = lmStandardEhtCode(llEht) Then
                            blDisplayEht = False
                            Exit For
                        End If
                    Next llEht
                End If
            End If
            If blDisplayEht Then
                Do
                    If llRow >= grdExport.Rows Then
                        grdExport.AddItem ""
                    End If
                    grdExport.Row = llRow
                    grdExport.Col = GENINDEX
                    grdExport.CellFontName = "Monotype Sorts"
                    grdExport.TextMatrix(llRow, EHTTYPECHARINDEX) = rst_Eht!ehtExportType
                    grdExport.TextMatrix(llRow, EHTINFOINDEX) = ""
                    Select Case rst_Eht!ehtExportType
                        Case "A"
                            grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "Aff Logs"
                        Case "C"
                            grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "C & C"
                        Case "D"
                            grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "IDC"
                        Case "I"
                            grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "ISCI"
                        Case "R"
                            grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "ISCI C/R"
                        Case "4"
                            grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "RCS 4"
                        Case "5"
                            grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "RCS 5"
                        Case "S"
                            grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "StarGuide"
                        Case "W"
                            grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "Compel"
                        Case "X"
                            grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "X-Digital"
                        Case "P"
                             grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) = "IPump"
                    End Select
                    grdExport.TextMatrix(llRow, EXPORTNAMEINDEX) = rst_Eht!ehtExportName
                    grdExport.TextMatrix(llRow, LASTDATEINDEX) = Format(rst_Eht!ehtLDE, sgShowDateForm)
                    grdExport.TextMatrix(llRow, LEADTIMEINDEX) = rst_Eht!ehtLeadTime
                    grdExport.TextMatrix(llRow, CYCLEINDEX) = rst_Eht!ehtCycle
                    grdExport.TextMatrix(llRow, STARTDATEINDEX) = DateAdd("D", 1, grdExport.TextMatrix(llRow, LASTDATEINDEX))
                    grdExport.TextMatrix(llRow, ENDDATEINDEX) = DateAdd("D", grdExport.TextMatrix(llRow, CYCLEINDEX) - 1, grdExport.TextMatrix(llRow, STARTDATEINDEX))
                    grdExport.TextMatrix(llRow, WORKDATEINDEX) = DateAdd("D", -grdExport.TextMatrix(llRow, LEADTIMEINDEX), grdExport.TextMatrix(llRow, STARTDATEINDEX))
                    grdExport.TextMatrix(llRow, EHTCODEINDEX) = rst_Eht!ehtCode
                    
                    '3/16/17: Check if dates span Sunday and export type must not permit date range to across Sunday
                    For ilLoop = LBound(tgSpecInfo) To UBound(tgSpecInfo) Step 1
                        If tgSpecInfo(ilLoop).sType = grdExport.TextMatrix(llRow, EHTTYPECHARINDEX) Then
                            If tgSpecInfo(ilLoop).sCheckDateSpan = "Y" Then
                                slStartDate = grdExport.TextMatrix(llRow, STARTDATEINDEX)
                                slEndDate = grdExport.TextMatrix(llRow, ENDDATEINDEX)
                                If gWeekDayLong(gDateValue(slEndDate)) <= gWeekDayLong(gDateValue(slStartDate)) Then
                                    bmDateError = True
                                    grdExport.CellBackColor = vbRed
                                End If
                            End If
                            Exit For
                        End If
                    Next ilLoop
                    
                    SQLQuery = "SELECT Count(evtCode) FROM evt_Export_Vehicles WHERE evtEhtCode = " & rst_Eht!ehtCode
                    Set rst_Evt = gSQLSelectCall(SQLQuery)
                    If Not rst_Evt.EOF Then
                        grdExport.TextMatrix(llRow, VEHICLEINDEX) = rst_Evt(0).Value
                    Else
                        grdExport.TextMatrix(llRow, VEHICLEINDEX) = ""
                    End If
                    'Determine if Partial created
                    blFound = False
                    For llEht = 0 To UBound(tmSplitEhtInfo) - 1 Step 1
                        If Not tmSplitEhtInfo(llEht).blRemoved Then
                            If rst_Eht!ehtCode = tmSplitEhtInfo(llEht).lStdEhtCode Then
                                blFound = True
                                ilCount = 0
                                llNext = tmSplitEhtInfo(llEht).lFirstEvt
                                Do While llNext <> -1
                                    ilCount = ilCount + 1
                                    llNext = tmSplitEvtInfo(llNext).lNextEvt
                                Loop
                                grdExport.TextMatrix(llRow, VEHICLEINDEX) = Trim$(Str$(ilCount)) & " of " & grdExport.TextMatrix(llRow, VEHICLEINDEX)
                                tmEhtStdColor(llEhtIndex).sGenFont = tmSplitEhtInfo(llEht).sGenFont
                                tmEhtStdColor(llEhtIndex).sGen = tmSplitEhtInfo(llEht).sGen
                                '3/24/13
                                tmEhtStdColor(llEhtIndex).sLogStatus = tmSplitEhtInfo(llEht).sLogStatus
                                tmEhtStdColor(llEhtIndex).sCopyStatus = tmSplitEhtInfo(llEht).sCopyStatus
                                tmSplitEhtInfo(llEht).lStdEhtCode = -tmSplitEhtInfo(llEht).lStdEhtCode
                                grdExport.TextMatrix(llRow, EHTINFOINDEX) = llEht
                                Exit For
                            End If
                        End If
                    Next llEht
                    grdExport.TextMatrix(llRow, EHTCODEINDEX) = rst_Eht!ehtCode
                    
                    slWorkDate = Format(grdExport.TextMatrix(llRow, WORKDATEINDEX), "m/d/yy")
                    slNowDate = Format(gNow(), "m/d/yy")
                    If tmEhtStdColor(llEhtIndex).sGenFont = "N" Then
                        If gDateValue(slWorkDate) <= gDateValue(slNowDate) And (Not blFirstStep) Then
                            grdExport.TextMatrix(llRow, GENINDEX) = "4"
                            tmEhtStdColor(llEhtIndex).sGenFont = "M"
                            tmEhtStdColor(llEhtIndex).sGen = "4"
                            blFirstStep = True
                        Else
                            If tmEhtStdColor(llEhtIndex).sGenFont = "M" Then
                                grdExport.CellFontName = "Monotype Sorts"
                                grdExport.TextMatrix(llRow, GENINDEX) = tmEhtStdColor(llEhtIndex).sGen
                            ElseIf tmEhtStdColor(llEhtIndex).sGenFont = "A" Then
                                grdExport.CellFontName = "Arial"
                                grdExport.TextMatrix(llRow, GENINDEX) = tmEhtStdColor(llEhtIndex).sGen
                            Else
                                grdExport.CellFontName = "Monotype Sorts"
                                grdExport.TextMatrix(llRow, GENINDEX) = tmEhtStdColor(llEhtIndex).sGen
                            End If
                        End If
                    Else
                        If Not blFirstStep Then
                            If tmEhtStdColor(llEhtIndex).sGenFont = "M" Then
                                grdExport.CellFontName = "Monotype Sorts"
                                grdExport.TextMatrix(llRow, GENINDEX) = tmEhtStdColor(llEhtIndex).sGen
                            ElseIf tmEhtStdColor(llEhtIndex).sGenFont = "A" Then
                                grdExport.CellFontName = "Arial"
                                grdExport.TextMatrix(llRow, GENINDEX) = tmEhtStdColor(llEhtIndex).sGen
                            Else
                                grdExport.CellFontName = "Monotype Sorts"
                                grdExport.TextMatrix(llRow, GENINDEX) = tmEhtStdColor(llEhtIndex).sGen
                            End If
                            blFirstStep = True
                        End If
                    End If
                    For ilLogAndCopyColor = 0 To UBound(tmLogAndCopyColor) - 1 Step 1
                        If rst_Eht!ehtCode = tmLogAndCopyColor(ilLogAndCopyColor).lEhtCode Then
                            tmEhtStdColor(llEhtIndex).sLogStatus = tmLogAndCopyColor(ilLogAndCopyColor).sLogColor
                            tmEhtStdColor(llEhtIndex).sCopyStatus = tmLogAndCopyColor(ilLogAndCopyColor).sCopyColor
                            Exit For
                        End If
                    Next ilLogAndCopyColor
                    If (tmEhtStdColor(llEhtIndex).sLogStatus = "R") Or (tmEhtStdColor(llEhtIndex).sCopyStatus = "R") Then
                        '3/18/17: If Priority set, don't clear
                        If tmEhtStdColor(llEhtIndex).sGenFont <> "A" Then
                            grdExport.TextMatrix(llRow, GENINDEX) = ""
                            tmEhtStdColor(llEhtIndex).sGenFont = "N"
                            tmEhtStdColor(llEhtIndex).sGen = ""
                        End If
                    End If
                    '3/16/17: Dates span Sunday, disallow generation
                    If (grdExport.CellBackColor = vbRed) Then
                        grdExport.TextMatrix(llRow, GENINDEX) = ""
                        tmEhtStdColor(llEhtIndex).sGenFont = "N"
                        tmEhtStdColor(llEhtIndex).sGen = ""
                    End If
                    'Determine if more splits required to be processed
                    If blFound Then
                        blFound = False
                        For llEht = 0 To UBound(tmSplitEhtInfo) - 1 Step 1
                            If Not tmSplitEhtInfo(llEht).blRemoved Then
                                If rst_Eht!ehtCode = tmSplitEhtInfo(llEht).lStdEhtCode Then
                                    blFound = True
                                    Exit For
                                End If
                            End If
                        Next llEht
                    End If
                    slDate = Trim$(gDateValue(grdExport.TextMatrix(llRow, WORKDATEINDEX)))
                    Do While Len(slDate) < 6
                        slDate = "0" & slDate
                    Loop
                    grdExport.TextMatrix(llRow, SORTINDEX) = slDate & Trim$(grdExport.TextMatrix(llRow, EXPORTTYPEINDEX)) & Trim$(grdExport.TextMatrix(llRow, EXPORTNAMEINDEX))
                    llRow = llRow + 1
                Loop While blFound
                For llEht = 0 To UBound(tmSplitEhtInfo) - 1 Step 1
                    If tmSplitEhtInfo(llEht).lStdEhtCode < 0 Then
                        tmSplitEhtInfo(llEht).lStdEhtCode = -tmSplitEhtInfo(llEht).lStdEhtCode
                    End If
                Next llEht
            End If
        End If
        rst_Eht.MoveNext
    Loop
    
    ilLastSColSorted = -1
    ilLastSSort = -1
    gGrid_SortByCol grdExport, EXPORTTYPEINDEX, SORTINDEX, ilLastSColSorted, ilLastSSort, False
    Exit Sub
ErrHand:
    gSetMousePointer grdExport, grdStatus, vbDefault
    gSetMousePointer grdAlerts, grdStations, vbDefault
    gHandleError "AffErrorLog.txt", "Export-mExportPopulate"

End Sub



Private Function mColOk() As Integer
    mColOk = True
    If (grdExport.CellBackColor = LIGHTYELLOW) Or (grdExport.CellBackColor = LIGHTGREENCOLOR) Or (grdExport.CellBackColor = vbRed) Then
        mColOk = False
        Exit Function
    End If
End Function


Private Sub mSetStatusGridColor()
    Dim llRow As Long
    Dim llCol As Long
    
    For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
        For llCol = SEXPORTTYPEINDEX To SCLOSEINDEX Step 1
            grdStatus.Row = llRow
            grdStatus.Col = llCol
            If grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) <> "" Then
                If (llCol <> SPRIORITYINDEX) And (llCol <> SSTATUSINDEX) And (llCol <> SCLOSEINDEX) Then
                    grdStatus.CellBackColor = LIGHTYELLOW
                ElseIf llCol = SCLOSEINDEX Then
                    If (grdStatus.TextMatrix(llRow, SUSTCODEINDEX) = igUstCode) And (Val(grdStatus.TextMatrix(llRow, SPRIORITYINDEX)) > 1) Then
                        grdStatus.CellBackColor = vbRed
                        grdStatus.CellForeColor = vbWhite
                        grdStatus.TextMatrix(llRow, SCLOSEINDEX) = "X"
                    Else
                        grdStatus.CellBackColor = LIGHTYELLOW
                    End If
                ElseIf llCol = SSTATUSINDEX Then
                    If (grdStatus.TextMatrix(llRow, llCol) <> "Completed") And (grdStatus.TextMatrix(llRow, llCol) <> "Error") Or (grdStatus.TextMatrix(llRow, SPRIORITYINDEX) = "Custom") Then
                        grdStatus.CellBackColor = LIGHTYELLOW
                    Else
                        grdStatus.CellBackColor = LIGHTGREENCOLOR
                    End If
                Else
                    If (grdStatus.TextMatrix(llRow, llCol) = "") Or (grdStatus.TextMatrix(llRow, llCol) = "1") Or (Val(grdStatus.TextMatrix(llRow, SPRIORITYINDEX)) <= 0) Or (grdStatus.TextMatrix(llRow, llCol) = "Custom") Then
                        grdStatus.CellBackColor = LIGHTYELLOW
                    Else
                        If smUserChgPriority = "Y" Then
                            grdStatus.CellBackColor = vbWhite
                        Else
                            If Val(grdStatus.TextMatrix(llRow, SUSTCODEINDEX)) = igUstCode Then
                                grdStatus.CellBackColor = vbWhite
                            Else
                                grdStatus.CellBackColor = LIGHTYELLOW
                            End If
                        End If
                    End If
                End If
            Else
                grdStatus.CellBackColor = LIGHTYELLOW
            End If
        Next llCol
    Next llRow
End Sub

Private Sub mAlertPopulate()

    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slTime As String
    Dim llTime As Long
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim ilShowAuf As Integer
    Dim llDate As Long
    Dim slVehicleName As String
    Dim mItem As ListItem
    Dim tlAuf As AUF
    Dim llRet As Long
    Dim llRow As Long
    Dim llCol As Long

    On Error GoTo ErrHand
    grdAlerts.Row = 0
    llRow = grdAlerts.FixedRows
    
    '8230 add '4'
    For ilLoop = 0 To 4 Step 1
        '7/23/12: Bypass Traffic alters
        If ilLoop <> 2 Then
            ReDim tmAufView(0 To 0) As AUFVIEW
            Select Case ilLoop
                Case 0              'Export Spots
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE (aufType = 'F' or aufType = 'R') AND aufSubType = 'S' AND aufStatus = 'R'"
                Case 1              'Expot ISCI
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE (aufType = 'F' or aufType = 'R') AND aufSubType = 'I' AND aufStatus = 'R'"
                Case 2              'Traffic Logs
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'L' AND (aufSubType = 'S' or aufSubType = 'C') AND aufStatus = 'R'"
                Case 3              'Agreement
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'P' AND aufSubType = 'A' AND aufStatus = 'R'"
                Case 4              'web vendors
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'V' AND aufSubType = 'E' AND aufStatus = 'R'"
            End Select
            Set rstAlert = gSQLSelectCall(SQLAlertQuery)
            Do While Not rstAlert.EOF
                tlAuf.sType = rstAlert!aufType
                tlAuf.sStatus = rstAlert!aufStatus
                tlAuf.sSubType = rstAlert!aufSubType
                tlAuf.iVefCode = rstAlert!aufVefCode
                tlAuf.lCode = rstAlert!aufCode
                If ilLoop = 4 Then
                    'error for queue
                    tlAuf.iCountdown = rstAlert!aufcountdown
                    'error for monitoring
                    tlAuf.lCefCode = rstAlert!aufcefcode
                    'wve code.  I use below to determine which type of error message to show
                    tlAuf.lChfCode = rstAlert!aufChfCode
                    'att code
                    tlAuf.lUlfCode = rstAlert!aufulfcode
                End If
                If IsNull(rstAlert!aufMoWeekDate) Then
                    tlAuf.lMoWeekDate = 0
                ElseIf Not gIsDate(rstAlert!aufMoWeekDate) Then
                    tlAuf.lMoWeekDate = 0
                Else
                    tlAuf.lMoWeekDate = DateValue(gAdjYear(Format$(rstAlert!aufMoWeekDate, sgShowDateForm)))
                End If
                tlAuf.lEnteredDate = DateValue(gAdjYear(Format$(rstAlert!aufEnteredDate, sgShowDateForm)))
                tlAuf.lEnteredTime = gTimeToLong(Format$(rstAlert!aufEnteredTime, sgShowTimeWOSecForm), False)
                slDate = Trim$(Str$(tlAuf.lEnteredDate))
                Do While Len(slDate) < 5
                    slDate = "0" & slDate
                Loop
                slTime = Trim$(Str$(tlAuf.lEnteredTime))
                Do While Len(slTime) < 6
                    slTime = "0" & slTime
                Loop
                tmAufView(UBound(tmAufView)).sKey = slDate & slTime
                tmAufView(UBound(tmAufView)).tAuf = tlAuf
                ReDim Preserve tmAufView(0 To UBound(tmAufView) + 1) As AUFVIEW
                rstAlert.MoveNext
            Loop
            
            If UBound(tmAufView) - 1 > 0 Then
                ArraySortTyp fnAV(tmAufView(), 0), UBound(tmAufView), 1, LenB(tmAufView(0)), 0, LenB(tmAufView(0).sKey), 0
            End If
            
            For ilIndex = 0 To UBound(tmAufView) - 1 Step 1
                'Test if status "R" is still valid.
                tlAuf = tmAufView(ilIndex).tAuf
                If tlAuf.sType = "F" Then   'Final Log Generated
                    'If tlAuf.lMoWeekDate + 6 < DateValue(Format$(gNow(), "m/d/yy")) Then
                    If tlAuf.lMoWeekDate + 6 < DateValue(gAdjYear(Format$(gNow(), "m/d/yy"))) Then
                        ilShowAuf = False
                    Else
                        'ilShowAuf = True
                        If ((tlAuf.sSubType = "I") And (sgExptISCIAlert <> "N")) Or ((tlAuf.sSubType = "S") And (sgExptSpotAlert <> "N")) Then
                            ilShowAuf = True
                        Else
                            ilShowAuf = False
                        End If
                    End If
                ElseIf tlAuf.sType = "R" Then   'Reprint Log generated
                    'If tlAuf.lMoWeekDate + 6 < DateValue(Format$(gNow(), "m/d/yy")) Then
                    If tlAuf.lMoWeekDate + 6 < DateValue(gAdjYear(Format$(gNow(), "m/d/yy"))) Then
                        ilShowAuf = False
                    Else
                        'ilShowAuf = True
                        If ((tlAuf.sSubType = "I") And (sgExptISCIAlert <> "N")) Or ((tlAuf.sSubType = "S") And (sgExptSpotAlert <> "N")) Then
                            ilShowAuf = True
                        Else
                            ilShowAuf = False
                        End If
                    End If
                ElseIf tlAuf.sType = "L" Then
                    If sgTrafLogAlert <> "N" Then
                        ilShowAuf = True
                    Else
                        ilShowAuf = False
                    End If
                ElseIf tlAuf.sType = "P" Then
                    ilShowAuf = True
                '8230
                ElseIf tlAuf.sType = "V" Then
                    ilShowAuf = True
                Else
                    ilShowAuf = False
                End If
                If ilShowAuf Then
                    If llRow >= grdAlerts.Rows Then
                        grdAlerts.AddItem ""
                    End If
                                    
                    Select Case ilLoop
                        Case 0
                            grdAlerts.TextMatrix(llRow, AACTIONINDEX) = "Export" '"Spots"
                        Case 1
                            grdAlerts.TextMatrix(llRow, AACTIONINDEX) = "Export" '"ISCI"
                        Case 2
                            grdAlerts.TextMatrix(llRow, AACTIONINDEX) = "Gen Log"
                        Case 3
                            grdAlerts.TextMatrix(llRow, AACTIONINDEX) = "Chk Agree"
                        '8230
                        Case 4
                            grdAlerts.TextMatrix(llRow, AACTIONINDEX) = "Web Vendor"
                    End Select
                    grdAlerts.TextMatrix(llRow, ACREATIONDATEINDEX) = Format$(tlAuf.lEnteredDate, sgShowDateForm)
                    grdAlerts.TextMatrix(llRow, ACREATIONTIMEINDEX) = Format$(gLongToTime(tlAuf.lEnteredTime), sgShowTimeWOSecForm)
                    '8230
                    If tlAuf.sType = "V" Then
                        If tlAuf.lUlfCode > 0 Then
                            slVehicleName = gVendorInitials(tlAuf.iVefCode) & "-" & mVendorStationVehicle(tlAuf.lUlfCode)
                        Else
                            slVehicleName = gVendorName(tlAuf.iVefCode)
                        End If
                    Else
                        For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                            If tgVehicleInfo(ilVef).iCode = tlAuf.iVefCode Then
                                slVehicleName = tgVehicleInfo(ilVef).sVehicle
                                Exit For
                            End If
                        Next ilVef
                    End If
                    grdAlerts.TextMatrix(llRow, AVEHICLEINDEX) = Trim$(slVehicleName)
                    grdAlerts.TextMatrix(llRow, ADATEINDEX) = Format$(tlAuf.lMoWeekDate, sgShowDateForm)
                    If tlAuf.sType = "F" Then    'Export
                        If (ilLoop = 0) And (tlAuf.sSubType = "S") Then
                            grdAlerts.TextMatrix(llRow, AREASONINDEX) = "Final Log"
                        ElseIf (ilLoop = 1) And (tlAuf.sSubType = "I") Then
                            grdAlerts.TextMatrix(llRow, AREASONINDEX) = "Final Log"
                        End If
                    ElseIf tlAuf.sType = "R" Then    'Export
                        If (ilLoop = 0) And (tlAuf.sSubType = "S") Then
                            grdAlerts.TextMatrix(llRow, AREASONINDEX) = "Reprint Log"
                        ElseIf (ilLoop = 1) And (tlAuf.sSubType = "I") Then
                            grdAlerts.TextMatrix(llRow, AREASONINDEX) = "ISCI"
                        End If
                    ElseIf tlAuf.sType = "L" Then    'Log
                        If tlAuf.sSubType = "C" Then
                            grdAlerts.TextMatrix(llRow, AREASONINDEX) = "Copy Changed"
                        ElseIf tlAuf.sSubType = "S" Then
                            grdAlerts.TextMatrix(llRow, AREASONINDEX) = "Spot Changed"
                        Else
                            grdAlerts.TextMatrix(llRow, AREASONINDEX) = ""
                        End If
                    ElseIf tlAuf.sType = "P" Then    'Agreement
                        If (tlAuf.sSubType = "A") Then
                            grdAlerts.TextMatrix(llRow, AREASONINDEX) = "Program Changed"
                        End If
                        grdAlerts.TextMatrix(llRow, AAUFCODEINDEX) = Trim$(Str$(tlAuf.lCode))
                    '8230
                    ElseIf tlAuf.sType = "V" Then
                        If tlAuf.lChfCode > 0 Then
                            grdAlerts.TextMatrix(llRow, AREASONINDEX) = gVendorIssue(True, tlAuf.iCountdown)
                        Else
                            grdAlerts.TextMatrix(llRow, AREASONINDEX) = gVendorWvmIssue(tlAuf.lCefCode)
                        End If
                        grdAlerts.TextMatrix(llRow, AAUFCODEINDEX) = Trim$(Str$(tlAuf.lCode))
                    End If
                    'Set color
                    grdAlerts.Row = llRow
                    For llCol = AEXPORTINDEX To ADELETEINDEX Step 1
                        grdAlerts.Col = llCol
                        If llCol = AEXPORTINDEX Then
                            If ((tlAuf.sType = "F") Or (tlAuf.sType = "R")) And (tlAuf.sSubType <> "I") Then
                                grdAlerts.CellBackColor = vbWhite
                            Else
                                grdAlerts.CellBackColor = LIGHTYELLOW
                            End If
                        ElseIf llCol <> ADELETEINDEX Then
                            grdAlerts.CellBackColor = LIGHTYELLOW
                        ElseIf (llCol = ADELETEINDEX) And (tlAuf.sType <> "P") Then
                            grdAlerts.CellBackColor = LIGHTYELLOW
                        End If
                    Next llCol
                    llRow = llRow + 1
                Else
                    If (tlAuf.sType = "F") Or (tlAuf.sType = "R") Then
                        SQLAlertQuery = "UPDATE AUF_ALERT_USER SET "
                        SQLAlertQuery = SQLAlertQuery & "aufStatus = 'C'" & ", "
                        SQLAlertQuery = SQLAlertQuery & "aufClearDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                        SQLAlertQuery = SQLAlertQuery & "aufClearTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
                        SQLAlertQuery = SQLAlertQuery & "aufClearUstCode = " & igUstCode & " "
                        SQLAlertQuery = SQLAlertQuery & "WHERE aufCode = " & tlAuf.lCode
                        'cnn.Execute SQLAlertQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLAlertQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand1:
                            gSetMousePointer grdStatus, grdStatus, vbDefault
                            gSetMousePointer grdAlerts, grdStations, vbDefault
                            gHandleError "AffErrorLog.txt", "Export-mAlertPopulate"
                            Exit Sub
                        End If
                        On Error GoTo ErrHand
                    End If
                End If
            Next ilIndex
        End If
    Next ilLoop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Export-mAlertPopulate"
    Exit Sub
End Sub
Private Function mVendorStationVehicle(llAttCode As Long) As String
    '8230
    Dim slAgreementInfo As String
    Dim slSql As String
    
    slAgreementInfo = ""
    If llAttCode = 0 Then
        slAgreementInfo = ""
    Else
        slSql = "Select shttcallletters, vefName from att inner join VEF_Vehicles on attvefcode = vefcode inner join shtt on attshfcode = shttcode where attcode = " & llAttCode
        Set rst = gSQLSelectCall(slSql)
        If Not rst.EOF Then
            slAgreementInfo = Trim$(rst!shttCallLetters) & "\" & Trim$(rst!vefName)
        End If
    End If
    mVendorStationVehicle = slAgreementInfo
End Function
Private Sub mSetExportGridColors()
    Dim llRow As Long
    Dim llCol As Long
    Dim llEht As Long
    
    For llRow = grdExport.FixedRows To grdExport.Rows - 1 Step 1
        For llCol = GENINDEX To CLOSEINDEX Step 1
            grdExport.Row = llRow
            grdExport.Col = llCol
            If grdExport.TextMatrix(llRow, LASTDATEINDEX) <> "" Then
                'If (llCol = EXPORTTYPEINDEX) Or (llCol = EXPORTNAMEINDEX) Or (llCol = WORKDATEINDEX) Or (llCol = LASTDATEINDEX) Or ((grdExport.TextMatrix(llRow, EHTINFOINDEX) = "") And (llCol = CLOSEINDEX)) Then
                If ((llCol <> GENINDEX) And (llCol <> VEHICLEINDEX) And (llCol <> CLOSEINDEX)) Or ((grdExport.TextMatrix(llRow, EHTINFOINDEX) = "") And (llCol = CLOSEINDEX)) Then
                    grdExport.CellBackColor = LIGHTYELLOW
                ElseIf (llCol = VEHICLEINDEX) Then
                    grdExport.CellBackColor = LIGHTGREENCOLOR 'GRAY
                ElseIf (grdExport.TextMatrix(llRow, EHTINFOINDEX) <> "") And (llCol = CLOSEINDEX) Then
                    'grdExport.CellBackColor = vbWhite
                    'Set grdExport.CellPicture = pbcUndo.Picture
                    grdExport.CellBackColor = vbRed
                    grdExport.CellForeColor = vbWhite
                    grdExport.TextMatrix(llRow, CLOSEINDEX) = "X"
                ElseIf (grdExport.TextMatrix(llRow, EHTINFOINDEX) <> "") And (llCol <> GENINDEX) Then
                    grdExport.CellBackColor = LIGHTYELLOW
                Else
                    If (grdExport.CellBackColor <> vbRed) Or (llCol <> GENINDEX) Then
                        grdExport.CellBackColor = vbWhite
                    End If
                End If
            Else
                'If (llCol <> LASTDATEINDEX) And (llCol <> LEADTIMEINDEX) And (llCol <> CYCLEINDEX) Then
                    grdExport.CellBackColor = LIGHTYELLOW
                'Else
                '    grdExport.CellBackColor = vbWhite
                'End If
            End If
        Next llCol
        'Log Color
        grdExport.Col = LOGSTATUSINDEX
        If grdExport.TextMatrix(llRow, EHTINFOINDEX) = "" Then
            If grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "" Then
                lgExportEhtCode = Val(grdExport.TextMatrix(llRow, EHTCODEINDEX))
                For llEht = 0 To UBound(tmEhtStdColor) - 1 Step 1
                    If lgExportEhtCode = tmEhtStdColor(llEht).lEhtCode Then
                        If tmEhtStdColor(llEht).sLogStatus = "G" Then
                            grdExport.CellBackColor = MIDGREENCOLOR
                        ElseIf tmEhtStdColor(llEht).sLogStatus = "R" Then
                            grdExport.CellBackColor = vbRed
                        Else
                            grdExport.CellBackColor = GRAY
                        End If
                        Exit For
                    End If
                Next llEht
            End If
        Else
            llEht = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
            If tmSplitEhtInfo(llEht).sLogStatus = "G" Then
                grdExport.CellBackColor = MIDGREENCOLOR
            ElseIf tmSplitEhtInfo(llEht).sLogStatus = "R" Then
                grdExport.CellBackColor = vbRed
            Else
                grdExport.CellBackColor = GRAY
            End If
        End If
        
        'Copy Color
        grdExport.Col = COPYSTATUSINDEX
        If grdExport.TextMatrix(llRow, EHTINFOINDEX) = "" Then
            If grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "" Then
                lgExportEhtCode = Val(grdExport.TextMatrix(llRow, EHTCODEINDEX))
                For llEht = 0 To UBound(tmEhtStdColor) - 1 Step 1
                    If lgExportEhtCode = tmEhtStdColor(llEht).lEhtCode Then
                        If tmEhtStdColor(llEht).sCopyStatus = "G" Then
                            grdExport.CellBackColor = MIDGREENCOLOR
                        ElseIf tmEhtStdColor(llEht).sCopyStatus = "R" Then
                            grdExport.CellBackColor = vbRed
                        Else
                            grdExport.CellBackColor = GRAY
                        End If
                        Exit For
                    End If
                Next llEht
            End If
        Else
            llEht = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
            If tmSplitEhtInfo(llEht).sCopyStatus = "G" Then
                grdExport.CellBackColor = MIDGREENCOLOR
            ElseIf tmSplitEhtInfo(llEht).sCopyStatus = "R" Then
                grdExport.CellBackColor = vbRed
            Else
                grdExport.CellBackColor = GRAY
            End If
        End If
        
    Next llRow
End Sub

Private Sub mEnableBox()

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llLeft As Long
    Dim llTop As Long
    Dim llRow As Long
    Dim ilPriority As Integer
    
    If (grdExport.Row < grdExport.FixedRows) Or (grdExport.Row >= grdExport.Rows) Or (grdExport.Col < grdExport.FixedCols) Or (grdExport.Col > CLOSEINDEX) Then
        Exit Sub
    End If
    lmEnableRow = grdExport.Row
    lmEnableCol = grdExport.Col
    imCtrlVisible = True
    'pbcArrow.Visible = False
    'pbcArrow.Move grdExport.Left - pbcArrow.Width, grdExport.Top + grdExport.RowPos(grdExport.Row) + (grdExport.RowHeight(grdExport.Row) - pbcArrow.Height) / 2
    'pbcArrow.Visible = True
    llLeft = frcTab(0).Left
    llTop = frcTab(0).Top
    Select Case grdExport.Col
        Case GENINDEX
            'pbcGen.Move llLeft + grdExport.Left + grdExport.ColPos(grdExport.Col) + 30, llTop + grdExport.Top + grdExport.RowPos(grdExport.Row) + 15, grdExport.ColWidth(grdExport.Col) - 30, grdExport.RowHeight(grdExport.Row) - 15
            'If grdExport.TextMatrix(lmEnableRow, lmEnableCol) = "4" Then
            '    smGen = "Y"
            'Else
            '    smGen = "N"
            'End If
            edcDropdown.Move llLeft + grdExport.Left + grdExport.ColPos(grdExport.Col) + 30, llTop + grdExport.Top + grdExport.RowPos(grdExport.Row) + 15, grdExport.ColWidth(grdExport.Col) - 30, grdExport.RowHeight(grdExport.Row) - 15
            If grdExport.CellFontName = "Monotype Sorts" Then
                ilPriority = imNextPriority - 1
                For llRow = grdExport.FixedRows To grdExport.Rows - 1 Step 1
                    If grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "" Then
                        grdExport.Row = llRow
                        If grdExport.CellFontName = "Arial" Then
                            If Val(grdExport.TextMatrix(llRow, GENINDEX)) > ilPriority Then
                                ilPriority = Val(grdExport.TextMatrix(llRow, GENINDEX))
                            End If
                        End If
                    End If
                Next llRow
                grdExport.Row = lmEnableRow
                edcDropdown.Text = ilPriority + 1
            Else
                edcDropdown.Text = grdExport.TextMatrix(lmEnableRow, lmEnableCol)
            End If
        Case LEADTIMEINDEX
            edcDropdown.Move llLeft + grdExport.Left + grdExport.ColPos(grdExport.Col) + 30, llTop + grdExport.Top + grdExport.RowPos(grdExport.Row) + 15, grdExport.ColWidth(grdExport.Col) - 30, grdExport.RowHeight(grdExport.Row) - 15
            edcDropdown.Text = grdExport.TextMatrix(lmEnableRow, lmEnableCol)
        Case CYCLEINDEX
            edcDropdown.Move llLeft + grdExport.Left + grdExport.ColPos(grdExport.Col) + 30, llTop + grdExport.Top + grdExport.RowPos(grdExport.Row) + 15, grdExport.ColWidth(grdExport.Col) - 30, grdExport.RowHeight(grdExport.Row) - 15
            edcDropdown.Text = grdExport.TextMatrix(lmEnableRow, lmEnableCol)
        Case STARTDATEINDEX
            edcDropdown.Move llLeft + grdExport.Left + grdExport.ColPos(grdExport.Col) + 30, llTop + grdExport.Top + grdExport.RowPos(grdExport.Row) + 15, grdExport.ColWidth(grdExport.Col) - 30, grdExport.RowHeight(grdExport.Row) - 15
            edcDropdown.Text = grdExport.TextMatrix(lmEnableRow, lmEnableCol)
        Case ENDDATEINDEX
            edcDropdown.Move llLeft + grdExport.Left + grdExport.ColPos(grdExport.Col) + 30, llTop + grdExport.Top + grdExport.RowPos(grdExport.Row) + 15, grdExport.ColWidth(grdExport.Col) - 30, grdExport.RowHeight(grdExport.Row) - 15
            edcDropdown.Text = grdExport.TextMatrix(lmEnableRow, lmEnableCol)
    End Select
    mSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow()
    Dim slStr As String
    Dim llSvEnableCol As Long
    Dim llSvEnableRow As Long
    Dim blChgVehicleList As Boolean
    Dim slNowDate As String
    Dim slWorkDate As String
    Dim llRow As Long
    Dim blMatchFound As Boolean
    Dim ilPriority As Integer
    Dim ilOrigPriority As Integer
    Dim llEht As Long
    Dim llIndex As Long

    llSvEnableCol = grdExport.Col
    llSvEnableRow = grdExport.Row
    blChgVehicleList = False
    If (lmEnableRow >= grdExport.FixedRows) And (lmEnableRow < grdExport.Rows) Then
        grdExport.Col = lmEnableCol
        grdExport.Row = lmEnableRow
        Select Case lmEnableCol
            Case GENINDEX
                'If smGen = "Y" Then
                '    grdExport.TextMatrix(lmEnableRow, lmEnableCol) = "4"
                'Else
                '    grdExport.TextMatrix(lmEnableRow, lmEnableCol) = " "
                'End If
                If Trim(edcDropdown.Text) = "" Then
                    If grdExport.CellFontName = "Arial" Then
                        'Adjust priorities
                        If Trim$(grdExport.TextMatrix(lmEnableRow, lmEnableCol)) <> "" Then
                            ilPriority = grdExport.TextMatrix(lmEnableRow, lmEnableCol)
                            For llRow = grdExport.FixedRows To grdExport.Rows - 1 Step 1
                                If grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "" Then
                                    grdExport.Row = llRow
                                    If grdExport.CellFontName = "Arial" Then
                                        If Val(grdExport.TextMatrix(llRow, GENINDEX)) >= ilPriority Then
                                            grdExport.TextMatrix(llRow, GENINDEX) = Val(grdExport.TextMatrix(llRow, GENINDEX)) - 1
                                            mSavePriority llRow
                                        End If
                                    End If
                                End If
                            Next llRow
                        End If
                    End If
                    grdExport.Row = lmEnableRow
                    grdExport.CellFontName = "Monotype Sorts"
                    slNowDate = Format(gNow(), "m/d/yy")
                    slWorkDate = grdExport.TextMatrix(lmEnableRow, WORKDATEINDEX)
                    If gDateValue(slWorkDate) <= gDateValue(slNowDate) Then
                        grdExport.TextMatrix(lmEnableRow, GENINDEX) = "4"
                    Else
                        grdExport.TextMatrix(lmEnableRow, GENINDEX) = " "
                    End If
                Else
                    ilOrigPriority = 9999
                    If grdExport.CellFontName = "Arial" Then
                        ilOrigPriority = Val(grdExport.TextMatrix(lmEnableRow, lmEnableCol))
                    End If
                    grdExport.CellFontName = "Arial"
                    ilPriority = edcDropdown.Text
                    blMatchFound = False
                    For llRow = grdExport.FixedRows To grdExport.Rows - 1 Step 1
                        If (grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "") And (llRow <> lmEnableRow) Then
                            grdExport.Row = llRow
                            If grdExport.CellFontName = "Arial" Then
                                If Val(grdExport.TextMatrix(llRow, GENINDEX)) = ilPriority Then
                                    blMatchFound = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next llRow
                    If blMatchFound Then
                        For llRow = grdExport.FixedRows To grdExport.Rows - 1 Step 1
                            If (grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "") And (llRow <> lmEnableRow) Then
                                grdExport.Row = llRow
                                If grdExport.CellFontName = "Arial" Then
                                    If ilPriority < ilOrigPriority Then
                                        If (Val(grdExport.TextMatrix(llRow, GENINDEX)) >= ilPriority) And (Val(grdExport.TextMatrix(llRow, GENINDEX)) < ilOrigPriority) Then
                                            grdExport.TextMatrix(llRow, GENINDEX) = Val(grdExport.TextMatrix(llRow, GENINDEX)) + 1
                                            mSavePriority llRow
                                        End If
                                    Else
                                        If (Val(grdExport.TextMatrix(llRow, GENINDEX)) > ilOrigPriority) And (Val(grdExport.TextMatrix(llRow, GENINDEX)) <= ilPriority) Then
                                            grdExport.TextMatrix(llRow, GENINDEX) = Val(grdExport.TextMatrix(llRow, GENINDEX)) - 1
                                            mSavePriority llRow
                                        End If
                                    End If
                                End If
                            End If
                        Next llRow
                    End If
                    grdExport.Row = lmEnableRow
                    grdExport.TextMatrix(lmEnableRow, lmEnableCol) = ilPriority
                End If
                If grdExport.TextMatrix(lmEnableRow, EHTINFOINDEX) = "" Then
                    lgExportEhtCode = grdExport.TextMatrix(lmEnableRow, EHTCODEINDEX)
                    For llEht = 0 To UBound(tmEhtStdColor) - 1 Step 1
                        If lgExportEhtCode = tmEhtStdColor(llEht).lEhtCode Then
                            tmEhtStdColor(llEht).sGenFont = Left$(grdExport.CellFontName, 1)
                            tmEhtStdColor(llEht).sGen = grdExport.TextMatrix(lmEnableRow, GENINDEX)
                            Exit For
                        End If
                    Next llEht
                Else
                    llIndex = Val(grdExport.TextMatrix(lmEnableRow, EHTINFOINDEX))
                    tmSplitEhtInfo(llIndex).sGenFont = Left$(grdExport.CellFontName, 1)
                    tmSplitEhtInfo(llIndex).sGen = grdExport.TextMatrix(lmEnableRow, GENINDEX)
                End If
            Case LEADTIMEINDEX
                If grdExport.TextMatrix(lmEnableRow, lmEnableCol) <> edcDropdown.Text Then
                    blChgVehicleList = True
                End If
                grdExport.TextMatrix(lmEnableRow, lmEnableCol) = Trim$(edcDropdown.Text)
            Case CYCLEINDEX
                If grdExport.TextMatrix(lmEnableRow, lmEnableCol) <> edcDropdown.Text Then
                    blChgVehicleList = True
                End If
                grdExport.TextMatrix(lmEnableRow, lmEnableCol) = Trim$(edcDropdown.Text)
            Case STARTDATEINDEX
                If grdExport.TextMatrix(lmEnableRow, lmEnableCol) <> edcDropdown.Text Then
                    blChgVehicleList = True
                End If
                grdExport.TextMatrix(lmEnableRow, lmEnableCol) = Trim$(edcDropdown.Text)
            Case ENDDATEINDEX
                If grdExport.TextMatrix(lmEnableRow, lmEnableCol) <> edcDropdown.Text Then
                    blChgVehicleList = True
                End If
                grdExport.TextMatrix(lmEnableRow, lmEnableCol) = Trim$(edcDropdown.Text)
        End Select
        If blChgVehicleList Then
            grdExport.Col = VEHICLEINDEX
            grdExport.Row = lmEnableRow
            grdExport.CellBackColor = LIGHTYELLOW
        End If
    End If
    grdExport.Col = llSvEnableCol
    grdExport.Row = llSvEnableRow
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    'pbcArrow.Visible = False
    edcDropdown.Visible = False
    pbcGen.Visible = False
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer
    Dim llColWidth As Long

    If (grdExport.Row < grdExport.FixedRows) Or (grdExport.Row >= grdExport.Rows) Or (grdExport.Col < grdExport.FixedCols) Or (grdExport.Col > CLOSEINDEX) Then
        Exit Sub
    End If
    imCtrlVisible = True
    Select Case grdExport.Col
        Case GENINDEX
            'pbcGen.Visible = True
            'pbcGen.SetFocus
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case LEADTIMEINDEX
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case CYCLEINDEX
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case STARTDATEINDEX
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case ENDDATEINDEX
            edcDropdown.Visible = True
            edcDropdown.SetFocus
    End Select
End Sub
Private Sub mStatusEnableBox()

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llLeft As Long
    Dim llTop As Long
    Dim llRow As Long
    Dim ilPriority As Integer
    
    If (grdStatus.Row < grdStatus.FixedRows) Or (grdStatus.Row >= grdStatus.Rows) Or (grdStatus.Col < grdStatus.FixedCols) Or (grdStatus.Col > CLOSEINDEX) Then
        Exit Sub
    End If
    lmStatusEnableRow = grdStatus.Row
    lmStatusEnableCol = grdStatus.Col
    imStatusCtrlVisible = True
    'pbcArrow.Visible = False
    'pbcArrow.Move grdStatus.Left - pbcArrow.Width, grdStatus.Top + grdStatus.RowPos(grdStatus.Row) + (grdStatus.RowHeight(grdStatus.Row) - pbcArrow.Height) / 2
    'pbcArrow.Visible = True
    llLeft = frcTab(0).Left
    llTop = frcTab(0).Top
    Select Case grdStatus.Col
        Case SPRIORITYINDEX
            edcStatusDropdown.Move llLeft + grdStatus.Left + grdStatus.ColPos(grdStatus.Col) + 30, llTop + grdStatus.Top + grdStatus.RowPos(grdStatus.Row) + 15, grdStatus.ColWidth(grdStatus.Col) - 30, grdStatus.RowHeight(grdStatus.Row) - 15
            edcStatusDropdown.Text = grdStatus.TextMatrix(lmStatusEnableRow, lmStatusEnableCol)
    End Select
    mStatusSetFocus
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mStatusSetShow()
    Dim slStr As String
    Dim llSvEnableCol As Long
    Dim llSvEnableRow As Long
    Dim llRow As Long
    Dim blMatchFound As Boolean
    Dim ilFdPriority As Integer
    Dim ilPriority As Integer

    llSvEnableCol = grdStatus.Col
    llSvEnableRow = grdStatus.Row
    If (lmStatusEnableRow >= grdStatus.FixedRows) And (lmStatusEnableRow < grdStatus.Rows) Then
        grdStatus.Col = lmStatusEnableCol
        grdStatus.Row = lmStatusEnableRow
        Select Case lmStatusEnableCol
            Case SPRIORITYINDEX
                ilFdPriority = Val(edcStatusDropdown.Text)
                ilPriority = Val(grdStatus.TextMatrix(lmStatusEnableRow, SPRIORITYINDEX))
                For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
                    If grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) <> "" Then
                        grdStatus.Row = llRow
                        If grdStatus.CellBackColor = vbWhite Then
                            If ilFdPriority = Val(grdStatus.TextMatrix(llRow, SPRIORITYINDEX)) Then
                                SQLQuery = "UPDATE eqt_Export_Queue SET "
                                SQLQuery = SQLQuery & "eqtPriority = " & ilPriority
                                SQLQuery = SQLQuery & " WHERE eqtCode = " & grdStatus.TextMatrix(llRow, SEQTCODEINDEX)
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/10/16: Replaced GoSub
                                    'GoSub ErrHand1:
                                    gSetMousePointer grdStatus, grdStatus, vbDefault
                                    gSetMousePointer grdAlerts, grdStations, vbDefault
                                    gHandleError "AffErrorLog.txt", "Export-mStatusSetShow"
                                    Exit Sub
                                End If
                                SQLQuery = "UPDATE eqt_Export_Queue SET "
                                SQLQuery = SQLQuery & "eqtPriority = " & ilFdPriority
                                SQLQuery = SQLQuery & " WHERE eqtCode = " & grdStatus.TextMatrix(lmStatusEnableRow, SEQTCODEINDEX)
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/10/16: Replaced GoSub
                                    'GoSub ErrHand1:
                                    gSetMousePointer grdStatus, grdStatus, vbDefault
                                    gSetMousePointer grdAlerts, grdStations, vbDefault
                                    gHandleError "AffErrorLog.txt", "Export-mStatusSetShow"
                                    Exit Sub
                                End If
                                gSetMousePointer grdExport, grdStatus, vbHourglass
                                gSetMousePointer grdAlerts, grdStations, vbHourglass
                                grdStatus.Redraw = False
                                mStatusPopulate
                                mSetStatusGridColor
                                gSetMousePointer grdExport, grdStatus, vbDefault
                                gSetMousePointer grdAlerts, grdStations, vbDefault
                                grdStatus.Redraw = True
                                Exit For
                            End If
                        End If
                    End If
                Next llRow
        End Select
    End If
    grdStatus.Col = llSvEnableCol
    grdStatus.Row = llSvEnableRow
    lmStatusEnableRow = -1
    lmStatusEnableCol = -1
    imStatusCtrlVisible = False
    edcStatusDropdown.Visible = False
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Export-mStatusSetShow"
    Exit Sub
''ErrHand1:
''    gHandleError "AffErrorLog.txt", "Export-mAdjustPriority"
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mStatusSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer
    Dim llColWidth As Long

    If (grdStatus.Row < grdStatus.FixedRows) Or (grdStatus.Row >= grdStatus.Rows) Or (grdStatus.Col < grdStatus.FixedCols) Or (grdStatus.Col > CLOSEINDEX) Then
        Exit Sub
    End If
    imStatusCtrlVisible = True
    Select Case grdStatus.Col
        Case SPRIORITYINDEX
            edcStatusDropdown.Visible = True
            edcStatusDropdown.SetFocus
    End Select
End Sub

Private Sub mStatusPopulate()
    Dim llRow As Long
    Dim slNowDateMinus1 As String
    Dim llDate As Long
    Dim llTime As Long
    Dim slDate As String
    Dim slTime As String
    Dim ilLastSColSorted As Integer
    Dim ilLastSSort As Integer
    Dim slPriority As String
    
    On Error GoTo ErrHand
    mClearGrid grdStatus
    imNextPriority = 0
    llRow = grdStatus.FixedRows
    slNowDateMinus1 = DateAdd("d", -1, Format(gNow(), "m/d/yy"))
    SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtDateEntered >= '" & Format(slNowDateMinus1, sgSQLDateForm) & "' AND eqtDateEntered <= '" & Format(gNow(), sgSQLDateForm) & "' AND eqtType <> 'T' ORDER BY eqtDateEntered Desc, eqtTimeEntered"
    Set rst_Eqt = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eqt.EOF
        If llRow >= grdStatus.Rows Then
            grdStatus.AddItem ""
        End If
        grdStatus.Row = llRow
        Select Case rst_Eqt!eqtType
            Case "A"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "Aff Logs"
            Case "C"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "C & C"
            Case "D"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "IDC"
            Case "I"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "ISCI"
            Case "R"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "ISCI C/R"
            Case "4"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "RCS 4"
            Case "5"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "RCS 5"
            Case "S"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "StarGuide"
            Case "W"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "Compel"
            Case "X"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "X-Digital"
            Case "1"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "Marketron"
            Case "2"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "Univision"
            Case "3"
                grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "CSI Web"
            Case "P"
                 grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "IPump"
        End Select
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & rst_Eqt!eqtEhtCode
        Set rst_Eht = gSQLSelectCall(SQLQuery)
        If Not rst_Eht.EOF Then
            grdStatus.TextMatrix(llRow, SEXPORTNAMEINDEX) = Trim$(rst_Eht!ehtExportName)
        Else
            grdStatus.TextMatrix(llRow, SEXPORTNAMEINDEX) = ""
        End If
        SQLQuery = "SELECT ustname, ustReportName, ustUserInitials, ustCode FROM Ust Where ustCode = " & rst_Eqt!eqtUstCode
        Set rst_Ust = gSQLSelectCall(SQLQuery)
        If Not rst_Ust.EOF Then
            If Trim$(rst_Ust!ustUserInitials) <> "" Then
                grdStatus.TextMatrix(llRow, SUSERINDEX) = Trim$(rst_Ust!ustUserInitials)
            Else
                If Trim$(rst_Ust!ustReportName) <> "" Then
                    grdStatus.TextMatrix(llRow, SUSERINDEX) = Trim$(rst_Ust!ustReportName)
                Else
                    grdStatus.TextMatrix(llRow, SUSERINDEX) = Trim$(rst_Ust!ustname)
                End If
            End If
            If Trim$(rst_Ust!ustReportName) <> "" Then
                grdStatus.TextMatrix(llRow, SUSERNAMEINDEX) = Trim$(rst_Ust!ustReportName)
            Else
                grdStatus.TextMatrix(llRow, SUSERNAMEINDEX) = Trim$(rst_Ust!ustname)
            End If
        End If
        grdStatus.TextMatrix(llRow, SUSTCODEINDEX) = rst_Eqt!eqtUstCode
        grdStatus.TextMatrix(llRow, STIMEREQUESTINDEX) = Format(rst_Eqt!eqtDateEntered, sgShowDateForm) & " " & Format(rst_Eqt!eqtTimeEntered, sgShowTimeWSecForm)
        grdStatus.TextMatrix(llRow, SSORTINDEX) = ""
        If (rst_Eqt!eqtPriority <= 0) Or (rst_Eht!ehtSubType = "C") Then
            grdStatus.TextMatrix(llRow, SPRIORITYINDEX) = "Custom"
            grdStatus.TextMatrix(llRow, SSORTINDEX) = "C"
        ElseIf rst_Eqt!eqtStatus = "P" Then
            grdStatus.TextMatrix(llRow, SPRIORITYINDEX) = "Processing"
            grdStatus.TextMatrix(llRow, SSORTINDEX) = "A"
            If rst_Eqt!eqtPriority > imNextPriority Then
                imNextPriority = rst_Eqt!eqtPriority
            End If
        ElseIf rst_Eqt!eqtStatus = "R" Then
            grdStatus.TextMatrix(llRow, SPRIORITYINDEX) = rst_Eqt!eqtPriority
            grdStatus.TextMatrix(llRow, SSORTINDEX) = "B"
            If rst_Eqt!eqtPriority > imNextPriority Then
                imNextPriority = rst_Eqt!eqtPriority
            End If
        Else
            grdStatus.TextMatrix(llRow, SPRIORITYINDEX) = ""
        End If
        SQLQuery = "SELECT Count(evtCode) FROM evt_Export_Vehicles WHERE evtEhtCode = " & rst_Eht!ehtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        If Not rst_Evt.EOF Then
            grdStatus.TextMatrix(llRow, SVEHICLEINDEX) = rst_Evt(0).Value
        Else
            grdStatus.TextMatrix(llRow, SVEHICLEINDEX) = ""
        End If
        If rst_Eht!ehtStandardEhtCode > 0 Then
            SQLQuery = "SELECT Count(evtCode) FROM evt_Export_Vehicles WHERE evtEhtCode = " & rst_Eht!ehtStandardEhtCode
            Set rst_Evt = gSQLSelectCall(SQLQuery)
            If Not rst_Evt.EOF Then
                grdStatus.TextMatrix(llRow, SVEHICLEINDEX) = grdStatus.TextMatrix(llRow, SVEHICLEINDEX) & " of " & rst_Evt(0).Value
            End If
        End If
        grdStatus.TextMatrix(llRow, SRESULTFILEINDEX) = ""
        Select Case rst_Eqt!eqtStatus
            Case "R"    'Requested
                grdStatus.TextMatrix(llRow, SSTATUSINDEX) = "Requested"
                grdStatus.TextMatrix(llRow, STIMESTARTEDINDEX) = ""
                grdStatus.TextMatrix(llRow, STIMEENDINDEX) = ""
            Case "P"    'Processing
                grdStatus.TextMatrix(llRow, SSTATUSINDEX) = "Processing"
                grdStatus.TextMatrix(llRow, STIMESTARTEDINDEX) = Format(rst_Eqt!eqtDateStarted, sgShowDateForm) & " " & Format(rst_Eqt!eqtTimeStarted, sgShowTimeWSecForm)
                grdStatus.TextMatrix(llRow, STIMEENDINDEX) = ""
            Case "C"    'Completed
                grdStatus.TextMatrix(llRow, SSTATUSINDEX) = "Completed"
                grdStatus.TextMatrix(llRow, STIMESTARTEDINDEX) = Format(rst_Eqt!eqtDateStarted, sgShowDateForm) & " " & Format(rst_Eqt!eqtTimeStarted, sgShowTimeWSecForm)
                grdStatus.TextMatrix(llRow, STIMEENDINDEX) = Format(rst_Eqt!eqtDateCompleted, sgShowDateForm) & " " & Format(rst_Eqt!eqtTimeCompleted, sgShowTimeWSecForm)
                If grdStatus.TextMatrix(llRow, SSORTINDEX) = "" Then
                    grdStatus.TextMatrix(llRow, SSORTINDEX) = "D"
                End If
                grdStatus.TextMatrix(llRow, SRESULTFILEINDEX) = Trim$(rst_Eqt!eqtResultFile)
           Case "E"    'Error
                grdStatus.TextMatrix(llRow, SSTATUSINDEX) = "Error"
                grdStatus.TextMatrix(llRow, STIMESTARTEDINDEX) = Format(rst_Eqt!eqtDateStarted, sgShowDateForm) & " " & Format(rst_Eqt!eqtTimeStarted, sgShowTimeWSecForm)
                grdStatus.TextMatrix(llRow, STIMEENDINDEX) = Format(rst_Eqt!eqtDateCompleted, sgShowDateForm) & " " & Format(rst_Eqt!eqtTimeCompleted, sgShowTimeWSecForm)
                If grdStatus.TextMatrix(llRow, SSORTINDEX) = "" Then
                    grdStatus.TextMatrix(llRow, SSORTINDEX) = "D"
                End If
                grdStatus.TextMatrix(llRow, SRESULTFILEINDEX) = Trim$(rst_Eqt!eqtResultFile)
        End Select
        grdStatus.TextMatrix(llRow, SEXPORTINFOINDEX) = Format(rst_Eqt!eqtStartDate, sgShowDateForm) & " #" & rst_Eqt!eqtNumberDays
        llDate = gDateValue(Format(rst_Eqt!eqtDateEntered, sgShowDateForm))
        slDate = Trim$(Str$(llDate))
        Do While Len(slDate) < 6
            slDate = "0" & slDate
        Loop
        llTime = gTimeToLong(Format(rst_Eqt!eqtTimeEntered, sgShowTimeWSecForm), True)
        slTime = Trim$(Str$(llTime))
        Do While Len(slTime) < 6
            slTime = "0" & slTime
        Loop
        slPriority = Trim(Str$(rst_Eqt!eqtPriority))
        Do While Len(slPriority) < 3
            slPriority = "0" & slPriority
        Loop
        grdStatus.TextMatrix(llRow, SSORTINDEX) = grdStatus.TextMatrix(llRow, SSORTINDEX) & slDate & slTime & slPriority
        grdStatus.TextMatrix(llRow, SEQTCODEINDEX) = rst_Eqt!eqtCode
        llRow = llRow + 1
        rst_Eqt.MoveNext
    Loop
    ilLastSColSorted = -1
    ilLastSSort = -1
    gGrid_SortByCol grdStatus, SEXPORTTYPEINDEX, SSORTINDEX, ilLastSColSorted, ilLastSSort
    imNextPriority = imNextPriority + 1
    Exit Sub
ErrHand:
    gSetMousePointer grdExport, grdStatus, vbDefault
    gSetMousePointer grdAlerts, grdStations, vbDefault
    gHandleError "AffErrorLog.txt", "ExportQueue-mStatusPopulate"
    grdStatus.Redraw = True
End Sub
Private Sub mCheckServiceStatus()
    Dim ilRet As Integer
    Dim llServiceDate As Long
    Dim llServiceTime As Long
    
    On Error GoTo ErrHand
    SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtType = 'T'"
    Set rst_Eqt = gSQLSelectCall(SQLQuery)
    If Not rst_Eqt.EOF Then
        llServiceDate = gDateValue(Format(rst_Eqt!eqtDateEntered, sgShowDateForm))
        llServiceTime = gTimeToLong(Format(rst_Eqt!eqtTimeEntered, sgShowTimeWSecForm), True)
        If lgLastServiceTime < 0 Then
            lgLastServiceDate = llServiceDate
            lgLastServiceTime = llServiceTime
            igCountTimeNotChanged = 0
            frmExport.BackColor = vbYellow
            Exit Sub
        End If
        If lgLastServiceTime <> llServiceTime Then
            igCountTimeNotChanged = 0
            If lgLastServiceTime + 300 >= llServiceTime Then
                lgLastServiceDate = llServiceDate
                lgLastServiceTime = llServiceTime
                frmExport.BackColor = DARKGREEN
                Exit Sub
            End If
            If (lgLastServiceTime > 85800) And (llServiceTime < 600) Then
                lgLastServiceDate = llServiceDate
                lgLastServiceTime = llServiceTime
                frmExport.BackColor = DARKGREEN
                Exit Sub
            End If
            If lm1970 <> llServiceDate Then
                lgLastServiceDate = llServiceDate
                lgLastServiceTime = llServiceTime
                frmExport.BackColor = DARKGREEN
                Exit Sub
            End If
        Else
            igCountTimeNotChanged = igCountTimeNotChanged + 1
            If igCountTimeNotChanged < 3 Then
                Exit Sub
            End If
        End If
    End If
    frmExport.BackColor = vbRed
    igCountTimeNotChanged = 0
    Exit Sub
ErrHand:
    frmExport.BackColor = vbYellow
    Exit Sub
End Sub

Private Sub TabStrip1_GotFocus()
    mSetShow
    mStatusSetShow
End Sub

Private Sub tmcClock_Timer()
    'mCheckServiceStatus
End Sub

Private Sub mAdjustPriority(ilPriority As Integer)
    On Error GoTo ErrHand
    SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE (eqtStatus = 'R' or eqtStatus = 'P') AND eqtPriority >= " & ilPriority
    Set rst_Eqt = gSQLSelectCall(SQLQuery)
    If Not rst_Eqt.EOF Then
        SQLQuery = "UPDATE eqt_Export_Queue SET "
        SQLQuery = SQLQuery & "eqtPriority = eqtPriority-1 "
        SQLQuery = SQLQuery & "WHERE eqtPriority >= " & ilPriority & " AND eqtStatus = 'R'"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            gSetMousePointer grdStatus, grdStatus, vbDefault
            gSetMousePointer grdAlerts, grdStations, vbDefault
            gHandleError "AffErrorLog.txt", "Export-mAdjustPriority"
            Exit Sub
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Export-mAdjustPriority"
    Exit Sub
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "Export-mAdjustPriority"
End Sub

Private Sub mRemoveOldEqt()
    Dim slDate As String
    Dim blEhtDelete As Boolean
    Dim llEqt As Long
    Dim llEht As Long
    Dim ilPass As Integer
    ReDim llEqtCode(0 To 0) As Long
    ReDim llEhtCode(0 To 0) As Long
    
    On Error GoTo ErrHand
    slDate = Format(gNow(), sgShowDateForm)
    slDate = DateAdd("d", -7, slDate)
    For ilPass = 0 To 1 Step 1
        If ilPass = 0 Then
            SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtDateCompleted > '" & Format("1/1/1970", sgSQLDateForm) & "' AND eqtDateCompleted <= '" & Format(slDate, sgSQLDateForm) & "' AND (eqtStatus = 'C' OR eqtStatus = 'E')"
        Else
            SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtDateEntered <= '" & Format(slDate, sgSQLDateForm) & "' AND eqtStatus = 'P'"
        End If
        Set rst_Eqt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Eqt.EOF
            blEhtDelete = False
            'If custom or Partial, remove EHT and associated files
            'If rst_Eqt!eqtPriority <= 0 Then    'Custom
            SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & rst_Eqt!eqtEhtCode
            Set rst_Eht = gSQLSelectCall(SQLQuery)
            If Not rst_Eht.EOF Then
                If (rst_Eqt!eqtPriority <= 0) Or (rst_Eht!ehtSubType = "C") Then
                    blEhtDelete = True
                Else
                    If rst_Eht!ehtStandardEhtCode > 0 Then
                        blEhtDelete = True
                    End If
                End If
            End If
            If blEhtDelete Then
                '9/8/14: Retain EQT that need to be removed
                llEqtCode(UBound(llEqtCode)) = rst_Eqt!eqtCode
                ReDim Preserve llEqtCode(0 To UBound(llEqtCode) + 1) As Long
                mRemoveEht rst_Eqt!eqtEhtCode
            End If
    
            rst_Eqt.MoveNext
        Loop
        For llEqt = 0 To UBound(llEqtCode) - 1 Step 1
            SQLQuery = "DELETE FROM eqt_Export_Queue WHERE eqtCode = " & llEqtCode(llEqt)
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand1:
                gSetMousePointer grdStatus, grdStatus, vbDefault
                gSetMousePointer grdAlerts, grdStations, vbDefault
                gHandleError "AffErrorLog.txt", "Export-mRemoveOldEqt"
                Exit Sub
            End If
        Next llEqt
    Next ilPass
    'Remove eht without eqt
    SQLQuery = "SELECT * FROM eht_Export_Header Where ehtSubtype = 'C'"
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtEhtCode = " & rst_Eht!ehtCode
        Set rst_Eqt = gSQLSelectCall(SQLQuery)
        If rst_Eqt.EOF Then
            llEhtCode(UBound(llEhtCode)) = rst_Eht!ehtCode
            ReDim Preserve llEhtCode(0 To UBound(llEhtCode) + 1) As Long
        End If
        rst_Eht.MoveNext
    Loop
    For llEht = 0 To UBound(llEhtCode) - 1 Step 1
        mRemoveEht llEhtCode(llEht)
    Next llEht
    Exit Sub
    
ErrHand:
    gHandleError "AffErrorLog.txt", "Export-mRemoveOldEqt"
    Exit Sub
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "Export-mRemoveOldEqt"
'    Return
End Sub

Private Sub mRemoveEht(llEhtCode As Long)
    On Error GoTo ErrHand
    SQLQuery = "DELETE FROM evt_Export_Vehicles"
    SQLQuery = SQLQuery & " WHERE (EvtEhtCode = " & llEhtCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gSetMousePointer grdStatus, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        gHandleError "AffErrorLog.txt", "Export-mRemoveEht"
        Exit Sub
    End If
    SQLQuery = "DELETE FROM est_Export_Station"
    SQLQuery = SQLQuery & " WHERE (EstEhtCode = " & llEhtCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gSetMousePointer grdStatus, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        gHandleError "AffErrorLog.txt", "Export-mRemoveEht"
        Exit Sub
    End If
    SQLQuery = "DELETE FROM ect_Export_Criteria"
    SQLQuery = SQLQuery & " WHERE (EctEhtCode = " & llEhtCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gSetMousePointer grdStatus, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        gHandleError "AffErrorLog.txt", "Export-mRemoveEht"
        Exit Sub
    End If
    SQLQuery = "DELETE FROM eht_Export_Header"
    SQLQuery = SQLQuery & " WHERE (ehtCode = " & llEhtCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gSetMousePointer grdStatus, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        gHandleError "AffErrorLog.txt", "Export-mRemoveEht"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Export-mRemoveEht"
    Exit Sub
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "Export-mRemoveEht"
'    Return
End Sub

Private Function mAddEht(llIndex As Long) As Long
    Dim llStdEhtCode As Long
    Dim llEhtCode As Long
    
    On Error GoTo ErrHand
    llStdEhtCode = tmSplitEhtInfo(llIndex).lStdEhtCode
    SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & llStdEhtCode
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    If Not rst_Eht.EOF Then
        SQLQuery = "Insert Into eht_Export_Header ( "
        SQLQuery = SQLQuery & "ehtCode, "
        SQLQuery = SQLQuery & "ehtExportType, "
        SQLQuery = SQLQuery & "ehtSubType, "
        SQLQuery = SQLQuery & "ehtStandardEhtCode, "
        SQLQuery = SQLQuery & "ehtExportName, "
        SQLQuery = SQLQuery & "ehtUstCode, "
        SQLQuery = SQLQuery & "ehtLDE, "
        SQLQuery = SQLQuery & "ehtLeadTime, "
        SQLQuery = SQLQuery & "ehtCycle, "
        SQLQuery = SQLQuery & "ehtUnused "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "Values ( "
        SQLQuery = SQLQuery & "Replace" & ", "
        SQLQuery = SQLQuery & "'" & gFixQuote(rst_Eht!ehtExportType) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("A") & "', "
        SQLQuery = SQLQuery & llStdEhtCode & ", "
        SQLQuery = SQLQuery & "'" & gFixQuote(rst_Eht!ehtExportName) & "', "
        SQLQuery = SQLQuery & igUstCode & ", "
        SQLQuery = SQLQuery & "'" & Format$(rst_Eht!ehtLDE, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & rst_Eht!ehtLeadTime & ", "
        SQLQuery = SQLQuery & rst_Eht!ehtCycle & ", "
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & ") "
        llEhtCode = gInsertAndReturnCode(SQLQuery, "eht_Export_Header", "ehtCode", "Replace")
        mAddEht = llEhtCode
    Else
        mAddEht = 0
    End If
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Export-mAddEht"
    mAddEht = 0
End Function


Private Function mAddEct(llEhtCode As Long, slLogType As String, slFieldType As String, slFieldName As String, llFieldValue As Long, slFieldString As String) As Long
    Dim llEctCode As Long
    
    On Error GoTo ErrHand
    SQLQuery = "Insert Into ect_Export_Criteria ( "
    SQLQuery = SQLQuery & "ectCode, "
    SQLQuery = SQLQuery & "ectEhtCode, "
    SQLQuery = SQLQuery & "ectLogType, "
    SQLQuery = SQLQuery & "ectFieldType, "
    SQLQuery = SQLQuery & "ectFieldName, "
    SQLQuery = SQLQuery & "ectFieldValue, "
    SQLQuery = SQLQuery & "ectFieldString, "
    SQLQuery = SQLQuery & "ectUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & llEhtCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slLogType) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slFieldType) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slFieldName) & "', "
    SQLQuery = SQLQuery & llFieldValue & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slFieldString) & "', "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    
    llEctCode = gInsertAndReturnCode(SQLQuery, "ect_Export_Criteria", "ectCode", "Replace")

    mAddEct = llEctCode
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Export-mAddEct"
    mAddEct = 0

End Function

Private Function mAddEvt(llEhtCode As Long, ilVefCode As Integer) As Long
    Dim llEvtCode As Long
    
    On Error GoTo ErrHand
    SQLQuery = "Insert Into evt_Export_Vehicles ( "
    SQLQuery = SQLQuery & "evtCode, "
    SQLQuery = SQLQuery & "evtEhtCode, "
    SQLQuery = SQLQuery & "evtVefCode, "
    SQLQuery = SQLQuery & "evtUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & llEhtCode & ", "
    SQLQuery = SQLQuery & ilVefCode & ", "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llEvtCode = gInsertAndReturnCode(SQLQuery, "evt_Export_Vehicles", "evtCode", "Replace")
    mAddEvt = llEvtCode
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Export-mAddEvt"
    mAddEvt = 0

End Function



Private Sub mGetAltered()
    '
    'Gather all splits not being processed (no eqt records)
    'This array is used to replace the standard in the Export grid
    '
    Dim llEhtInfo As Long
    Dim llNext As Long
    Dim llEvtInfo As Long
    Dim llEctInfo As Long
    Dim slStartDate As String
    Dim ilLogAndCopyColor As Integer
    
    On Error GoTo ErrHand
    ReDim tmSplitEhtInfo(0 To 0) As EHTINFO
    ReDim tmSplitEvtInfo(0 To 0) As EVTINFO
    ReDim tmSplitEctInfo(0 To 0) As ECTINFO
    ReDim lmStandardEhtCode(0 To 0) As Long
    SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtStandardEhtCode > 0 "
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        slStartDate = DateAdd("D", 1, Format(rst_Eht!ehtLDE, sgShowDateForm))
        'eqtStatus removed because need to know if split existed during the week
        'SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtEhtCode = " & rst_Eht!ehtCode '& " AND (eqtStatus = 'R' or eqtStatus = 'P')"
        SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtEhtCode = " & rst_Eht!ehtCode & " AND eqtStartdate = '" & Format(slStartDate, sgSQLDateForm) & "'"
        Set rst_Eqt = gSQLSelectCall(SQLQuery)
        If rst_Eqt.EOF Then
            llEhtInfo = UBound(tmSplitEhtInfo)
            tmSplitEhtInfo(llEhtInfo).lEhtCode = rst_Eht!ehtCode
            tmSplitEhtInfo(llEhtInfo).lFirstEvt = -1
            tmSplitEhtInfo(llEhtInfo).lFirstEct = -1
            tmSplitEhtInfo(llEhtInfo).lStdEhtCode = rst_Eht!ehtStandardEhtCode
            tmSplitEhtInfo(llEhtInfo).sLogStatus = "N"
            tmSplitEhtInfo(llEhtInfo).sCopyStatus = "N"
            For ilLogAndCopyColor = 0 To UBound(tmLogAndCopyColor) - 1 Step 1
                If rst_Eht!ehtCode = tmLogAndCopyColor(ilLogAndCopyColor).lEhtCode Then
                    tmSplitEhtInfo(llEhtInfo).sLogStatus = tmLogAndCopyColor(ilLogAndCopyColor).sLogColor
                    tmSplitEhtInfo(llEhtInfo).sCopyStatus = tmLogAndCopyColor(ilLogAndCopyColor).sCopyColor
                    Exit For
                End If
            Next ilLogAndCopyColor
            tmSplitEhtInfo(llEhtInfo).sGenFont = "N"
            tmSplitEhtInfo(llEhtInfo).sGen = ""
            SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & rst_Eht!ehtCode
            Set rst_Evt = gSQLSelectCall(SQLQuery)
            Do While Not rst_Evt.EOF
                If tmSplitEhtInfo(llEhtInfo).lFirstEvt = -1 Then
                    llNext = -1
                Else
                    llNext = tmSplitEhtInfo(llEhtInfo).lFirstEvt
                End If
                llEvtInfo = UBound(tmSplitEvtInfo)
                tmSplitEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
                tmSplitEvtInfo(llEvtInfo).iVefCode = rst_Evt!evtVefCode
                tmSplitEvtInfo(llEvtInfo).lNextEvt = llNext
                ReDim Preserve tmSplitEvtInfo(0 To UBound(tmSplitEvtInfo) + 1) As EVTINFO
                rst_Evt.MoveNext
            Loop
            
            SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
            Set rst_Ect = gSQLSelectCall(SQLQuery)
            Do While Not rst_Ect.EOF
                If tmSplitEhtInfo(llEhtInfo).lFirstEct = -1 Then
                    llNext = -1
                Else
                    llNext = tmSplitEhtInfo(llEhtInfo).lFirstEct
                End If
                llEctInfo = UBound(tmSplitEctInfo)
                tmSplitEhtInfo(llEhtInfo).lFirstEct = llEctInfo
                tmSplitEctInfo(llEctInfo).sLogType = rst_Ect!ectLogType
                tmSplitEctInfo(llEctInfo).sFieldType = rst_Ect!ectFieldType
                tmSplitEctInfo(llEctInfo).sFieldName = rst_Ect!ectFieldName
                tmSplitEctInfo(llEctInfo).lFieldValue = rst_Ect!ectFieldValue
                tmSplitEctInfo(llEctInfo).sFieldString = rst_Ect!ectFieldString
                tmSplitEctInfo(llEctInfo).lNextEct = llNext
                ReDim Preserve tmSplitEctInfo(0 To UBound(tmSplitEctInfo) + 1) As ECTINFO
                rst_Ect.MoveNext
            Loop
            tmSplitEhtInfo(llEhtInfo).blRemoved = False
            tmSplitEhtInfo(llEhtInfo).sGenFont = "N"
            tmSplitEhtInfo(llEhtInfo).sGen = ""
            ReDim Preserve tmSplitEhtInfo(0 To UBound(tmSplitEhtInfo) + 1) As EHTINFO
        Else
            If (rst_Eqt!eqtStatus <> "C") And (rst_Eqt!eqtStatus <> "E") Then
                lmStandardEhtCode(UBound(lmStandardEhtCode)) = rst_Eht!ehtStandardEhtCode
                ReDim Preserve lmStandardEhtCode(0 To UBound(lmStandardEhtCode) + 1) As Long
            End If
        End If
        rst_Eht.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Export-mRemoveEht"
    Exit Sub
ErrHand1:
    gHandleError "AffErrorLog.txt", "Export-mRemoveEht"
    Return
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    mPopListKey
    mCheckTaskMonitor
End Sub

Private Sub tmcStatus_Timer()
    If frmExport.BackColor = DARKGREEN Then
        tmcStatus.Enabled = False
        grdStatus.Redraw = False
        gSetMousePointer grdExport, grdStatus, vbHourglass
        gSetMousePointer grdAlerts, grdStations, vbHourglass
        mStatusPopulate
        mSetStatusGridColor
        gSetMousePointer grdExport, grdStatus, vbDefault
        gSetMousePointer grdAlerts, grdStations, vbDefault
        grdStatus.Redraw = True
        tmcStatus.Enabled = True
    End If
End Sub
Private Sub mPopListKey()
    Dim llMaxWidth As Long
    lbcKey.Clear
    'lbcKey.AddItem "Background Color"
    'lbcKey.AddItem "     Green: Background Export Program Running"
    'lbcKey.AddItem "     Red: Background Export Program Not Running"
    'lbcKey.AddItem "     Yellow: Unable to Determine Status of the "
    'lbcKey.AddItem "             Background Export Program"
    lbcKey.AddItem "Column L: Log"
    lbcKey.AddItem "     Green: Traffic Logs Generated"
    lbcKey.AddItem "     Red: Traffic Logs Not Generated"
    lbcKey.AddItem "     Gray: Traffic Logs Generation Not Checked"
    lbcKey.AddItem "Column C: Copy"
    lbcKey.AddItem "     Green: Copy Assigned"
    lbcKey.AddItem "     Red: Copy Missing"
    lbcKey.AddItem "     Gray: Copy Status Not Checked"
    
    pbcArial.FontBold = False
    pbcArial.FontName = "Arial"
    pbcArial.FontBold = False
    pbcArial.FontSize = 8
    llMaxWidth = (pbcArial.TextWidth("     Gray: Traffic Logs Generation Not Checked")) + 180
    lbcKey.Width = llMaxWidth
    lbcKey.FontBold = False
    lbcKey.FontName = "Arial"
    lbcKey.FontBold = False
    lbcKey.FontSize = 8
    lbcKey.Height = (lbcKey.ListCount) * 225
    lbcKey.Move imcKey.Left, imcKey.Top - lbcKey.Height
End Sub

Private Function mCheckLogs() As Integer
    Dim llRow As Long
    Dim llIndex As Long
    Dim llEvt As Long
    Dim llSvNext As Long
    Dim llCheck As Long
    Dim llEvtInfo As Long
    Dim slLLD As String
    Dim ilLoop As Integer
    Dim blSplit As Boolean
    Dim llNext As Long
    Dim llVef As Long
    Dim llEht As Long
    Dim slCopyColor As String
    Dim slGenFont As String
    Dim slGen As String
    
    blSplit = False
    For llRow = grdExport.FixedRows To grdExport.Rows - 1 Step 1
        If grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "" Then
            If Trim$(grdExport.TextMatrix(llRow, GENINDEX)) <> "" Then
                grdExport.Row = llRow
                grdExport.Col = COPYSTATUSINDEX
                If grdExport.CellBackColor = MIDGREENCOLOR Then
                    slCopyColor = "G"
                ElseIf grdExport.CellBackColor = MIDGREENCOLOR Then
                    slCopyColor = "R"
                Else
                    slCopyColor = "N"
                End If
                grdExport.Col = GENINDEX
                slGenFont = Left$(grdExport.CellFontName, 1)
                slGen = grdExport.Text
                'If grdExport.CellFontName = "Monotype Sorts" Then
                    ReDim ilVefCode(0 To 0) As Integer
                    ReDim ilTVefCode(0 To 0) As Integer
                    If grdExport.TextMatrix(llRow, EHTINFOINDEX) = "" Then
                        'Not split
                        lgExportEhtCode = grdExport.TextMatrix(llRow, EHTCODEINDEX)
                        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & lgExportEhtCode
                        Set rst_Evt = gSQLSelectCall(SQLQuery)
                        Do While Not rst_Evt.EOF
                            SQLQuery = "SELECT vpfLLD, vpfLNoDaysCycle"
                            SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
                            SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & rst_Evt!evtVefCode & ")"
                            
                            Set rst = gSQLSelectCall(SQLQuery)
                            If Not rst.EOF Then
                                If Not IsNull(rst!vpfLLD) Then
                                    If gIsDate(rst!vpfLLD) Then
                                        slLLD = Format$(rst!vpfLLD, "mm/dd/yyyy")
                                        If gDateValue(grdExport.TextMatrix(llRow, ENDDATEINDEX)) <= gDateValue(slLLD) Then
                                            ilVefCode(UBound(ilVefCode)) = rst_Evt!evtVefCode
                                            ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                                        End If
                                    End If
                                End If
                                ilTVefCode(UBound(ilTVefCode)) = rst_Evt!evtVefCode
                                ReDim Preserve ilTVefCode(0 To UBound(ilTVefCode) + 1) As Integer
                            End If
                            rst_Evt.MoveNext
                        Loop
                    Else
                        'Split
                        llIndex = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
                        llEvt = tmSplitEhtInfo(llIndex).lFirstEvt
                        Do While llEvt <> -1
                            SQLQuery = "SELECT vpfLLD, vpfLNoDaysCycle"
                            SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
                            SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & tmSplitEvtInfo(llEvt).iVefCode & ")"
                            
                            Set rst = gSQLSelectCall(SQLQuery)
                            If Not rst.EOF Then
                                If Not IsNull(rst!vpfLLD) Then
                                    If gIsDate(rst!vpfLLD) Then
                                        slLLD = Format$(rst!vpfLLD, "mm/dd/yyyy")
                                        If gDateValue(grdExport.TextMatrix(llRow, ENDDATEINDEX)) <= gDateValue(slLLD) Then
                                            ilVefCode(UBound(ilVefCode)) = tmSplitEvtInfo(llEvt).iVefCode
                                            ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                                        End If
                                    End If
                                End If
                                ilTVefCode(UBound(ilTVefCode)) = tmSplitEvtInfo(llEvt).iVefCode
                                ReDim Preserve ilTVefCode(0 To UBound(ilTVefCode) + 1) As Integer
                            End If
                            llEvt = tmSplitEvtInfo(llEvt).lNextEvt
                        Loop
                    End If
                    If (UBound(ilTVefCode) <> UBound(ilVefCode)) And (UBound(ilVefCode) <> 0) Then
                        blSplit = True
                        mPreSplit llRow
                        
                        lgExportEhtInfoIndex = 0
                        llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
                        Do While llNext <> -1
                            llSvNext = tgEvtInfo(llNext).lNextEvt
                            tgEvtInfo(llNext).lNextEvt = -9999
                            llNext = llSvNext
                        Loop
                        tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt = -1
                        For ilLoop = 0 To UBound(ilVefCode) - 1 Step 1
                            If tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt = -1 Then
                                llNext = -1
                            Else
                                llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
                            End If
                            llEvtInfo = UBound(tgEvtInfo)
                            For llCheck = 0 To UBound(tgEvtInfo) - 1 Step 1
                                If tgEvtInfo(llCheck).lNextEvt = -9999 Then
                                    llEvtInfo = llCheck
                                    Exit For
                                End If
                            Next llCheck
                            tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt = llEvtInfo
                            tgEvtInfo(llEvtInfo).iVefCode = ilVefCode(ilLoop)
                            If tgEvtInfo(llEvtInfo).lNextEvt = -9999 Then
                                tgEvtInfo(llEvtInfo).lNextEvt = llNext
                            Else
                                tgEvtInfo(llEvtInfo).lNextEvt = llNext
                                ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
                            End If
                        Next ilLoop
                        
                        llNext = tgEhtInfo(0).lFirstEvt
                        Do While llNext <> -1
                            For llVef = 0 To UBound(ilTVefCode) - 1 Step 1
                                If ilTVefCode(llVef) = tgEvtInfo(llNext).iVefCode Then
                                    ilTVefCode(llVef) = -1
                                End If
                            Next llVef
                            llNext = tgEvtInfo(llNext).lNextEvt
                        Loop
                        mPostSplit llRow, ilTVefCode(), slGenFont, slGen, LOGSTATUSINDEX, "G", "R", COPYSTATUSINDEX, slCopyColor, slCopyColor
                    Else
                        If grdExport.TextMatrix(llRow, EHTINFOINDEX) = "" Then
                            For llEht = 0 To UBound(tmEhtStdColor) - 1 Step 1
                                If lgExportEhtCode = tmEhtStdColor(llEht).lEhtCode Then
                                    If UBound(ilVefCode) = 0 Then
                                        tmEhtStdColor(llEht).sLogStatus = "R"
                                    Else
                                        tmEhtStdColor(llEht).sLogStatus = "G"
                                    End If
                                    Exit For
                                End If
                            Next llEht
                        Else
                            If UBound(ilVefCode) = 0 Then
                                tmSplitEhtInfo(llIndex).sLogStatus = "R"
                            Else
                                tmSplitEhtInfo(llIndex).sLogStatus = "G"
                            End If
                        End If
                        grdExport.Col = LOGSTATUSINDEX
                        grdExport.CellBackColor = MIDGREENCOLOR
                    End If
                'End If
            End If
        End If
    Next llRow
    If blSplit Then
        mExportPopulate
        mSetExportGridColors
    End If
    mCheckLogs = True
    Exit Function
ErrHand:
    mCheckLogs = False
    gHandleError "AffErrorLog.txt", "Export-mCheckLogs"
End Function

Private Function mCheckCopy() As Integer
    Dim llRow As Long
    Dim llIndex As Long
    Dim llEvt As Long
    Dim llSvNext As Long
    Dim llCheck As Long
    Dim llEvtInfo As Long
    Dim slLLD As String
    Dim ilLoop As Integer
    Dim blSplit As Boolean
    Dim llNext As Long
    Dim llVef As Long
    Dim llEht As Long
    Dim llColor As Long
    Dim slLogColor As String
    Dim slGenFont As String
    Dim slGen As String
    
    blSplit = False
    For llRow = grdExport.FixedRows To grdExport.Rows - 1 Step 1
        If grdExport.TextMatrix(llRow, EXPORTTYPEINDEX) <> "" Then
            If Trim$(grdExport.TextMatrix(llRow, GENINDEX)) <> "" Then
                grdExport.Row = llRow
                grdExport.Col = LOGSTATUSINDEX
                If grdExport.CellBackColor = MIDGREENCOLOR Then
                    slLogColor = "G"
                ElseIf grdExport.CellBackColor = vbRed Then
                    slLogColor = "R"
                Else
                    slLogColor = "N"
                End If
                grdExport.Col = GENINDEX
                slGenFont = Left$(grdExport.CellFontName, 1)
                slGen = grdExport.Text
                'If grdExport.CellFontName = "Monotype Sorts" Then
                    ReDim ilVefCode(0 To 0) As Integer
                    ReDim ilTVefCode(0 To 0) As Integer
                    If grdExport.TextMatrix(llRow, EHTINFOINDEX) = "" Then
                        'Not split
                        lgExportEhtCode = grdExport.TextMatrix(llRow, EHTCODEINDEX)
                        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & lgExportEhtCode
                        Set rst_Evt = gSQLSelectCall(SQLQuery)
                        Do While Not rst_Evt.EOF
                            SQLQuery = "SELECT Count(*)"
                            SQLQuery = SQLQuery + " FROM LST"
                            SQLQuery = SQLQuery + " WHERE (lstLogVefCode =" & rst_Evt!evtVefCode
                            SQLQuery = SQLQuery + " AND lstType = 0"
                            SQLQuery = SQLQuery + " AND lstLogDate >= '" & Format(grdExport.TextMatrix(llRow, STARTDATEINDEX), sgSQLDateForm) & "'"
                            SQLQuery = SQLQuery + " AND lstLogDate <= '" & Format(grdExport.TextMatrix(llRow, ENDDATEINDEX), sgSQLDateForm) & "'" & ")"
                            Set rst = gSQLSelectCall(SQLQuery)
                            If rst(0).Value > 0 Then
                                SQLQuery = "SELECT Count(*)"
                                SQLQuery = SQLQuery + " FROM LST"
                                SQLQuery = SQLQuery + " WHERE (lstLogVefCode =" & rst_Evt!evtVefCode
                                SQLQuery = SQLQuery + " AND lstType = 0"
                                SQLQuery = SQLQuery + " AND lstLogDate >= '" & Format(grdExport.TextMatrix(llRow, STARTDATEINDEX), sgSQLDateForm) & "'"
                                SQLQuery = SQLQuery + " AND lstLogDate <= '" & Format(grdExport.TextMatrix(llRow, ENDDATEINDEX), sgSQLDateForm) & "'"
                                SQLQuery = SQLQuery + " AND RTrim(lstCart) = '' AND RTrim(lstISCI) = ''" & ")"
                                Set rst = gSQLSelectCall(SQLQuery)
                                'If Not rst.EOF Then
                                    If rst(0).Value = 0 Then
                                        ilVefCode(UBound(ilVefCode)) = rst_Evt!evtVefCode
                                        ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                                    Else
                                        ilLoop = ilLoop
                                    End If
                                    ilTVefCode(UBound(ilTVefCode)) = rst_Evt!evtVefCode
                                    ReDim Preserve ilTVefCode(0 To UBound(ilTVefCode) + 1) As Integer
                                'End If
                            Else
                                ilTVefCode(UBound(ilTVefCode)) = rst_Evt!evtVefCode
                                ReDim Preserve ilTVefCode(0 To UBound(ilTVefCode) + 1) As Integer
                            End If
                            rst_Evt.MoveNext
                        Loop
                    Else
                        'Split
                        llIndex = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
                        llEvt = tmSplitEhtInfo(llIndex).lFirstEvt
                        Do While llEvt <> -1
                            SQLQuery = "SELECT Count(*)"
                            SQLQuery = SQLQuery + " FROM LST"
                            SQLQuery = SQLQuery + " WHERE (lstLogVefCode =" & tmSplitEvtInfo(llEvt).iVefCode
                            SQLQuery = SQLQuery + " AND lstType = 0"
                            SQLQuery = SQLQuery + " AND lstLogDate >= '" & Format(grdExport.TextMatrix(llRow, STARTDATEINDEX), sgSQLDateForm) & "'"
                            SQLQuery = SQLQuery + " AND lstLogDate <= '" & Format(grdExport.TextMatrix(llRow, ENDDATEINDEX), sgSQLDateForm) & "'" & ")"
                            Set rst = gSQLSelectCall(SQLQuery)
                            If rst(0).Value > 0 Then
                                SQLQuery = "SELECT Count(*)"
                                SQLQuery = SQLQuery + " FROM LST"
                                SQLQuery = SQLQuery + " WHERE (lstLogVefCode =" & tmSplitEvtInfo(llEvt).iVefCode
                                SQLQuery = SQLQuery + " AND lstType = 0"
                                SQLQuery = SQLQuery + " AND lstLogDate >= '" & Format(grdExport.TextMatrix(llRow, STARTDATEINDEX), sgSQLDateForm) & "'"
                                SQLQuery = SQLQuery + " AND lstLogDate <= '" & Format(grdExport.TextMatrix(llRow, ENDDATEINDEX), sgSQLDateForm) & "'"
                                SQLQuery = SQLQuery + " AND RTrim(lstCart) = '' AND RTrim(lstISCI) = ''" & ")"
                                Set rst = gSQLSelectCall(SQLQuery)
                                'If Not rst.EOF Then
                                    If rst(0).Value = 0 Then
                                        ilVefCode(UBound(ilVefCode)) = tmSplitEvtInfo(llEvt).iVefCode
                                        ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                                    Else
                                        ilLoop = ilLoop
                                    End If
                                    ilTVefCode(UBound(ilTVefCode)) = tmSplitEvtInfo(llEvt).iVefCode
                                    ReDim Preserve ilTVefCode(0 To UBound(ilTVefCode) + 1) As Integer
                                'End If
                            Else
                                ilTVefCode(UBound(ilTVefCode)) = tmSplitEvtInfo(llEvt).iVefCode
                                ReDim Preserve ilTVefCode(0 To UBound(ilTVefCode) + 1) As Integer
                            End If
                            llEvt = tmSplitEvtInfo(llEvt).lNextEvt
                        Loop
                    End If
                    If (UBound(ilTVefCode) <> UBound(ilVefCode)) And (UBound(ilVefCode) <> 0) Then
                        blSplit = True
                        mPreSplit llRow
                        
                        lgExportEhtInfoIndex = 0
                        llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
                        Do While llNext <> -1
                            llSvNext = tgEvtInfo(llNext).lNextEvt
                            tgEvtInfo(llNext).lNextEvt = -9999
                            llNext = llSvNext
                        Loop
                        tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt = -1
                        For ilLoop = 0 To UBound(ilVefCode) - 1 Step 1
                            If tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt = -1 Then
                                llNext = -1
                            Else
                                llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
                            End If
                            llEvtInfo = UBound(tgEvtInfo)
                            For llCheck = 0 To UBound(tgEvtInfo) - 1 Step 1
                                If tgEvtInfo(llCheck).lNextEvt = -9999 Then
                                    llEvtInfo = llCheck
                                    Exit For
                                End If
                            Next llCheck
                            tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt = llEvtInfo
                            tgEvtInfo(llEvtInfo).iVefCode = ilVefCode(ilLoop)
                            If tgEvtInfo(llEvtInfo).lNextEvt = -9999 Then
                                tgEvtInfo(llEvtInfo).lNextEvt = llNext
                            Else
                                tgEvtInfo(llEvtInfo).lNextEvt = llNext
                                ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
                            End If
                        Next ilLoop
                        
                        llNext = tgEhtInfo(0).lFirstEvt
                        Do While llNext <> -1
                            For llVef = 0 To UBound(ilTVefCode) - 1 Step 1
                                If ilTVefCode(llVef) = tgEvtInfo(llNext).iVefCode Then
                                    ilTVefCode(llVef) = -1
                                End If
                            Next llVef
                            llNext = tgEvtInfo(llNext).lNextEvt
                        Loop
                        mPostSplit llRow, ilTVefCode(), slGenFont, slGen, LOGSTATUSINDEX, slLogColor, slLogColor, COPYSTATUSINDEX, "G", "R"
                    Else
                        llColor = GRAY
                        If grdExport.TextMatrix(llRow, EHTINFOINDEX) = "" Then
                            For llEht = 0 To UBound(tmEhtStdColor) - 1 Step 1
                                If lgExportEhtCode = tmEhtStdColor(llEht).lEhtCode Then
                                    If UBound(ilVefCode) = 0 Then
                                        tmEhtStdColor(llEht).sCopyStatus = "R"
                                        tmEhtStdColor(llEht).sGenFont = "N"
                                        tmEhtStdColor(llEht).sGen = ""
                                        llColor = vbRed
                                    Else
                                        tmEhtStdColor(llEht).sCopyStatus = "G"
                                        tmEhtStdColor(llEht).sGenFont = slGenFont
                                        tmEhtStdColor(llEht).sGen = slGen
                                        llColor = MIDGREENCOLOR
                                    End If
                                    Exit For
                                End If
                            Next llEht
                        Else
                            If UBound(ilVefCode) = 0 Then
                                tmSplitEhtInfo(llIndex).sCopyStatus = "R"
                                tmSplitEhtInfo(llIndex).sGenFont = "N"
                                tmSplitEhtInfo(llIndex).sGen = ""
                                llColor = vbRed
                            Else
                                tmSplitEhtInfo(llIndex).sCopyStatus = "G"
                                tmSplitEhtInfo(llIndex).sGenFont = slGenFont
                                tmSplitEhtInfo(llIndex).sGen = slGen
                                llColor = MIDGREENCOLOR
                            End If
                        End If
                        grdExport.Col = COPYSTATUSINDEX
                        grdExport.CellBackColor = llColor
                        If llColor = vbRed Then
                            tmSplitEhtInfo(llIndex).sGenFont = "N"
                            tmSplitEhtInfo(llIndex).sGen = ""
                            grdExport.Text = ""
                        End If
                    End If
                'End If
            End If
        End If
    Next llRow
    If blSplit Then
        mExportPopulate
        mSetExportGridColors
    End If
    mCheckCopy = True
    Exit Function
ErrHand:
    mCheckCopy = False
    gHandleError "AffErrorLog.txt", "Export-mCheckCopy"
End Function

Private Sub mPreSplit(llRow As Long)
    Dim ilFound As Integer
    Dim llNext As Long
    Dim llEvtNext As Long
    Dim llEctNext As Long
    Dim llEhtInfo As Long
    Dim llEvtInfo As Long
    Dim llEctInfo As Long

    ReDim tgEhtInfo(0 To 0) As EHTINFO
    ReDim tgEvtInfo(0 To 0) As EVTINFO
    ReDim tgEctInfo(0 To 0) As ECTINFO
    sgExportTypeChar = grdExport.TextMatrix(llRow, EHTTYPECHARINDEX)
    sgExportName = grdExport.TextMatrix(llRow, EXPORTNAMEINDEX)
    lgExportEhtCode = Val(grdExport.TextMatrix(llRow, EHTCODEINDEX))
    If grdExport.TextMatrix(llRow, EHTINFOINDEX) = "" Then
        ilFound = False
        lgExportEhtInfoIndex = UBound(tgEhtInfo)
        llEhtInfo = lgExportEhtInfoIndex
        tgEhtInfo(llEhtInfo).lEhtCode = 0
        tgEhtInfo(llEhtInfo).lFirstEvt = -1
        tgEhtInfo(llEhtInfo).lFirstEct = -1
        tgEhtInfo(llEhtInfo).lStdEhtCode = lgExportEhtCode
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & lgExportEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            If tgEhtInfo(llEhtInfo).lFirstEvt = -1 Then
                llNext = -1
            Else
                llNext = tgEhtInfo(llEhtInfo).lFirstEvt
            End If
            llEvtInfo = UBound(tgEvtInfo)
            tgEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
            tgEvtInfo(llEvtInfo).iVefCode = rst_Evt!evtVefCode
            tgEvtInfo(llEvtInfo).lNextEvt = llNext
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
        
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & lgExportEhtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            If tgEhtInfo(llEhtInfo).lFirstEct = -1 Then
                llNext = -1
            Else
                llNext = tgEhtInfo(llEhtInfo).lFirstEct
            End If
            llEctInfo = UBound(tgEctInfo)
            tgEhtInfo(llEhtInfo).lFirstEct = llEctInfo
            tgEctInfo(llEctInfo).sLogType = rst_Ect!ectLogType
            tgEctInfo(llEctInfo).sFieldType = rst_Ect!ectFieldType
            tgEctInfo(llEctInfo).sFieldName = rst_Ect!ectFieldName
            tgEctInfo(llEctInfo).lFieldValue = rst_Ect!ectFieldValue
            tgEctInfo(llEctInfo).sFieldString = rst_Ect!ectFieldString
            tgEctInfo(llEctInfo).lNextEct = llNext
            ReDim Preserve tgEctInfo(0 To UBound(tgEctInfo) + 1) As ECTINFO
            rst_Ect.MoveNext
        Loop
        ReDim Preserve tgEhtInfo(0 To UBound(tgEhtInfo) + 1) As EHTINFO
    Else
        ilFound = True
        lgExportEhtInfoIndex = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
        
        llEhtInfo = UBound(tgEhtInfo)
        tgEhtInfo(llEhtInfo).lEhtCode = tmSplitEhtInfo(lgExportEhtInfoIndex).lEhtCode
        tgEhtInfo(llEhtInfo).lFirstEvt = -1
        tgEhtInfo(llEhtInfo).lFirstEct = -1
        tgEhtInfo(llEhtInfo).lStdEhtCode = lgExportEhtCode
        llEvtNext = tmSplitEhtInfo(lgExportEhtInfoIndex).lFirstEvt
        Do While llEvtNext <> -1
            If tgEhtInfo(llEhtInfo).lFirstEvt = -1 Then
                llNext = -1
            Else
                llNext = tgEhtInfo(llEhtInfo).lFirstEvt
            End If
            llEvtInfo = UBound(tgEvtInfo)
            tgEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
            tgEvtInfo(llEvtInfo).iVefCode = tmSplitEvtInfo(llEvtNext).iVefCode
            tgEvtInfo(llEvtInfo).lNextEvt = llNext
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            llEvtNext = tmSplitEvtInfo(llEvtNext).lNextEvt
        Loop
        llEctNext = tmSplitEhtInfo(lgExportEhtInfoIndex).lFirstEct
        Do While llEctNext <> -1
            If tgEhtInfo(llEhtInfo).lFirstEct = -1 Then
                llNext = -1
            Else
                llNext = tgEhtInfo(llEhtInfo).lFirstEct
            End If
            llEctInfo = UBound(tgEctInfo)
            tgEhtInfo(llEhtInfo).lFirstEct = llEctInfo
            tgEctInfo(llEctInfo).sLogType = tmSplitEctInfo(llEctNext).sLogType
            tgEctInfo(llEctInfo).sFieldType = tmSplitEctInfo(llEctNext).sFieldType
            tgEctInfo(llEctInfo).sFieldName = tmSplitEctInfo(llEctNext).sFieldName
            tgEctInfo(llEctInfo).lFieldValue = tmSplitEctInfo(llEctNext).lFieldValue
            tgEctInfo(llEctInfo).sFieldString = tmSplitEctInfo(llEctNext).sFieldString
            tgEctInfo(llEctInfo).lNextEct = llNext
            ReDim Preserve tgEctInfo(0 To UBound(tgEctInfo) + 1) As ECTINFO
            llEctNext = tmSplitEctInfo(llEctNext).lNextEct
        Loop
        ReDim Preserve tgEhtInfo(0 To UBound(tgEhtInfo) + 1) As EHTINFO
    End If
End Sub

Private Sub mPostSplit(llRow As Long, ilVefCode() As Integer, slGenFont As String, slGen As String, ilLogCol As Integer, slPass0LogColor As String, slPass1LogColor As String, ilCopyCol As Integer, slPass0CopyColor As String, slPass1CopyColor As String)
    Dim llEvtNext As Long
    Dim llVef As Long
    Dim ilPass As Integer
    Dim llNext As Long
    Dim ilSave As Integer
    Dim llEhtInfo As Long
    Dim llEvtInfo As Long
    Dim llCopyEct As Long
    Dim llEctInfo As Long
    Dim ilFound As Integer
    
    llEvtNext = tgEhtInfo(0).lFirstEvt
    llVef = 0
    If grdExport.TextMatrix(llRow, EHTINFOINDEX) = "" Then
        ilFound = False
    Else
        ilFound = True
    End If
    For ilPass = 0 To 1 Step 1
        If (ilPass = 0) And ilFound Then
            llEhtInfo = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
        Else
            llEhtInfo = UBound(tmSplitEhtInfo)
            tmSplitEhtInfo(llEhtInfo).lEhtCode = 0
        End If
        tmSplitEhtInfo(llEhtInfo).lFirstEvt = -1
        tmSplitEhtInfo(llEhtInfo).lFirstEct = -1
        tmSplitEhtInfo(llEhtInfo).lStdEhtCode = lgExportEhtCode
        Do
            If tmSplitEhtInfo(llEhtInfo).lFirstEvt = -1 Then
                llNext = -1
            Else
                llNext = tmSplitEhtInfo(llEhtInfo).lFirstEvt
            End If
            ilSave = False
            llEvtInfo = UBound(tmSplitEvtInfo)
            If ilPass = 0 Then
                If llEvtNext <> -1 Then
                    tmSplitEvtInfo(llEvtInfo).iVefCode = tgEvtInfo(llEvtNext).iVefCode
                    llEvtNext = tgEvtInfo(llEvtNext).lNextEvt
                    ilSave = True
                End If
            Else
                Do While llVef < UBound(ilVefCode)
                    If ilVefCode(llVef) > 0 Then
                        tmSplitEvtInfo(llEvtInfo).iVefCode = ilVefCode(llVef)
                        llVef = llVef + 1
                        ilSave = True
                        Exit Do
                    End If
                    llVef = llVef + 1
                Loop
            End If
            If ilSave Then
                tmSplitEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
                tmSplitEvtInfo(llEvtInfo).lNextEvt = llNext
                If ilPass = 0 Then
                    tmSplitEhtInfo(llEhtInfo).sGenFont = slGenFont
                    tmSplitEhtInfo(llEhtInfo).sGen = slGen
                End If
                ReDim Preserve tmSplitEvtInfo(0 To UBound(tmSplitEvtInfo) + 1) As EVTINFO
            Else
                Exit Do
            End If
        Loop
        llCopyEct = tgEhtInfo(0).lFirstEct
        Do While llCopyEct <> -1
            If tmSplitEhtInfo(llEhtInfo).lFirstEct = -1 Then
                llNext = -1
            Else
                llNext = tmSplitEhtInfo(llEhtInfo).lFirstEct
            End If
            llEctInfo = UBound(tmSplitEctInfo)
            tmSplitEhtInfo(llEhtInfo).lFirstEct = llEctInfo
            tmSplitEctInfo(llEctInfo).sLogType = tgEctInfo(llCopyEct).sLogType
            tmSplitEctInfo(llEctInfo).sFieldType = tgEctInfo(llCopyEct).sFieldType
            tmSplitEctInfo(llEctInfo).sFieldName = tgEctInfo(llCopyEct).sFieldName
            tmSplitEctInfo(llEctInfo).lFieldValue = tgEctInfo(llCopyEct).lFieldValue
            tmSplitEctInfo(llEctInfo).sFieldString = tgEctInfo(llCopyEct).sFieldString
            tmSplitEctInfo(llEctInfo).lNextEct = llNext
            ReDim Preserve tmSplitEctInfo(0 To UBound(tmSplitEctInfo) + 1) As ECTINFO
            llCopyEct = tgEctInfo(llCopyEct).lNextEct
        Loop
        If ilLogCol = LOGSTATUSINDEX Then
            tmSplitEhtInfo(llEhtInfo).sLogStatus = "N"
            If ilPass = 0 Then
                tmSplitEhtInfo(llEhtInfo).sLogStatus = slPass0LogColor
            ElseIf ilPass = 1 Then
                tmSplitEhtInfo(llEhtInfo).sLogStatus = slPass1LogColor
            End If
        End If
        If ilCopyCol = COPYSTATUSINDEX Then
            tmSplitEhtInfo(llEhtInfo).sCopyStatus = "N"
            If ilPass = 0 Then
                tmSplitEhtInfo(llEhtInfo).sCopyStatus = slPass0CopyColor
            ElseIf ilPass = 1 Then
                tmSplitEhtInfo(llEhtInfo).sCopyStatus = slPass1CopyColor
            End If
        End If
        If (ilPass = 1) Or (Not ilFound) Then
            tmSplitEhtInfo(llEhtInfo).blRemoved = False
            tmSplitEhtInfo(llEhtInfo).sGenFont = "N"
            tmSplitEhtInfo(llEhtInfo).sGen = ""
            ReDim Preserve tmSplitEhtInfo(0 To UBound(tmSplitEhtInfo) + 1) As EHTINFO
        End If
        If ilPass = 0 Then
            tmSplitEhtInfo(llEhtInfo).sGenFont = slGenFont
            tmSplitEhtInfo(llEhtInfo).sGen = slGen
        End If
    Next ilPass

End Sub

Private Sub mSavePriority(llRow As Long)
    Dim llSvRow As Long
    Dim llSvCol As Long
    Dim llEht As Long
    Dim llIndex As Integer
    
    llSvRow = grdExport.Row
    llSvCol = grdExport.Col
    grdExport.Row = llRow
    grdExport.Col = GENINDEX
    If grdExport.TextMatrix(llRow, EHTINFOINDEX) = "" Then
        lgExportEhtCode = grdExport.TextMatrix(llRow, EHTCODEINDEX)
        For llEht = 0 To UBound(tmEhtStdColor) - 1 Step 1
            If lgExportEhtCode = tmEhtStdColor(llEht).lEhtCode Then
                tmEhtStdColor(llEht).sGenFont = Left$(grdExport.CellFontName, 1)
                tmEhtStdColor(llEht).sGen = grdExport.TextMatrix(llRow, GENINDEX)
                Exit For
            End If
        Next llEht
    Else
        llIndex = Val(grdExport.TextMatrix(llRow, EHTINFOINDEX))
        tmSplitEhtInfo(llIndex).sGenFont = Left$(grdExport.CellFontName, 1)
        tmSplitEhtInfo(llIndex).sGen = grdExport.TextMatrix(llRow, GENINDEX)
    End If
    grdExport.Row = llSvRow
    grdExport.Col = llSvCol
End Sub

Private Sub mCheckTaskMonitor()
    Dim blAEQFound As Boolean
    Dim ilTask As Integer

    blAEQFound = False
    For ilTask = 0 To UBound(tgTaskInfo) Step 1
        If Trim$(tgTaskInfo(ilTask).sTaskCode) = "AEQ" Then
            If tgTaskInfo(ilTask).iMenuIndex > 0 Then
                blAEQFound = True
                If tgTaskInfo(ilTask).lColor <> DARKGREEN Then
                    MsgBox " Please verify that the Affiliate Export Queue Application is Running", vbInformation + vbOKOnly
                End If
            End If
            Exit For
        End If
    Next ilTask
    If blAEQFound = False Then
        MsgBox "Affiliate Export Queue: Please Contact Counterpoint Service to get help setting up the Task Monitor and/or the Application", vbInformation + vbOKOnly
    End If

End Sub

