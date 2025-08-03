VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExpMatrix 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5100
   ClientLeft      =   825
   ClientTop       =   2400
   ClientWidth     =   10020
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
   ScaleHeight     =   5100
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ckcInclCmmts 
      Caption         =   "Include Digital Avg Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   2160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CheckBox ckcPacingRange 
      Caption         =   "Date range Pacing"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox edcPacingEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox edcPacing 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame frcAmazon 
      Height          =   1455
      Left            =   8400
      TabIndex        =   34
      Top             =   4440
      Visible         =   0   'False
      Width           =   9735
      Begin VB.TextBox edcAmazonSubfolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   39
         ToolTipText     =   "(Optional) Amazon Web Bucket Subfolder Name.   Example: Counterpoint"
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox ckcKeepLocalFile 
         Caption         =   "Keep Local File"
         Height          =   195
         Left            =   6000
         TabIndex        =   42
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox edcBucketName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox edcAccessKey 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6000
         PasswordChar    =   "*"
         TabIndex        =   40
         ToolTipText     =   "The Access Key Assigned by AWS"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox edcPrivateKey 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6000
         PasswordChar    =   "*"
         TabIndex        =   41
         ToolTipText     =   "The Private Key Assigned by AWS"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox edcRegion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         ToolTipText     =   "Region/Endpoint - Example: USEast1, USEast2, USWest1 or USWest2"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Folder (optional)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lacExportFilename 
         Caption         =   "lacExportFilename"
         Height          =   255
         Left            =   7800
         TabIndex        =   47
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "BucketName"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "AccessKey"
         Height          =   255
         Left            =   4800
         TabIndex        =   45
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "PrivateKey"
         Height          =   255
         Left            =   4800
         TabIndex        =   44
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Region"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CheckBox ckcAmazon 
      Caption         =   "Upload to Amazon Web bucket"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   4560
      Width           =   3135
   End
   Begin VB.PictureBox plcTo 
      Height          =   375
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   5445
      TabIndex        =   35
      Top             =   3240
      Width           =   5505
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   30
         Width           =   5505
      End
   End
   Begin VB.CommandButton cmcTo 
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
      Left            =   7200
      TabIndex        =   18
      Top             =   3240
      Width           =   1485
   End
   Begin VB.CheckBox ckcInclAdj 
      Caption         =   "Include Adjustments"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1860
      Width           =   2295
   End
   Begin VB.CheckBox ckcInclMissed 
      Caption         =   "Include Missed Spots"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   9525
      Top             =   3600
   End
   Begin VB.ListBox lbcVehicle 
      Height          =   2205
      ItemData        =   "ExpMatrix.frx":0000
      Left            =   6120
      List            =   "ExpMatrix.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   16
      Top             =   720
      Width           =   3735
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Vehicles"
      Height          =   195
      Left            =   6120
      TabIndex        =   15
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox PlcNetBy 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   5895
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   840
      Width           =   5895
      Begin VB.OptionButton rbcNetBy 
         Caption         =   "Net"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   32
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton rbcNetBy 
         Caption         =   "T-Net"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   31
         Top             =   0
         Width           =   1080
      End
   End
   Begin VB.PictureBox plcMonthBy 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      ScaleHeight     =   360
      ScaleWidth      =   5895
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5895
      Begin VB.OptionButton rbcMonthBy 
         Caption         =   "Bill Method"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton rbcMonthBy 
         Caption         =   "Calendar Contract"
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   33
         Top             =   0
         Width           =   1935
      End
      Begin VB.OptionButton rbcMonthBy 
         Caption         =   "Calendar Spots"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
      Begin VB.OptionButton rbcMonthBy 
         Caption         =   "Standard"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.TextBox edcContract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   14
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9120
      Top             =   3600
   End
   Begin VB.CheckBox ckcNTR 
      Caption         =   "Include NTR Revenue"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox edcNoMonths 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   4920
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "13"
      Top             =   1200
      Width           =   360
   End
   Begin VB.TextBox edcYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1200
      Width           =   720
   End
   Begin VB.TextBox edcMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1200
      Width           =   600
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   8280
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
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
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   2355
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2355
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3615
      Visible         =   0   'False
      Width           =   525
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
      Left            =   3680
      TabIndex        =   20
      Top             =   4560
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
      Left            =   5380
      TabIndex        =   21
      Top             =   4560
      Width           =   1050
   End
   Begin VB.Label lacPacingRange 
      Alignment       =   2  'Center
      Caption         =   "to"
      Height          =   255
      Left            =   2280
      TabIndex        =   51
      Top             =   2550
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lacPacing 
      Caption         =   "Pacing Date"
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   2550
      Width           =   1095
   End
   Begin VB.Label lacSaveIn 
      Appearance      =   0  'Flat
      Caption         =   "Save In"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   360
      TabIndex        =   36
      Top             =   3270
      Width           =   810
   End
   Begin VB.Label lacContract 
      Appearance      =   0  'Flat
      Caption         =   "Contract #"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   29
      Top             =   2940
      Width           =   1065
   End
   Begin VB.Label lacNoMonths 
      Appearance      =   0  'Flat
      Caption         =   "# months"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3960
      TabIndex        =   28
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   480
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Label lacMonth 
      Appearance      =   0  'Flat
      Caption         =   "Start Month"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1230
      Width           =   1185
   End
   Begin VB.Label lacStartYear 
      Appearance      =   0  'Flat
      Caption         =   "Start Year"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2040
      TabIndex        =   23
      Top             =   1230
      Width           =   1035
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8760
      Top             =   3600
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   2520
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   7335
   End
End
Attribute VB_Name = "ExpMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ExpMatrix.frm on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: ExpMatrix.Frm
'
' Release: 5.1
'
' Description:
'   This file contains the Export Matrix input screen
'   7-9-15 Tableau STd & Cal export:  same as Matrix format
Option Explicit
Option Compare Text
Dim hmMsg As Integer
Dim hmMatrix As Integer

Dim smExportName As String
Dim imFirstActivate As Integer
Dim lmCntrNo As Long    'for debugging purposes to filter a single contract

Dim imFirstTime As Integer
Dim tmChfAdvtExt() As CHFADVTEXT

Dim lmProject(0 To 24) As Long          'projection $, max 2 years, index zero ignored
'2-28-14 implement tnet option
Dim lmAcquisition(0 To 24) As Long      'Acquisition $, max 2 years, index zero ignored

Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF

Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF

Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF

Dim hmAgf As Integer            'Agency file handle
Dim imAgfRecLen As Integer      'AGF record length
Dim tmAgfSrchKey As INTKEY0     'AGF key image
Dim tmAgf As AGF

Dim tmRvf As RVF

Dim hmSof As Integer            'Office file handle
Dim imSofRecLen As Integer      'SOF record length
Dim tmSof() As SOF

Dim hmSlf As Integer            'Salesperson file handle

Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnfSrchKey As INTKEY0
Dim tmMnf As MNF
Dim tmMnfSS() As MNF                    'array of Sales Sources MNF
Dim tmMnfGroups() As MNF

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer

Dim hmVff As Integer            'Vfhicle features file handle
Dim tmVff As VFF                'VfF record image
Dim imVffRecLen As Integer        'VfF record length
Dim tmSrchVffKey As INTKEY0

Dim hmSbf As Integer        'SBF file handle
Dim tmSbf As SBF            'SBF record image
Dim imSbfRecLen As Integer  'SBF record length

'spots for calendar export
Dim hmSdf As Integer        'SDF file handle
Dim tmSdf As SDF            'SDF record image
Dim imSdfRecLen As Integer  'SDF record length
Dim hmSmf As Integer

'Product
Dim hmPrf As Integer        'Prf Handle
Dim tmPrf As PRF
Dim imPrfRecLen As Integer      'Prf record length
Dim tmPrfSrchKey As LONGKEY0  'Prf key record image

Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long

Dim lmSlfSplit() As Long           '9-30-02 make 4 tables common: slsp slsp share %
Dim imSlfCode() As Integer
Dim imslfcomm() As Integer             'slsp under comm %, unused in this rept but reqd for common subroutine
Dim imslfremnant() As Integer          'slsp under remnant %, unused in this rept but reqd for common subroutine
Dim lmSlfSplitRev() As Long           '1-31-04 Rev % split for B & B Sales comm. if 0% and Slsp comm % in chf is non zero, calc
Dim lmTempGross As Long
Dim lmTempNet As Long
Dim lmTempPct As Long
Dim lmTempAcquisition As Long       '2-28-14
Dim bmStdExport As Boolean
Dim tmSdfExtSort() As SDFEXTSORT
Dim tmSdfExt() As SDFEXT
Dim smClientName As String
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcVehicle)
Dim imIncludeCodes As Integer
Dim imUseCodes() As Integer
Dim imExportOption As Integer       'lbcExport.ItemData(lbcExport.ListIndex)
Dim smExportOptionName As String    'matrix or tableau or RAB (1-29-20)
Dim smMonthBy As String
'TTP 9992 - Custom Logging and Amazon support stuff
Dim myBucket As CsiToAmazonS3.ApiCaller
Dim imVehCount As Integer
Dim lgExportCount As Long
Dim smExportFilename As String
'CEF Type
Dim hmCef As Integer        'TTP 9992 - comment for Salemsan (user) Email Addresses
Dim tmCef As CEF
Dim imCefRecLen As Integer
Dim tmCefSrchKey0 As LONGKEY0

Dim tmBillCycle As BILLCYCLERAB               '1-27-21 - If pulling RAB by Bill method, then an extra set of dates needs to be maintained

'2/25/21 Podcast Ad Server for RAB
Dim ilPcfLoop As Integer
Dim hmPcf As Integer            'cpm podcast handle
Dim tmPcf() As PCF
Dim imPcfRecLen As Integer    'PCF record length

Dim tmSBFAdjust() As ADJUSTLIST
Dim tmRvfArray() As RVF
Dim lmPacingDate As Long 'TTP 10163
Dim imPacingDay As Integer 'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range

'TTP 10205
Dim imLastAdfCode As Integer
Dim smLastADFXRefId As String
Dim lmLastADFXCRMID As Long
Dim imLastAgyCode As Integer
Dim smLastAGFXRefId As String
Dim lmLastAGFXCRMID As Long
Dim lmLastVehicleId As Integer
Dim imLastOwnerId As Integer
Dim smLastOwner As String

'            Matrix Export - Gather Projection data from contracts
'            Loop thru contracts within date last std bdcst billing
'            through the number of months requested (up to 24 months).
'
'********************************************************************************************
Function mCrMatrixProj() As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slStdStart As String            'start date to gather (std start)
    Dim slStdEnd As String              'end date to gather (end of std year)
    Dim slCalStart As String            'start date to gather (Cal start)
    Dim slCalEnd As String              'end date to gather (end of Cal year)
    Dim slTempStart As String           'Start date of period requested (minus 1 month) to handle makegoods outside end date of contrct
    Dim slCntrStatus As String          'list of contract status to gather (working, order, hold, etc)
    Dim slCntrType As String            'list of contract types to gather (Per inq, Direct Response, Remnants, etc)
    Dim ilHOState As Integer            'which type of HO cntr states to include (whether revisions should be included)
    Dim llContrCode As Long
    Dim ilCurrentRecd As Integer
    Dim ilLoop As Integer
    Dim ilClf As Integer                'loop count for lines
    Dim ilTemp As Integer
    Dim llStdStart As Long              'requested start date to gather (std serial date)
    Dim llStdEnd As Long                'requested end date to gather (std serial date)
    Dim llCalStart As Long              'requested start date to gather (Cal serial date)
    Dim llCalEnd As Long                'requested end date to gather (Cal serial date)
    Dim llDate As Long
    Dim ilCorT As Integer
    Dim ilStartCorT As Integer
    Dim ilEndCorT As Integer
    Dim slCashAgyComm As String
    Dim slPctTrade As String
    Dim slNet As String
    Dim llNet As Long
    Dim slGross As String
    Dim llGross As Long
    Dim slGrossPct As String
    Dim llAmt As Long
    Dim ilFoundMonth As Integer
    Dim ilFirstProjInxStd As Integer
    Dim ilFirstProjInxCal As Integer
    ReDim llTempGross(0 To 24) As Long  'max 24 months projection, gross $, index zero ignored
    ReDim llTempNet(0 To 24) As Long    'max 24 months projection, net $, index zero ignored
    Dim tlMatrixInfo As MATRIXINFO
    Dim ilLoopOnMonth As Integer
    ReDim tlSbf(0 To 0) As SBF
    Dim ilError As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim slDate As String
    Dim tlSBFTypes As SBFTypes
    Dim ilLoopOnSlsp As Integer
    Dim ilMnfSubCo As Integer
    Dim ilReverseSign As Integer
    Dim llSlsRevNetShare As Long
    Dim llSlsRevGrossShare As Long
    Dim blValidVehicle As Boolean
    Dim ilAcqCommPct As Integer
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim llAcqNet As Long
    Dim llAcqComm As Long
    Dim blAcqOK As Boolean
    Dim tlPriceTypes As PRICETYPES      '1-23-20
    Dim ilUseAcquisitionCost As Integer
    Dim ilAdjustDays As Integer
    Dim ilValidDays(0 To 6) As Integer
    Dim lsCntrStartDate As String
    Dim lsCntrEndDate As String
    Dim ilMatchCntr As Integer
    '2/25/21 Podcast Ad Server for RAB
    Dim tmTranTypes As TRANTYPES
    ReDim tlRvf(0 To 0) As RVF
    Dim slRvfStart As String            'earliest date to retrieve billed adserver trans
    Dim slRvfEnd As String              'latest date to retrieve billed adserver trans
    Dim tlPcfArray() As PCF
    Dim tlPcf As PCF
    Dim blFound As Boolean
    Dim ilAdjust As Integer
    Dim llAmount As Long
    Dim ilRemainMonths As Integer
    Dim llCPMStartDate As Long
    Dim llCPMEndDate As Long
    Dim ilMatrixLoop As Integer
    Dim sTmpBillCycle As String
    Dim ilVehLoop As Integer
    Dim llBilledAmt As Long 'Boostr Phase 2
    Dim llOrderedAmt As Long 'Boostr Phase 2
    Dim llRemainingAmt As Long 'Boostr Phase 2
    Dim llMonthlyAmt As Long 'Boostr Phase 2
    Dim llTempStartDate As Long 'Boostr Phase 2
    Dim llTempEndDate As Long 'Boostr Phase 2
    '-----------------------------------------------
    'defaults for retrieval of billed rvf items [if there are any adserver $ to project]
    tmTranTypes.iInv = True
    tmTranTypes.iWriteOff = False
    tmTranTypes.iPymt = False
    tmTranTypes.iCash = True
    tmTranTypes.iTrade = False
    tmTranTypes.iMerch = False
    tmTranTypes.iPromo = False
    tmTranTypes.iNTR = False
    tmTranTypes.iAirTime = True
    tmTranTypes.iAdj = True
    ilError = False             'assume everything is OK
    
    'TTP 10163
    lmPacingDate = 0
    ilAdjustDays = 30
    If (imExportOption = EXP_CUST_REV) Then
        If Trim(edcPacing.Text) <> "" Then
            lmPacingDate = gDateValue(Trim(edcPacing.Text))
            ilAdjustDays = 90
            ilFirstProjInxStd = 1
        End If
    End If
    
    '-----------------------------------------------
    'determine First Projected Index (Std)
    For ilLoop = 1 To igPeriods Step 1
        If tmBillCycle.lStdBillCycleLastBilled >= tmBillCycle.lStdBillCycleStartDates(ilLoop) And tmBillCycle.lStdBillCycleLastBilled < tmBillCycle.lStdBillCycleStartDates(ilLoop + 1) Then
            ilFirstProjInxStd = ilLoop + 1
            Exit For
        End If
    Next ilLoop
    If ilFirstProjInxStd = 0 Then
        ilFirstProjInxStd = 1                          'all projections, no actuals
    End If
    
    '-----------------------------------------------
    'determine First Projected Index (Cal)
    For ilLoop = 1 To igPeriods Step 1
        If tmBillCycle.lStdBillCycleLastBilled >= tmBillCycle.lCalBillCycleStartDates(ilLoop) And tmBillCycle.lStdBillCycleLastBilled < tmBillCycle.lCalBillCycleStartDates(ilLoop + 1) Then
            ilFirstProjInxCal = ilLoop + 1
            Exit For
        End If
    Next ilLoop
    'TTP 10734 - RAB cal contract: missing invoiced month for digital line contract that was revised after invoicing
    If ilFirstProjInxCal = 0 Or rbcMonthBy(2).Value = True Then
        ilFirstProjInxCal = 1                          'all projections, no actuals
    End If
    
    '-----------------------------------------------
    'determine what month index the actual is (versus the future dates)
    If tmBillCycle.ilBillCycle = 0 Or tmBillCycle.ilBillCycle = 2 Then                        'std or Bill Method
        'For ilLoop = 1 To igPeriods Step 1
        '    If tmBillCycle.lStdBillCycleLastBilled >= tmBillCycle.lStdBillCycleStartDates(ilLoop) And tmBillCycle.lStdBillCycleLastBilled < tmBillCycle.lStdBillCycleStartDates(ilLoop + 1) Then
        '        ilFirstProjInxStd = ilLoop + 1
        '        slStdStart = Format$(tmBillCycle.lStdBillCycleStartDates(ilFirstProjInxStd), "m/d/yy")
        '        Exit For
        '    End If
        'Next ilLoop
        'If ilFirstProjInxStd = 0 Then
        '    ilFirstProjInxStd = 1                          'all projections, no actuals
        'End If
        If tmBillCycle.lStdBillCycleLastBilled >= tmBillCycle.lStdBillCycleStartDates(igPeriods + 1) And lmPacingDate = 0 Then   'all data was in the past only, dont do contracts (Unless were pacing)
            Exit Function
        End If

        'llStdStart = tmBillCycle.lStdBillCycleStartDates(ilFirstProjInxStd)  'first date to project
        'llStdEnd = tmBillCycle.lStdBillCycleStartDates(igPeriods + 1)                'end date to project
        'slStdStart = Format$(tmBillCycle.lStdBillCycleStartDates(1), "m/d/yy")       'assume first date of proj is the quarter entered
        'slStdEnd = Format$(tmBillCycle.lStdBillCycleStartDates(igPeriods + 1), "m/d/yy") 'end date to project
    Else                                    'calendar by sch line
        ilAdjustDays = (tmBillCycle.lCalBillCycleStartDates(igPeriods + 1) - tmBillCycle.lCalBillCycleStartDates(1)) + 1
        ReDim llCalSpots(0 To ilAdjustDays) As Long        'init buckets for daily calendar values
        ReDim llCalAmt(0 To ilAdjustDays) As Long
        ReDim llCalAcqAmt(0 To ilAdjustDays) As Long
        ReDim llCalAcqNetAmt(0 To ilAdjustDays) As Long
        For ilLoop = 0 To 6                         'days of the week
            ilValidDays(ilLoop) = True              'force alldays as valid
        Next ilLoop
        'dont bother trying to calculate rates that are 0, since this report doesnt need spot counts
        tlPriceTypes.iCharge = True     'Chargeable lines
        tlPriceTypes.iZero = False      '.00 lines
        tlPriceTypes.iADU = False     'adu lines
        tlPriceTypes.iBonus = False          'bonus lines
        tlPriceTypes.iNC = False          'N/C lines
        tlPriceTypes.iRecap = False        'recapturable
        tlPriceTypes.iSpinoff = False      'spinoff
        ilUseAcquisitionCost = False
        'ilFirstProjInxCal = 1
        'tmBillCycle.iCalBillCycleLastBilledInx = 1

        'llStdStart = tmBillCycle.lCalBillCycleStartDates(ilFirstProjInxCal)  'first date to project
        'llStdEnd = tmBillCycle.lCalBillCycleStartDates(igPeriods + 1)                'end date to project
        'slStdStart = Format$(tmBillCycle.lStdBillCycleStartDates(1), "m/d/yy")       'assume first date of proj is the quarter entered
        'slStdEnd = Format$(tmBillCycle.lStdBillCycleStartDates(igPeriods + 1), "m/d/yy") 'end date to project
    End If
    
    '-----------------------------------------------
    'Get the Start and End Dates (Std)
    llStdStart = tmBillCycle.lStdBillCycleStartDates(ilFirstProjInxStd)  'first date to project
    slStdStart = Format$(llStdStart, "m/d/yy")       'assume first date of proj is the quarter entered
    llStdEnd = tmBillCycle.lStdBillCycleStartDates(igPeriods + 1)                'end date to project
    slStdEnd = Format$(llStdEnd, "m/d/yy") 'end date to project
    '-----------------------------------------------
    'Get the Start and End Dates (Cal)
    llCalStart = tmBillCycle.lCalBillCycleStartDates(ilFirstProjInxCal)  'first date to project
    slCalStart = Format$(llCalStart, "m/d/yy")       'assume first date of proj is the quarter entered
    llCalEnd = tmBillCycle.lCalBillCycleStartDates(igPeriods + 1)                'end date to project
    slCalEnd = Format$(llCalEnd, "m/d/yy") 'end date to project
    '-----------------------------------------------
    'Get the Start and End Dates (Bill Cycle)
    If tmBillCycle.ilBillCycle = 2 Then                       'Use Bill Method, expand Date range to include min/max of Start and End dates from both Std Bcast Cal + monthly Cal, we will sort the extra's out later
        'ilFirstProjInxStd already calc'd, now calc ilFirstProjInxCal
        'For ilLoop = 1 To igPeriods Step 1
        '    If tmBillCycle.lCalBillCycleLastBilled >= tmBillCycle.lCalBillCycleStartDates(ilLoop) And tmBillCycle.lCalBillCycleLastBilled < tmBillCycle.lCalBillCycleStartDates(ilLoop + 1) Then
        '        ilFirstProjInxCal = ilLoop + 1
        '        Exit For
        '    End If
        'Next ilLoop
        
        llStdStart = IIF(tmBillCycle.lCalBillCycleStartDates(ilFirstProjInxCal) <= tmBillCycle.lStdBillCycleStartDates(ilFirstProjInxStd), tmBillCycle.lCalBillCycleStartDates(ilFirstProjInxCal), tmBillCycle.lStdBillCycleStartDates(ilFirstProjInxStd))  'first date to project
        slStdStart = Format$(llStdStart, "m/d/yy")
        llStdEnd = IIF(tmBillCycle.lCalBillCycleStartDates(igPeriods + 1) >= tmBillCycle.lStdBillCycleStartDates(igPeriods + 1), tmBillCycle.lCalBillCycleStartDates(igPeriods + 1), tmBillCycle.lStdBillCycleStartDates(igPeriods + 1)) 'end date to project
        slStdEnd = Format$(llStdEnd, "m/d/yy")
    End If
   
    '-----------------------------------------------
    'setup type statement as to which type of SBF records to retrieve (only NTR)
    tlSBFTypes.iNTR = True          'include NTR billing
    tlSBFTypes.iInstallment = False      'exclude Installment billing
    tlSBFTypes.iImport = False           'exclude rep import billing

    '-----------------------------------------------
    'set contract types to retrieve
    slCntrStatus = "HOGN"                 'statuses: hold, order, unsch hold, uns order
    slCntrType = "CTRQ"         'all types: PI, DR, Remnant.  Ignore Reservations(V since they are not invoiced) and ignore PSA(p) and Promo(m)
    ilHOState = 2                       'get latest orders & revisions  (HOGN plus any revised orders WCI)

    '-----------------------------------------------
    'TTP 10163 (Re-do, to reset Index
    lmPacingDate = 0
    'ilAdjustDays = 30
    If (imExportOption = EXP_CUST_REV) Then
        If Trim(edcPacing.Text) <> "" Then
            lmPacingDate = gDateValue(Trim(edcPacing.Text))
            lmPacingDate = lmPacingDate + imPacingDay 'TTP 10596 - : Custom Revenue Export: add capability of running pacing version for a date range
            ilAdjustDays = 90
            ilFirstProjInxStd = 1
            llStdStart = tmBillCycle.lStdBillCycleStartDates(ilFirstProjInxStd)  'first date to project
            slStdStart = Format$(llStdStart, "m/d/yy")
        End If
    End If

    '------------------------------------------------------------------------------------------------
    'build table (into tmChfAdvtExt) of all contracts that fall within the dates required
    'slStdStart = Format$((gDateValue(slStdStart) - ilAdjustDays), "m/d/yy")
    If lmCntrNo > 0 Then
        ReDim tmChfAdvtExt(0 To 1) As CHFADVTEXT
        tmChfSrchKey1.lCntrNo = lmCntrNo
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, Len(tmChf), tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
           ReDim tmChfAdvtExt(0 To 0) As CHFADVTEXT
        Else
            'setup 1 entry in the active contract array for processing single contract
            tmChfAdvtExt(0).lCntrNo = tmChf.lCntrNo
            tmChfAdvtExt(0).lCode = tmChf.lCode
            tmChfAdvtExt(0).iSlfCode(0) = tmChf.iSlfCode(0)
            tmChfAdvtExt(0).iAdfCode = tmChf.iAdfCode
        End If
    Else
        'Debug.Print "obtain Contracts for: " & slStdStart & " to " & slStdEnd;
        'ilRet = gObtainCntrForDate(ExpMatrix, slStdStart, slStdEnd, slCntrStatus, slCntrType, ilHOState, tmChfAdvtExt())
        If tmBillCycle.ilBillCycle = 0 Or tmBillCycle.ilBillCycle = 2 Then                        'std or Bill Method
            slTempStart = Format$((gDateValue(slStdStart) - ilAdjustDays), "m/d/yy")
            Debug.Print "mCrMatrixProj - obtain Contracts for (std): " & slTempStart & " to " & slStdEnd
            ilRet = gObtainCntrForDate(ExpMatrix, slTempStart, slStdEnd, slCntrStatus, slCntrType, ilHOState, tmChfAdvtExt())
        Else
            slTempStart = Format$((gDateValue(slCalStart) - ilAdjustDays), "m/d/yy")
            Debug.Print "mCrMatrixProj - obtain Contracts for (cal): " & slTempStart & " to " & slCalEnd
            ilRet = gObtainCntrForDate(ExpMatrix, slTempStart, slCalEnd, slCntrStatus, slCntrType, ilHOState, tmChfAdvtExt())
        End If
    End If
    
    For ilLoop = 1 To 24            'init the projected gross & net values
        llTempNet(ilLoop) = 0
        lmProject(ilLoop) = 0
        lmAcquisition(ilLoop) = 0
        llTempGross(ilLoop) = 0
    Next ilLoop
    
    '------------------------------------------------------------------------------------------------
    'Contract Loop
    For ilCurrentRecd = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1                                          'loop while llCurrentRecd < llRecsRemaining
        'TTP 10163
        If lmPacingDate > 0 Then
            llContrCode = gPaceCntr(tmChfAdvtExt(ilCurrentRecd).lCntrNo, lmPacingDate, hmCHF, tmChf)
        Else
            llContrCode = tmChfAdvtExt(ilCurrentRecd).lCode
        End If
        
        If llContrCode = 0 Then
            'no contract found
        Else
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())
Debug.Print "mCrMatrixProj, processing Contract #" & tmChfAdvtExt(ilCurrentRecd).lCntrNo & " (" & tmChfAdvtExt(ilCurrentRecd).lCode & ")"
            sTmpBillCycle = "S"  'ilBillCycle '0=Std, 1=Cal, 2=billing cycle of contract
            If tmBillCycle.ilBillCycle = 1 Then sTmpBillCycle = "C"
            If tmBillCycle.ilBillCycle = 2 Then sTmpBillCycle = IIF(tgChfCT.sBillCycle = "C", "C", "S") 'Use Contract to determine Cal
            
            ilMatchCntr = False
            If tmBillCycle.ilBillCycle = 2 Then
                gUnpackDate tgChfCT.iStartDate(0), tgChfCT.iStartDate(1), lsCntrStartDate
                gUnpackDate tgChfCT.iEndDate(0), tgChfCT.iEndDate(1), lsCntrEndDate
                
                'Determine If This contract should be ignored; check contract within date range using the Contract's BillCycle
                ilMatchCntr = False
                If sTmpBillCycle = "C" Then
                    'Check if Contract is in Date span using Monthly Cal
                    If DateValue(lsCntrEndDate) > DateValue(Format(tmBillCycle.lCalBillCycleStartDates(1), "m/d/yy")) And DateValue(lsCntrStartDate) < DateValue(Format(tmBillCycle.lCalBillCycleStartDates(igPeriods + 1), "m/d/yy")) Then
                        ilMatchCntr = True
                    End If
                Else
                    'Check if Contract is in Date span using Std BCast Cal
                    If DateValue(lsCntrEndDate) >= DateValue(Format(tmBillCycle.lStdBillCycleStartDates(1), "m/d/yy")) And DateValue(lsCntrStartDate) < DateValue(Format(tmBillCycle.lStdBillCycleStartDates(igPeriods + 1), "m/d/yy")) Then
                        ilMatchCntr = True
                    End If
                End If
            Else
                'the date span is ok
                ilMatchCntr = True
            End If
                    
            If (lmCntrNo = 0) Or (lmCntrNo <> 0 And lmCntrNo = tgChfCT.lCntrNo) And ilMatchCntr Then    'single contract for debugging only
                 'obtain agency for commission, lines need them for net calculation, as well as for NTR
                 If tgChfCT.iAgfCode > 0 Then
                    slCashAgyComm = ".00"
                    tmAgfSrchKey.iCode = tgChfCT.iAgfCode
                    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        slCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                    End If
                 Else
                     slCashAgyComm = ".00"
                 End If              'iagfcode > 0
    
                '------------------------------------------------------------------------------------------------
                'prepare common data to get exported
                'tlMatrixInfo.iSlfCode = tgChfCT.iSlfCode(0)
                tlMatrixInfo.iAgfCode = tgChfCT.iAgfCode
                tlMatrixInfo.iAdfCode = tgChfCT.iAdfCode
                tlMatrixInfo.sProduct = tgChfCT.sProduct
                '1-22-12 obtain primary and secondary competitive codes
                tlMatrixInfo.iMnfComp1 = tgChfCT.iMnfComp(0)
                tlMatrixInfo.iMnfComp2 = tgChfCT.iMnfComp(1)
                '4-3-13 Order type:  standard, psa, promo, dr, pi, etc
                tlMatrixInfo.sOrderType = tgChfCT.sType
                tlMatrixInfo.lCntrNo = tgChfCT.lCntrNo                      '1-28-20
                tlMatrixInfo.lExtCntrNo = tgChfCT.lExtCntrNo   'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
                tlMatrixInfo.iNTRType = 0                                   '1-28-20
                gUnpackDateLong tgChfCT.iOHDDate(0), tgChfCT.iOHDDate(1), tlMatrixInfo.lOHDPacingDate           '6-17-20 added for RAB
                For ilLoop = 1 To 24                '1-22-12
                    tlMatrixInfo.lDirect(ilLoop) = 0
                Next ilLoop
                
                '------------------------------------------------------------------------------------------------
                'get NTR
                If tgChfCT.sNTRDefined = "Y" And ExpMatrix!ckcNTR.Value = vbChecked Then        'this has NTR billing
                    If tmBillCycle.ilBillCycle = 0 Or tmBillCycle.ilBillCycle = 2 Then                        'std
                        mMatrixNTR tlSBFTypes, tmBillCycle.lStdBillCycleStartDates(), tlMatrixInfo, ilFirstProjInxStd, slCashAgyComm
                    Else
                        mMatrixNTR tlSBFTypes, tmBillCycle.lCalBillCycleStartDates(), tlMatrixInfo, ilFirstProjInxCal, slCashAgyComm
                    End If
                End If
                
                '------------------------------------------------------------------------------------------------
                '2/25/21 Podcast Ad Server for RAB
                'TTP 10103: Podcast Ad Server buys: prior to invoicing, Matrix, Tableau, RAB, Cust Rev and Efficio (projections) (standard broadcast calendar) do not include ad server revenue; and should
                tlMatrixInfo.iNTRType = 0                   'JW 3-2-23 reinit ntr type for AdServer, found this issue while working on TTP 10665
                ilRet = Asc(tgSaf(0).sFeatures8)
                If ((ilRet And PODCASTCPMTAG) = PODCASTCPMTAG) And (tgChfCT.sAdServerDefined = "Y") And (imExportOption = EXP_RAB Or imExportOption = EXP_MATRIX Or imExportOption = EXP_EFFICIOPROJ Or imExportOption = EXP_TABLEAU Or imExportOption = EXP_CUST_REV) Then
                    ReDim tmSBFAdjust(0 To 0) As ADJUSTLIST  'build new for every contract
                    'Insure the common monthly buckets are initialized for the schedule lines
                    For ilLoop = 1 To 24
                        lmProject(ilLoop) = 0
                        lmAcquisition(ilLoop) = 0
                    Next ilLoop
                    
                    '-------------------------------------
                    'Get PCF (pcf_Pod_CPM_Cntr)
                    ilRet = gObtainPcf(hmPcf, tgChfCT.lCode, tmPcf())       'obtain all pcm for matching contract code
                    If ilRet Then
                        '-------------------------------------
                        'Get Receivable info
                        gUnpackDate tgChfCT.iStartDate(0), tgChfCT.iStartDate(1), slRvfStart 'Contract Start
                        If sTmpBillCycle = "C" Then 'Cal
                            slRvfEnd = Format(tmBillCycle.lCalBillCycleStartDates(ilFirstProjInxCal) - 1, "ddddd")
                        Else
                            slRvfEnd = Format(tmBillCycle.lStdBillCycleStartDates(ilFirstProjInxStd) - 1, "ddddd")
                        End If
                        '-------------------------------------
                        'gather all the receivables that exist for this contract
                        ilRet = gObtainPhfRvfbyCntr(Me, tgChfCT.lCntrNo, slRvfStart, slRvfEnd, tmTranTypes, tlRvf())
                        If ilRet Then
                            '-------------------------------------
                            'create entry into tmSBFAdjust, for all Lines ORDERED from PCF
                            For ilPcfLoop = LBound(tmPcf) To UBound(tmPcf) - 1
                                tlPcf = tmPcf(ilPcfLoop)
                                blValidVehicle = True
                                If Not gFilterLists(tlPcf.iVefCode, imIncludeCodes, imUseCodes()) Then
                                    blValidVehicle = False        'not a selected vehicle; bypass
                                End If
                                If blValidVehicle And tlPcf.sType <> "P" Then            'get only the hidden lines or standard line IDs, ignore Pkg
                                    blFound = False
                                    For ilLoop = LBound(tmSBFAdjust) To UBound(tmSBFAdjust) - 1 Step 1
                                        If tmSBFAdjust(ilLoop).lPodCode = tlPcf.lCode Then
                                            tmSBFAdjust(ilLoop).lOrderedCPMCost = tmSBFAdjust(ilLoop).lOrderedCPMCost + tlPcf.lTotalCost
                                            blFound = True
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If Not (blFound) Then
                                        ilLoop = UBound(tmSBFAdjust)
                                        tmSBFAdjust(ilLoop).iVefCode = tlPcf.iVefCode
                                        tmSBFAdjust(ilLoop).lOrderedCPMCost = tlPcf.lTotalCost
                                        tmSBFAdjust(ilLoop).lPodCode = tlPcf.lCode
                                        'JW - 5/18/23 - Fixed RAB and B&B for V81 TTP 10725 – new issue 5-17-23.zip
                                        tmSBFAdjust(ilLoop).iPodCPMID = tlPcf.iPodCPMID
                                        ReDim Preserve tmSBFAdjust(0 To UBound(tmSBFAdjust) + 1) As ADJUSTLIST
                                    End If
                                End If
                            Next ilPcfLoop
                            
                            '-------------------------------------
                            'get the adserver Received amts (RVF)
                            'TTP 10734 - RAB cal contract: missing invoiced month for digital line contract that was revised after invoicing
                            If rbcMonthBy(2).Value = True Then 'Cal Contract
                                'Dont Load RVF for Cal Contract, we will get #'s from Contract Only.
                            Else
                                For ilLoop = LBound(tlRvf) To UBound(tlRvf) - 1
                                    For ilAdjust = LBound(tmSBFAdjust) To UBound(tmSBFAdjust) - 1
                                        'TTP 10725 - Billed and Booked Cal Spots and RAB Cal Spots: not including digital line contract that should be included
                                        'If tlRvf(ilLoop).lPcfCode = tmSBFAdjust(ilAdjust).lPodCode Then
                                        'JW - 5/18/23 - Fixed RAB and B&B for V81 TTP 10725 – new issue 5-17-23.zip
                                        'If gObtaipnPcfCPMID(tlRvf(ilLoop).lPcfCode) = tlPcf.iPodCPMID Then
                                        If gObtainPcfCPMID(tlRvf(ilLoop).lPcfCode) = tmSBFAdjust(ilAdjust).iPodCPMID Then
                                            'accumulate the billing so far
                                            gPDNToLong tlRvf(ilLoop).sGross, llAmount
                                            tmSBFAdjust(ilAdjust).lBilledCPMCost = tmSBFAdjust(ilAdjust).lBilledCPMCost + llAmount
    'Debug.Print " - RVF for PodCode:" & tmSBFAdjust(ilAdjust).lPodCode & ", tmSBFAdjust(" & ilAdjust & ") amount:" & llAmount & " - on Veh:" & tmSBFAdjust(ilAdjust).iVefCode
                                            Exit For
                                        End If
                                    Next ilAdjust
                                Next ilLoop
                            End If
                            
                            '-------------------------------------
                            'get the amounts of adserver Ordered (PCF)
                            For ilPcfLoop = LBound(tmPcf) To UBound(tmPcf) - 1
                                tlPcf = tmPcf(ilPcfLoop)
                                gUnpackDateLong tlPcf.iStartDate(0), tlPcf.iStartDate(1), llCPMStartDate
                                gUnpackDateLong tlPcf.iEndDate(0), tlPcf.iEndDate(1), llCPMEndDate
'Debug.Print " mCrMatrixProj, Processing PCF(" & tlPcf.lCode & ") between " & Format(llCPMStartDate, "ddddd") & " and " & Format(llCPMEndDate, "ddddd") & " - using BillCycle:" & sTmpBillCycle
                                'TTP 10734 - RAB cal contract: missing invoiced month for digital line contract that was revised after invoicing
                                If rbcMonthBy(2).Value = True Then 'Cal Contract
                                    ilRemainMonths = gObtainMonthsOfCPMID(llCPMStartDate, llCPMEndDate, "C")
                                Else
                                    If sTmpBillCycle = "C" Then
                                        'TTP 10725 - Cal Contract when Contract's Bill Cycle ="S", use lStdBillCycleLastBilled for Last Billed Date
                                        ilRemainMonths = gObtainMonthsOfCPMID(IIF(llCPMStartDate > IIF(tgChfCT.sBillCycle = "C", tmBillCycle.lCalBillCycleLastBilled, tmBillCycle.lStdBillCycleLastBilled), llCPMStartDate, IIF(tgChfCT.sBillCycle = "C", tmBillCycle.lCalBillCycleLastBilled, tmBillCycle.lStdBillCycleLastBilled) + 1), llCPMEndDate, sTmpBillCycle)
                                    Else
                                        ilRemainMonths = gObtainMonthsOfCPMID(IIF(llCPMStartDate > tmBillCycle.lStdBillCycleLastBilled, llCPMStartDate, tmBillCycle.lStdBillCycleLastBilled + 1), llCPMEndDate, sTmpBillCycle)
                                    End If
                                End If
'Debug.Print " mCrMatrixProj, PCF Remaining Months (" & tlPcf.lCode & "):" & ilRemainMonths

                                '---------------------------------------------
                                'Boostr Phase 2 (RAB export Std broadcast/Cal Cntr : update to use new daily or monthly method depending on Site setting for future periods)
                                If tgSpfx.iLineCostType = 1 And tlPcf.sPriceType = "F" And imExportOption = EXP_RAB Then
                                    'Flat Rate Calculation using Daily Amount * Days per period
                                    llBilledAmt = 0
                                    llOrderedAmt = 0
                                    For ilAdjust = LBound(tmSBFAdjust) To UBound(tmSBFAdjust) - 1
                                        If tmSBFAdjust(ilAdjust).lPodCode = tlPcf.lCode Then
                                            llBilledAmt = tmSBFAdjust(ilAdjust).lBilledCPMCost
                                            llOrderedAmt = tmSBFAdjust(ilAdjust).lOrderedCPMCost
                                            llRemainingAmt = llOrderedAmt - llBilledAmt
                                            llMonthlyAmt = 0
                                            For ilLoop = 1 To igPeriods
                                                If sTmpBillCycle = "C" Then
                                                    llTempStartDate = tmBillCycle.lCalBillCycleStartDates(ilLoop)
                                                    llTempEndDate = tmBillCycle.lCalBillCycleStartDates(ilLoop + 1) - 1
                                                Else
                                                    llTempStartDate = tmBillCycle.lStdBillCycleStartDates(ilLoop)
                                                    llTempEndDate = tmBillCycle.lStdBillCycleStartDates(ilLoop + 1) - 1
                                                End If
                                                If rbcMonthBy(2).Value = False Then 'Not Cal Contract
                                                    If llTempStartDate > tmBillCycle.lStdBillCycleLastBilled Then
                                                        'future
                                                        llMonthlyAmt = mDeterminePeriodAmountByDaily(Format(tmBillCycle.lStdBillCycleLastBilled, "ddddd"), Format(llTempStartDate, "ddddd"), Format(llTempEndDate, "ddddd"), Format$(llCPMStartDate, "ddddd"), Format$(llCPMEndDate, "ddddd"), llBilledAmt, llOrderedAmt) * 100
                                                    End If
                                                Else
                                                    'Cal Contract (fake the last billed date to the day before Line starts)
                                                    llMonthlyAmt = mDeterminePeriodAmountByDaily(Format$(llCPMStartDate - 1, "ddddd"), Format(llTempStartDate, "ddddd"), Format(llTempEndDate, "ddddd"), Format$(llCPMStartDate, "ddddd"), Format$(llCPMEndDate, "ddddd"), 0, llOrderedAmt) * 100
                                                End If
                                                'Fix Boostr Phase 2 issues - for Joel - Issue 4 & Issue 7
                                                tlMatrixInfo.iLineNo = tmSBFAdjust(ilAdjust).iPodCPMID
                                                tmSBFAdjust(ilAdjust).lProject(ilLoop) = llMonthlyAmt
                                                lmAcquisition(ilLoop) = lmAcquisition(ilLoop) + llMonthlyAmt
                                                lmProject(ilLoop) = lmProject(ilLoop) + llMonthlyAmt
                                                If llTempStartDate > llCPMEndDate Then Exit For
                                            Next ilLoop
                                        End If
                                    Next ilAdjust
                                Else
                                    'Monthly Average
                                    For ilAdjust = LBound(tmSBFAdjust) To UBound(tmSBFAdjust) - 1
                                        If tmSBFAdjust(ilAdjust).lPodCode = tlPcf.lCode Then
                                            'TTP 10838 - Fix "Digital Line ID" column on the Standard Broadcast export - 8/19/23
                                            tmSBFAdjust(ilAdjust).iPodCPMID = tlPcf.iPodCPMID
                                            If ilRemainMonths > 0 Then
                                                llAmount = (tmSBFAdjust(ilAdjust).lOrderedCPMCost - tmSBFAdjust(ilAdjust).lBilledCPMCost) / ilRemainMonths
                                                For ilLoop = 1 To igPeriods
                                                    If sTmpBillCycle = "C" Then 'Cal
                                                        If llCPMEndDate >= tmBillCycle.lCalBillCycleStartDates(ilLoop) And llCPMStartDate < tmBillCycle.lCalBillCycleStartDates(ilLoop + 1) Then
                                                            'TTP 10734 - RAB cal contract: missing invoiced month for digital line contract that was revised after invoicing
                                                            'If tmBillCycle.lCalBillCycleStartDates(ilLoop) > tmBillCycle.lCalBillCycleLastBilled Then
                                                            If tmBillCycle.lCalBillCycleStartDates(ilLoop) > tmBillCycle.lCalBillCycleLastBilled Or rbcMonthBy(2).Value = True Then
                                                                tmSBFAdjust(ilAdjust).lProject(ilLoop) = tmSBFAdjust(ilAdjust).lProject(ilLoop) + llAmount
                                                                lmAcquisition(ilLoop) = lmAcquisition(ilLoop) + llAmount
                                                                lmProject(ilLoop) = lmProject(ilLoop) + llAmount
                                                                'TTP 10838 - Fix "Digital Line ID" column on the Standard Broadcast export - 8/19/23
                                                                tlMatrixInfo.iLineNo = tmSBFAdjust(ilAdjust).iPodCPMID
    'Debug.Print " - mCrMatrixProj, Applied Amt:" & llAmount & " to Cal Period:" & ilLoop & " (" & Format(tmBillCycle.lCalBillCycleStartDates(ilLoop), "ddddd") & ")"
                                                            End If
                                                        End If
                                                    Else 'Std
                                                        If llCPMEndDate >= tmBillCycle.lStdBillCycleStartDates(ilLoop) And llCPMStartDate < tmBillCycle.lStdBillCycleStartDates(ilLoop + 1) Then
                                                            'TTP 10734 - RAB cal contract: missing invoiced month for digital line contract that was revised after invoicing
                                                            'If tmBillCycle.lCalBillCycleStartDates(ilLoop) > tmBillCycle.lCalBillCycleLastBilled Then
                                                            If tmBillCycle.lCalBillCycleStartDates(ilLoop) > tmBillCycle.lCalBillCycleLastBilled Or rbcMonthBy(2).Value = True Then
                                                                tmSBFAdjust(ilAdjust).lProject(ilLoop) = tmSBFAdjust(ilAdjust).lProject(ilLoop) + llAmount
                                                                lmAcquisition(ilLoop) = lmAcquisition(ilLoop) + llAmount
                                                                lmProject(ilLoop) = lmProject(ilLoop) + llAmount
                                                                'TTP 10838 - Fix "Digital Line ID" column on the Standard Broadcast export - 8/19/23
                                                                tlMatrixInfo.iLineNo = tmSBFAdjust(ilAdjust).iPodCPMID
    'Debug.Print " - mCrMatrixProj, Applied Amt:" & llAmount & " to Std Period:" & ilLoop & " (" & Format(tmBillCycle.lStdBillCycleStartDates(ilLoop), "ddddd") & ")"
                                                            End If
                                                        End If
                                                    End If
    
                                                Next ilLoop
                                            End If
                                        End If
                                    Next ilAdjust
                                End If
                                '-------------------------------------
                                'Apply Amounts and Write line
                                'TTP 10665 - RAB Cal Contract: digital/ad server contract not appearing when lines are for Jan and the start month is Jan
                                'If sTmpBillCycle = "C" Then 'Cal
                                '    ilRet = mSplitAndCreate(tmBillCycle.lCalBillCycleStartDates, tlMatrixINfo, tmBillCycle.iCalBillCycleLastBilledInx + 1, slCashAgyComm, tlPcf.iVefCode)
                                'Else
                                '    ilRet = mSplitAndCreate(tmBillCycle.lStdBillCycleStartDates, tlMatrixINfo, tmBillCycle.iStdBillCycleLastBilledInx + 1, slCashAgyComm, tlPcf.iVefCode)
                                'End If
                                If sTmpBillCycle = "C" Then 'Cal
                                    ilRet = mSplitAndCreate(tmBillCycle.lCalBillCycleStartDates, tlMatrixInfo, ilFirstProjInxCal, slCashAgyComm, tlPcf.iVefCode)
                                Else
                                    ilRet = mSplitAndCreate(tmBillCycle.lStdBillCycleStartDates, tlMatrixInfo, ilFirstProjInxStd, slCashAgyComm, tlPcf.iVefCode)
                                End If
                                '-------------------------------------
                                'init the projected gross & net values
                                For ilLoop = 1 To 24
                                    llTempNet(ilLoop) = 0
                                    lmProject(ilLoop) = 0
                                    lmAcquisition(ilLoop) = 0
                                    llTempGross(ilLoop) = 0
                                Next ilLoop
                            Next ilPcfLoop
                        End If
                    End If
                End If
                
                '------------------------------------------------------------------------------------------------
                'Contract Flights
                For ilLoop = 1 To 24
                    lmProject(ilLoop) = 0
                    lmAcquisition(ilLoop) = 0
                Next ilLoop
                
                slPctTrade = gIntToStrDec(tgChfCT.iPctTrade, 0) 're-establish if went to NTR
                tlMatrixInfo.iNTRType = 0                   '1-28-20 reinit ntr type for airtime
'Debug.Print " mCrMatrixProj, Flights: " & UBound(tgClfCT)
                tlMatrixInfo.iLineNo = 0 'JW 9/25/23  - RAB export digital line ID showing for air time lines
                
                For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                    tmClf = tgClfCT(ilClf).ClfRec
                    blValidVehicle = True
                    If Not gFilterLists(tmClf.iVefCode, imIncludeCodes, imUseCodes()) Then
                        blValidVehicle = False        'not a selected vehicle; bypass
                    End If
                    
                    If (tmClf.sType <> "A" And tmClf.sType <> "O" And tmClf.sType <> "E") And (blValidVehicle) Then
'Debug.Print " - mCrMatrixProj, Processing CLF(" & tmClf.lCode & ") "
                        If tmBillCycle.ilBillCycle = 0 Or tmBillCycle.ilBillCycle = 2 Then            'std                        'calc acq net if necessary
                            mBuildExportFlights ilClf, tmBillCycle.lStdBillCycleStartDates(), ilFirstProjInxStd, igPeriods + 1
                            For ilLoopOnMonth = ilFirstProjInxStd To igPeriods
                                '7/31/15 implement acq commission  if applicable
                                If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE And ExpMatrix!rbcNetBy(1).Value Then
                                    ilAcqCommPct = 0
                                    blAcqOK = gGetAcqCommInfoByVehicle(tmClf.iVefCode, ilAcqLoInx, ilAcqHiInx)
                                    ilAcqCommPct = gGetEffectiveAcqComm(tmBillCycle.lStdBillCycleStartDates(ilLoopOnMonth), ilAcqLoInx, ilAcqHiInx)
                                    gCalcAcqComm ilAcqCommPct, lmAcquisition(ilLoopOnMonth), llAcqNet, llAcqComm
                                    lmAcquisition(ilLoopOnMonth) = llAcqNet
                                End If
                            Next ilLoopOnMonth
                            
                            ilRet = mSplitAndCreate(tmBillCycle.lStdBillCycleStartDates(), tlMatrixInfo, ilFirstProjInxStd, slCashAgyComm, tmClf.iVefCode)
                        Else 'Cal Month
                            'TTP 10665 - RAB Cal Contract: digital/ad server contract not appearing when lines are for Jan and the start month is Jan
                            'gCalendarFlightsWithNetAcq tgClfCT(ilClf), tgCffCT(), llStdStart, llStdEnd, ilValidDays(), True, llCalAmt(), llCalSpots(), llCalAcqAmt(), llCalAcqNetAmt(), ilUseAcquisitionCost, tlPriceTypes
                            gCalendarFlightsWithNetAcq tgClfCT(ilClf), tgCffCT(), llCalStart, llCalEnd, ilValidDays(), True, llCalAmt(), llCalSpots(), llCalAcqAmt(), llCalAcqNetAmt(), ilUseAcquisitionCost, tlPriceTypes
                            'gAccumCalFromDaysWithAcqNet tmBillCycle.lCalBillCycleStartDates(), llCalAmt(), llCalAcqAmt(), llCalAcqNetAmt(), False, ilUseAcquisitionCost, "G", lmProject(), lmAcquisition(), igPeriods + 1
                            gAccumCalFromDaysWithAcqNet tmBillCycle.lCalBillCycleStartDates(), llCalAmt(), llCalAcqAmt(), llCalAcqNetAmt(), False, ilUseAcquisitionCost, "G", lmProject(), lmAcquisition(), igPeriods + 1, ilFirstProjInxCal
                            ilRet = mSplitAndCreate(tmBillCycle.lCalBillCycleStartDates(), tlMatrixInfo, ilFirstProjInxCal, slCashAgyComm, tmClf.iVefCode)
                        End If
                    End If              'tmclf.stype
                    
                    For ilLoop = 1 To 24            'init the projected gross & net values
                        llTempNet(ilLoop) = 0
                        lmProject(ilLoop) = 0
                        lmAcquisition(ilLoop) = 0
                        llTempGross(ilLoop) = 0
                    Next ilLoop
                Next ilClf                          'next schedule line
            End If                                  'selective contract #
        End If 'not llContrCode = 0
    Next ilCurrentRecd
    Erase llTempGross, llTempNet
    mCrMatrixProj = ilError
End Function

'                   mBuildexportFlights - Loop through the flights of the schedule line
'                                   and build the projections dollars into lmprojmonths array
'                   <input> ilclf = sched line index into tgClfCt
'                           llStdStartDates() - up to 24 std month start dates
'                           ilFirstProjInx - index of 1st month to start projecting
'                           ilHowManyPer - # entries containing a date to test in date array
'                   <output> lmProject = array of 24 months data corresponding to
'                                           24 std start months
'                           lmAcquisition - array of 24 months acquisition costs
'       2-28-14 implement tnet / net option
Sub mBuildExportFlights(ilClf As Integer, llStdStartDates() As Long, ilFirstProjInx As Integer, ilHowManyPer As Integer)
    Dim ilCff As Integer
    Dim slStr As String
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim llDate2 As Long
    Dim llSpots As Long
    Dim ilTemp As Integer
    Dim llStdStart As Long
    Dim llStdEnd As Long
    Dim ilMonthInx As Integer
    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilHowManyPer)
    ilCff = tgClfCT(ilClf).iFirstCff
    Do While ilCff <> -1
    tmCff = tgCffCT(ilCff).CffRec

    'first decide if its Cancel Before Start
    gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
    llFltStart = gDateValue(slStr)
    gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
    llFltEnd = gDateValue(slStr)
    If llFltEnd < llFltStart Then
        Exit Sub
    End If
    'gUnpackDate tmcff.iStartDate(0), tmcff.iStartDate(1), slStr
    'llFltStart = gDateValue(slStr)
    'backup start date to Monday
    ilLoop = gWeekDayLong(llFltStart)
    Do While ilLoop <> 0
        llFltStart = llFltStart - 1
        ilLoop = gWeekDayLong(llFltStart)
    Loop
    'gUnpackDate tmcff.iEndDate(0), tmcff.iEndDate(1), slStr
    'llFltEnd = gDateValue(slStr)
    'the flight dates must be within the start and end of the projection periods,
    'not be a CAncel before start flight, and have a cost > 0
    If (llFltStart < llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart And tmCff.lActPrice > 0) Then
        'only retrieve for projections, anything in the past has already
        'been invoiced and has been retrieved from history or receive files
        'adjust the gather dates from flights: use flight start date or requested start date, whichever is later
        If llStdStart > llFltStart Then
            llFltStart = llStdStart
        End If
        'use flight end date or requsted end date, whichever is lesser
        If llStdEnd < llFltEnd Then
            llFltEnd = llStdEnd
        End If

        For llDate = llFltStart To llFltEnd Step 7
            'Loop on the number of weeks in this flight
            'calc week into of this flight to accum the spot count
            If tmCff.sDyWk = "W" Then            'weekly
                llSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
            Else                                        'daily
                If ilLoop + 6 < llFltEnd Then           'we have a whole week
                    llSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)
                Else
                    llFltEnd = llDate + 6
                    If llDate > llFltEnd Then
                        llFltEnd = llFltEnd       'this flight isn't 7 days
                    End If
                    For llDate2 = llDate To llFltEnd Step 1
                        ilTemp = gWeekDayLong(llDate2)
                        llSpots = llSpots + tmCff.iDay(ilTemp)
                    Next llDate2
                End If
            End If
            'determine month that this week belongs in, then accumulate the gross and net $
            'currently, the projections are based on STandard bdcst
            For ilMonthInx = ilFirstProjInx To ilHowManyPer - 1 Step 1       '5-22-03 (-1 to adjust to not exceed max dates) loop thru months to find the match
                If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                    If tgClfCT(ilClf).ClfRec.sType = "E" Then            'pckage pricing where spot rate equals the entire price for the week
                        lmProject(ilMonthInx) = lmProject(ilMonthInx) + tmCff.lActPrice
                        lmAcquisition(ilMonthInx) = lmAcquisition(ilMonthInx) + tmClf.lAcquisitionCost
                    Else
                        lmProject(ilMonthInx) = lmProject(ilMonthInx) + (llSpots * tmCff.lActPrice)
                        lmAcquisition(ilMonthInx) = lmAcquisition(ilMonthInx) + (llSpots * tmClf.lAcquisitionCost)
                    Exit For
                End If
                End If
            Next ilMonthInx
        Next llDate                                     'for llDate = llFltStart To llFltEnd
    End If                                          '
    ilCff = tgCffCT(ilCff).iNextCff                   'get next flight record from mem
    Loop                                            'while ilcff <> -1
End Sub

'           mCloseMatrixfiles - Close all applicable files for
'                       Matrix Export
'
Sub mCloseMatrixFiles()
    Dim ilRet As Integer
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmSbf)
    

    btrDestroy hmAgf
    btrDestroy hmCHF
    btrDestroy hmVef
    btrDestroy hmMnf
    btrDestroy hmSof
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmAgf
    btrDestroy hmSbf
    If Not bmStdExport Then         'not standard, must be calendar
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVsf)
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmVsf
    End If
    
    'TTP 9992
    If imExportOption = EXP_CUST_REV Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
    End If
End Sub

'           mOpenMatrixFiles - open files applicable to Matrix Export
'                           The Export takes all Receivables/History for up to 24 months
'                           and Contract projections and creates a text file of vehicles and
'                           their monthly gross & net $
'
'
Function mOpenMatrixFiles() As Integer
    Dim ilRet As Integer
    Dim ilTemp As Integer
    Dim ilError As Integer
    Dim slStamp As String
    Dim tlSof As SOF

    ilError = False

    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenMatrixFilesErr
    gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen AGF)", ExpMatrix
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)

    hmSbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenMatrixFilesErr
    gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen SBF)", ExpMatrix
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf)

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenMatrixFilesErr
    gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen CHF)", ExpMatrix
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenMatrixFilesErr
    gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen VEF)", ExpMatrix
    On Error GoTo 0
    imVefRecLen = Len(tmVef)

    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenMatrixFilesErr
    gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen MNF)", ExpMatrix
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)

    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenMatrixFilesErr
    gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen SOF)", ExpMatrix
    On Error GoTo 0
    imSofRecLen = Len(tlSof)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenMatrixFilesErr
    gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen CLF)", ExpMatrix
    On Error GoTo 0
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenMatrixFilesErr
    gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen CFF)", ExpMatrix
    On Error GoTo 0
    imCffRecLen = Len(tmCff)

    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenMatrixFilesErr
    gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen PRF)", ExpMatrix
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)

    'build array of sales offices
    ilTemp = 0
    ilRet = btrGetFirst(hmSof, tlSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tmSof(0 To ilTemp) As SOF
        tmSof(ilTemp) = tlSof
        ilRet = btrGetNext(hmSof, tlSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop
    ilRet = gObtainMnfForType("S", slStamp, tmMnfSS())        'build sales sources

    If Not bmStdExport Then         'not standard, its calendar exporting
        hmSdf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenMatrixFilesErr
        gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen SDF)", ExpMatrix
        On Error GoTo 0
        imSdfRecLen = Len(tmSdf)
        
        hmSmf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenMatrixFilesErr
        gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen SMF)", ExpMatrix
        On Error GoTo 0

        hmVsf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenMatrixFilesErr
        gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen VSF)", ExpMatrix
        On Error GoTo 0
    End If
    
    'TTP 9992 - Comment for Salesman (user) email Address
    If imExportOption = EXP_CUST_REV Then
        hmCef = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenMatrixFilesErr
        gBtrvErrorMsg ilRet, "gOpenMatrixFiles (btrOpen:CEF)", ExpMatrix
        On Error GoTo 0
    End If
    
    '2/25/21 Podcast Ad Server for RAB
    'TTP 10103: Close PCF for Matrix, Tableau, RAB, Cust Rev and Efficio (projections)
    ilRet = Asc(tgSaf(0).sFeatures8)
    If ((ilRet And PODCASTCPMTAG) = PODCASTCPMTAG) And (imExportOption = EXP_RAB Or imExportOption = EXP_MATRIX Or imExportOption = EXP_EFFICIOPROJ Or imExportOption = EXP_TABLEAU Or imExportOption = EXP_CUST_REV) Then      'using podcast Ad Server
        'open  podcast cpm file
        hmPcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmPcf, "", sgDBPath & "Pcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilError = ilRet
        End If
    End If
    
    imFirstTime = True              'set to create the header record in text file only once
    mOpenMatrixFiles = ilError
    Exit Function

mOpenMatrixFilesErr:
    ilError = True
    Return
End Function

'           mWriteExportRec - gather all the information for a month and write
'           a record to the export .csv file
'
'           <input> tlMatrixInfo - structure containing all the info required to write up to 24 months of data from
'                                   either the receivables or contract files (NTR included)
'           Return - true if error, otherwise false
'
'           9-23-11 Add more fields:  all 5 vehicle groups & split the slsp revenue
Private Function mWriteExportRec(tlMatrixInfo As MATRIXINFO) As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slStrCustomRev As String
    Dim ilIndex As Integer
    Dim ilOfficeInx As Integer
    Dim ilSSInx As Integer
    Dim slVehicle As String
    Dim ilVehicleID As Long
    Dim llVefExtId As Long
    Dim slSS As String
    Dim slOffice As String
    Dim ilOfficeID As Long
    Dim slSlsp As String
    Dim ilSlspID As Integer
    Dim slSlspEmail As String
    Dim ilSlspEmailID As Integer
    Dim slAdvt As String
    Dim ilAdvtID As Integer
    Dim slAdvtCreditStatus As String
    Dim slExtAdvtID As String
    Dim slAgency As String
    Dim slAgencyCreditStatus As String
    Dim ilAgencyID As Integer
    Dim slExtAgencyID As String
    Dim ilError As Integer
    Dim slStripCents As String
    Dim ilRemainder As Integer
    Dim slMarket As String
    Dim slResearch As String
    Dim slSubCompany As String
    Dim slFormat As String
    Dim slSubTotals As String
    Dim ilVGIndex As Integer
    Dim slPrimComp As String
    Dim slSecComp As String
    Dim ilRet As Integer
    Dim llTNet As Long
    Dim llAdjPromoMerch As Long
    Dim slContract As String                '1-28-20
    Dim slNTRType As String
    Dim ilOwnerID As Integer
    Dim slOwner As String
    Dim llAGFXCRMID As Long 'TTP 10599 - Fix TTP 10205 & TTP 10503 / JW - 11/28/22 - slight Performance fix with ADFX/AGFX lookups (External Agency ID, External Advertiser ID, Advertiser CRMID, and Agency CRMID)
    Dim llADFXCRMID As Long
    
    ilError = False
    If imFirstTime Then         'create the header record
        lgExportCount = 0
        If Trim$(smExportOptionName) = "RAB" Then
            'slStr = "Vehicle, Market, Research, Sub-Company, Format, SubTotals, Sales Source, Sales Office, Salesperson, Agency, Advertiser, Product, Order Type, Cash/Trade, AirTime/NTR, Primary Competitive Code, Secondary Competitive Code, Year, Month, Gross Direct, Gross Split, Net Split, NTR Type, Contract#, Pacing Date"
            'slStr = "Vehicle, Market, Research, Sub-Company, Format, SubTotals, Sales Source, Sales Office, Salesperson, Agency, Advertiser, Product, Order Type, Cash/Trade, AirTime/NTR, Primary Competitive Code, Secondary Competitive Code, Year, Month, Gross Direct, Gross Split, Net Split, NTR Type, Contract#, Pacing Date, Transaction Type"  '11-03-20 for RAB, TTP 10004
            'slStr = "VehicleID, Vehicle, OwnerID, Owner, Market, Research, Sub-Company, Format, SubTotals, Sales Source, Sales Office, Salesperson, Agency, Advertiser, Product, Order Type, Cash/Trade, AirTime/NTR, Primary Competitive Code, Secondary Competitive Code, Year, Month, Gross Direct, Gross Split, Net Split, NTR Type, Contract#, Pacing Date, Transaction Type"  'TTP 10447 - RAB Export: add VefCode, Participant Name, and participant code
            'slStr = "VehicleID, Vehicle, OwnerID, Owner, Market, Research, Sub-Company, Format, SubTotals, Sales Source, Sales Office, Salesperson, AgencyID, Agency, AdvertiserID, Advertiser, Product, Order Type, Cash/Trade, AirTime/NTR, Primary Competitive Code, Secondary Competitive Code, Year, Month, Gross Direct, Gross Split, Net Split, NTR Type, Contract#, Pacing Date, Transaction Type"  'TTP 10454 - RAB export: add advertiser ID (adfCode) and agency ID (agfCode)
            'slStr = "VehicleID, Vehicle, OwnerID, Owner, Market, Research, Sub-Company, Format, SubTotals, Sales Source, Sales Office, Salesperson, AgencyID, Agency, Agency Credit Status, AdvertiserID, Advertiser, Advertiser Credit Status, Product, Order Type, Cash/Trade, AirTime/NTR, Primary Competitive Code, Secondary Competitive Code, Year, Month, Gross Direct, Gross Split, Net Split, NTR Type, Contract#, Pacing Date, Transaction Type"  'TTP 10460 - RAB export: add Advertiser Credit Status and Agency Credit Status
            'slStr = "VehicleID, Vehicle, OwnerID, Owner, Market, Research, Sub-Company, Format, SubTotals, Sales Source, Sales Office, Salesperson, AgencyID, AgencyCRMID, Agency, Agency Credit Status, AdvertiserID, AdvertiserCRMID, Advertiser, Advertiser Credit Status, Product, Order Type, Cash/Trade, AirTime/NTR, Primary Competitive Code, Secondary Competitive Code, Year, Month, Gross Direct, Gross Split, Net Split, NTR Type, Contract#, ExtContractNo, Pacing Date, Transaction Type"  'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
            slStr = "VehicleID, ExtProductID, Vehicle, OwnerID, Owner, Market, Research, Sub-Company, Format, SubTotals, Sales Source, Sales Office, Salesperson, AgencyID, AgencyCRMID, Agency, Agency Credit Status, AdvertiserID, AdvertiserCRMID, Advertiser, Advertiser Credit Status, Product, Order Type, Cash/Trade, AirTime/NTR, Primary Competitive Code, Secondary Competitive Code, Year, Month, Gross Direct, Gross Split, Net Split, NTR Type, Contract#, ExtContractNo, Pacing Date, Transaction Type"  'TTP 10572 - RAB export: add Boostr Product ID (vefExtId) to export output
            
            'TTP 10838 - RAB export: include digital line number on automated export (Cal Spots and Broadcast)
            If rbcMonthBy(0).Value Or rbcMonthBy(1).Value Then  '0=Std Bcast cal or 1=Cal by Spots
                slStr = slStr & ",Digital Line ID"
            End If
            
            'TTP 10666 - Show Comment on Export?
            If ckcInclCmmts.Value = vbChecked Then
                'slStr = slStr & ",LineNo" 'TTP 10743 - RAB export: add line numbers
                slStr = slStr & ",Comment,CurrMonthAvg,NextMonthAvg"
            End If
            
        ElseIf Trim$(smExportOptionName) = "CustomRevenueExport" Then 'TTP 9992
            'Contract Number added 3/23/21
            slStrCustomRev = "Contract Number,Vehicle Name,Vehicle ID,Market,Research,Sub-Company,Format,SubTotals,Sales Office,Sales Office ID,Salesperson,Salesperson email,Salesperson ID,Agency name,Agency ID,External Agency ID,Advertiser name,Advertiser ID,External Advertiser ID,Product name,Cash/Trade,Air Time/NTR,Year,Month,Gross Direct,Gross Split,Net Split,Transaction Type,Receivables Date Entered"
            If edcPacing.Text <> "" Then
                'TTP 10163 - export Pacing Date when using Pacing Date
                slStrCustomRev = "Pacing Date,Contract Number,Vehicle Name,Vehicle ID,Market,Research,Sub-Company,Format,SubTotals,Sales Office,Sales Office ID,Salesperson,Salesperson email,Salesperson ID,Agency name,Agency ID,External Agency ID,Advertiser name,Advertiser ID,External Advertiser ID,Product name,Cash/Trade,Air Time/NTR,Year,Month,Gross Direct,Gross Split,Net Split,Transaction Type,Receivables Date Entered"
            End If
        Else
            slStr = "Vehicle, Market, Research, Sub-Company, Format, SubTotals, Sales Source, Sales Office, Salesperson, Agency, Advertiser, Product, Order Type, Cash/Trade, AirTime/NTR, Primary Competitive Code, Secondary Competitive Code, Year, Month, Gross Direct, Gross Split, Net Split"
            If ExpMatrix!rbcNetBy(1).Value Then             'tnet
                slStr = "Vehicle, Market, Research, Sub-Company, Format, SubTotals, Sales Source, Sales Office, Salesperson, Agency, Advertiser, Product, Order Type, Cash/Trade, AirTime/NTR, Primary Competitive Code, Secondary Competitive Code, Year, Month, Gross Direct, Gross Split, TNet Split"
            End If
        End If
        If Trim$(smExportOptionName) = "CustomRevenueExport" Then
            On Error GoTo mWriteExportRecErr
            Print #hmMatrix, slStrCustomRev     'write header description
            On Error GoTo 0
        Else
            On Error GoTo mWriteExportRecErr
            Print #hmMatrix, slStr     'write header description
            On Error GoTo 0
    
            slStr = "As of " & Format$(gNow(), "mm/dd/yy") & " "
            slStr = slStr & Format$(gNow(), "h:mm:ssAM/PM")
    
            On Error GoTo mWriteExportRecErr
            Print #hmMatrix, slStr        'write header description
            On Error GoTo 0
        End If
        imFirstTime = False         'do the heading and time stamp only once
    End If

    'format the month info for a contract/vehicle
    slVehicle = ""
    slMarket = ""
    slResearch = ""
    slSubCompany = ""
    slFormat = ""
    slSubTotals = ""
    slPrimComp = ""
    slSecComp = ""
    slNTRType = ""

    ilIndex = gBinarySearchVef(tlMatrixInfo.iVefCode)

    If ilIndex <> -1 Then
        slVehicle = Trim$(tgMVef(ilIndex).sName)
        ilVehicleID = Trim$(tgMVef(ilIndex).iCode)
        llVefExtId = Trim$(tgMVef(ilIndex).lExtId) 'TTP 10572 - RAB export: add Boostr Product ID (vefExtId) to export output
        mGetVGName tgMVef(ilIndex), tgMVef(ilIndex).iMnfVehGp3Mkt, slMarket
        mGetVGName tgMVef(ilIndex), tgMVef(ilIndex).iMnfVehGp5Rsch, slResearch
        mGetVGName tgMVef(ilIndex), tgMVef(ilIndex).iMnfVehGp6Sub, slSubCompany
        mGetVGName tgMVef(ilIndex), tgMVef(ilIndex).iMnfVehGp4Fmt, slFormat
        mGetVGName tgMVef(ilIndex), tgMVef(ilIndex).iMnfVehGp2, slSubTotals
        'TTP 10447 - RAB Export: add VefCode, Participant Name, and participant code
        ilOwnerID = mGetOwnerID(ilVehicleID)
        slOwner = mGetOwnerName(ilOwnerID)
    End If
    
    '1-22-12 obtain primary and secondary competitive codes to place into export
    tmMnfSrchKey.iCode = tlMatrixInfo.iMnfComp1
    ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        tmMnf.sName = ""
    End If
    slPrimComp = Trim$(tmMnf.sName)
    
    If tlMatrixInfo.iMnfComp2 > 0 Then
        tmMnfSrchKey.iCode = tlMatrixInfo.iMnfComp2
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmMnf.sName = ""
        End If
        slSecComp = Trim$(tmMnf.sName)
    End If
    
    If tlMatrixInfo.iNTRType > 0 Then
        tmMnfSrchKey.iCode = tlMatrixInfo.iNTRType          '1-28-20
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmMnf.sName = ""
        End If
        slNTRType = Trim$(tmMnf.sName)
    End If

    slSlsp = ""
    slOffice = ""
    slSS = ""
    ilIndex = gBinarySearchSlf(tlMatrixInfo.iSlfCode)
    If ilIndex <> -1 Then
        slSlsp = Trim$(tgMSlf(ilIndex).sFirstName) & " " & Trim$(tgMSlf(ilIndex).sLastName)
        'TTP 9992 - Salesman (user) ID, and Email
        ilSlspID = tgMSlf(ilIndex).iCode
        slSlspEmail = ""
        ilSlspEmailID = -1
        If Trim$(smExportOptionName) = "CustomRevenueExport" Then
            'Salemsan Comment ID: SalemanID -> UserID -> Email (SLF ->  URFSLFCode -> URF)
            For ilLoop = LBound(tgPopUrf) To UBound(tgPopUrf) Step 1
                If tgPopUrf(ilLoop).iSlfCode = ilSlspID Then
                    ilSlspEmailID = tgPopUrf(ilLoop).lEMailCefCode
                    Exit For
                End If
            Next ilLoop
            'Email Address:  SalemsanID -> UserID -> Email (SLF ->  URFSLFCode -> URF)
            If ilSlspEmailID <> -1 Then
                tmCefSrchKey0.lCode = ilSlspEmailID ' Look for the comment for this Saleman (user)
                imCefRecLen = Len(tmCef)
                ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slSlspEmail = gStripChr0(tmCef.sComment)
                End If
            End If
        End If
        
        For ilOfficeInx = LBound(tmSof) To UBound(tmSof)
            If tmSof(ilOfficeInx).iCode = tgMSlf(ilIndex).iSofCode Then
                slOffice = Trim$(tmSof(ilOfficeInx).sName)
                ilOfficeID = Trim$(tmSof(ilOfficeInx).iCode)
                'now detrmine sales source from office
                For ilSSInx = LBound(tmMnfSS) To UBound(tmMnfSS) - 1
                    If tmMnfSS(ilSSInx).iCode = tmSof(ilOfficeInx).iMnfSSCode Then
                        slSS = tmMnfSS(ilSSInx).sName
                        Exit For
                    End If
                Next ilSSInx
            End If
        Next ilOfficeInx
    End If

    slAgency = ""
    ilAgencyID = -1
    slExtAgencyID = ""
    If tlMatrixInfo.iAgfCode = 0 Then       'Direct
        slAgency = "Direct"
        'TTP 10456 - RAB, Matrix, Tableau Export: shows "-1" for agency value for direct advertiser, should show "direct"
        'slAgency = -1
    Else
        'do the binary search because if coming from past the agency wont be in memory
        ilIndex = gBinarySearchAgf(tlMatrixInfo.iAgfCode)
        If ilIndex <> -1 Then
            slAgency = Trim$(tgCommAgf(ilIndex).sName)
            'TTP 9992
            ilAgencyID = tgCommAgf(ilIndex).iCode
            'TTP 10460 - RAB export: add advertiser credit status and agency credit status
            slAgencyCreditStatus = ""
            Select Case Trim$(gStripChr0(tgCommAgf(ilIndex).sCrdApp))
                Case "A": slAgencyCreditStatus = "approved"
                Case "D": slAgencyCreditStatus = "denied"
                Case "R": slAgencyCreditStatus = "requires checking"
            End Select
            'slExtAgencyID = Trim$(tgCommAgf(ilIndex).sCodeStn)
            'slExtAgencyID = gStripChr0(tgCommAgf(ilIndex).sRefID)
            'TTP 10205: swap in Agency GUID field for the External Agency ID.
            mGetAgfxCodes tlMatrixInfo.iAgfCode, slExtAgencyID, llAGFXCRMID
            tlMatrixInfo.lAgfCRMId = llAGFXCRMID ' mGetAgfxCRMID(tlMatrixINfo.iAgfCode) 'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
        End If
    End If

    slAdvt = ""     'Advertiser Name
    ilAdvtID = -1     'Advertiser ID
    slExtAdvtID = ""     'External Advertiser ID - Use "station advertiser code" field
    ilIndex = gBinarySearchAdf(tlMatrixInfo.iAdfCode)
    If ilIndex <> -1 Then
        slAdvt = Trim$(tgCommAdf(ilIndex).sName)
        'TTP 9992
        ilAdvtID = tgCommAdf(ilIndex).iCode
        'TTP 10460 - RAB export: add advertiser credit status and agency credit status
        slAdvtCreditStatus = ""
        Select Case Trim$(gStripChr0(tgCommAdf(ilIndex).sCrdApp))
            Case "A": slAdvtCreditStatus = "approved"
            Case "D": slAdvtCreditStatus = "denied"
            Case "R": slAdvtCreditStatus = "requires checking"
        End Select
        'slExtAdvtID = Trim$(tgCommAdf(ilIndex).sCodeStn)
        'TTP 10205: swap in advertiser GUID field for the External Advertiser ID.
        mGetAdfxCodes tlMatrixInfo.iAdfCode, slExtAdvtID, llADFXCRMID
        'slExtAdvtID = mGetAdfxRefID(tlMatrixINfo.iAdfCode)
        'tlMatrixINfo.lAdfCRMId = mGetAdfxCRMID(tlMatrixINfo.iAdfCode) 'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
        tlMatrixInfo.lAdfCRMId = llADFXCRMID 'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
   End If
    
    'product, cash/trade, airtime/ntr are already strings
    For ilLoop = 1 To igPeriods       '24 months max
        If tlMatrixInfo.lNet(ilLoop) <> 0 Or (tlMatrixInfo.lAcquisition(ilLoop) <> 0 And ExpMatrix!rbcNetBy(1).Value) Then       'do not create $0
            slStr = ""
            'TTP 10447 - RAB Export: add VefCode, Participant Name, and participant code
            If Trim$(smExportOptionName) = "RAB" Then
                slStr = slStr & Trim$(ilVehicleID) & ","
                
                'TTP 10572 - RAB export: add Boostr Product ID (vefExtId) to export output
                If llVefExtId <> 0 Then
                    slStr = slStr & Trim$(llVefExtId) & ","
                Else
                    slStr = slStr & ","
                End If
            End If
            slStr = slStr & """" & Trim$(slVehicle) & """" & ","
            'TTP 10447 - RAB Export: add VefCode, Participant Name, and participant code
            If Trim$(smExportOptionName) = "RAB" Then
                slStr = slStr & Trim$(ilOwnerID) & ","
                slStr = slStr & """" & Trim$(slOwner) & """" & ","
            End If
            
            slStr = slStr & """" & Trim$(slMarket) & """" & ","
            slStr = slStr & """" & Trim$(slResearch) & """" & ","
            slStr = slStr & """" & Trim$(slSubCompany) & """" & ","
            slStr = slStr & """" & Trim$(slFormat) & """" & ","
            slStr = slStr & """" & Trim$(slSubTotals) & """" & ","
            slStr = slStr & """" & Trim$(slSS) & """" & ","
            slStr = slStr & """" & Trim$(slOffice) & """" & ","
            slStr = slStr & """" & Trim$(slSlsp) & """" & ","
            'TTP 10454 - RAB export: add advertiser ID (adfCode) and agency ID (agfCode)
            If Trim$(smExportOptionName) = "RAB" Then
                slStr = slStr & Trim$(tlMatrixInfo.iAgfCode) & ","
                If tlMatrixInfo.lAgfCRMId = 0 Then 'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
                    slStr = slStr & ","
                Else
                    slStr = slStr & Trim$(tlMatrixInfo.lAgfCRMId) & "," 'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
                End If
            End If
            slStr = slStr & """" & Trim$(slAgency) & """" & ","
            'TTP 10460 - RAB export: add advertiser credit status and agency credit status
            If Trim$(smExportOptionName) = "RAB" Then
                slStr = slStr & """" & slAgencyCreditStatus & """" & ","
            End If
            
            'TTP 10454 - RAB export: add advertiser ID (adfCode) and agency ID (agfCode)
            If Trim$(smExportOptionName) = "RAB" Then
                slStr = slStr & Trim$(tlMatrixInfo.iAdfCode) & ","
                If tlMatrixInfo.lAdfCRMId = 0 Then 'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
                    slStr = slStr & ","
                Else
                    slStr = slStr & Trim$(tlMatrixInfo.lAdfCRMId) & "," 'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
                End If
            End If
            slStr = slStr & """" & Trim$(slAdvt) & """" & ","
            'TTP 10460 - RAB export: add advertiser credit status and agency credit status
            If Trim$(smExportOptionName) = "RAB" Then
                slStr = slStr & """" & slAdvtCreditStatus & """" & ","
            End If
            
            slStr = slStr & """" & Trim$(tlMatrixInfo.sProduct) & """" & ","
            slStr = slStr & """" & tlMatrixInfo.sOrderType & """" & ","             '4-3-13
            slStr = slStr & """" & tlMatrixInfo.sCashTrade & """" & ","
            slStr = slStr & """" & tlMatrixInfo.sAirNTR & """" & ","
            slStr = slStr & """" & slPrimComp & """" & ","                  '1-22-12
            slStr = slStr & """" & slSecComp & """" & ","
            slStr = slStr & Trim$(str$(tlMatrixInfo.iYear(ilLoop))) & ","
            slStr = slStr & Trim$(str$(tlMatrixInfo.iMonth(ilLoop))) & ","
            
            'TTP 9992 - Custom Rev Export:
            slStrCustomRev = ""
            If edcPacing.Text <> "" Then
                lmPacingDate = gDateValue(edcPacing.Text)
                lmPacingDate = lmPacingDate + imPacingDay 'TTP 10596 - : Custom Revenue Export: add capability of running pacing version for a date range
                slStrCustomRev = slStrCustomRev & Format(lmPacingDate, "mm/dd/yy") & ","
            End If
            slStrCustomRev = slStrCustomRev & Trim$(str$(tlMatrixInfo.lCntrNo)) & "," 'Contract Number; added 3/23/21
            slStrCustomRev = slStrCustomRev & """" & Trim$(slVehicle) & """" & "," 'Vehicle Name
            slStrCustomRev = slStrCustomRev & ilVehicleID & "," 'Vehicle ID [#]
            slStrCustomRev = slStrCustomRev & """" & Trim$(slMarket) & """" & "," 'Market
            slStrCustomRev = slStrCustomRev & """" & Trim$(slResearch) & """" & "," 'Research
            slStrCustomRev = slStrCustomRev & """" & Trim$(slSubCompany) & """" & "," 'Sub-Company
            slStrCustomRev = slStrCustomRev & """" & Trim$(slFormat) & """" & "," 'Format
            slStrCustomRev = slStrCustomRev & """" & Trim$(slSubTotals) & """" & "," 'SubTotals
            slStrCustomRev = slStrCustomRev & """" & Trim$(slOffice) & """" & "," 'Sales Office
            slStrCustomRev = slStrCustomRev & Trim$(ilOfficeID) & "," 'Sales Office ID [#]
            slStrCustomRev = slStrCustomRev & """" & Trim$(slSlsp) & """" & "," 'SalesPerson
            slStrCustomRev = slStrCustomRev & """" & Trim$(slSlspEmail) & """" & "," 'SalesPerson email
            slStrCustomRev = slStrCustomRev & Trim$(ilSlspID) & "," 'SalesPerson ID [#]
            slStrCustomRev = slStrCustomRev & """" & Trim$(slAgency) & """" & "," 'Agency Name
            slStrCustomRev = slStrCustomRev & Trim$(ilAgencyID) & "," 'Agency ID [#]
            slStrCustomRev = slStrCustomRev & """" & Trim$(slExtAgencyID) & """" & "," 'External Agency ID
            slStrCustomRev = slStrCustomRev & """" & Trim$(slAdvt) & """" & "," 'Advertiser Name
            slStrCustomRev = slStrCustomRev & Trim$(ilAdvtID) & "," 'Advertiser ID [#]
            slStrCustomRev = slStrCustomRev & """" & Trim$(slExtAdvtID) & """" & ","  'External Advertiser ID
            slStrCustomRev = slStrCustomRev & """" & Trim$(tlMatrixInfo.sProduct) & """" & "," 'Product Name
            slStrCustomRev = slStrCustomRev & """" & Trim$(tlMatrixInfo.sCashTrade) & """" & "," 'Cash/Trade
            slStrCustomRev = slStrCustomRev & """" & Trim$(tlMatrixInfo.sAirNTR) & """" & "," 'Air Time / NTR
            slStrCustomRev = slStrCustomRev & Trim$(str$(tlMatrixInfo.iYear(ilLoop))) & "," 'Year
            slStrCustomRev = slStrCustomRev & Trim$(str$(tlMatrixInfo.iMonth(ilLoop))) & "," 'Month
                        
            If ExpMatrix!rbcNetBy(0).Value Then                 'use net vs tnet
                '1-22-12 the whole amt goes into the first slsp, as well as its split amt
                ilRemainder = tlMatrixInfo.lDirect(ilLoop) Mod 100
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gLongToStrDec(tlMatrixInfo.lDirect(ilLoop), 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & ","
                    slStrCustomRev = slStrCustomRev & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & "," 'TTP:9992-Gross Direct
                Else
                    slStr = slStr & Trim$(gLongToStrDec(tlMatrixInfo.lDirect(ilLoop), 2)) & ","
                    slStrCustomRev = slStrCustomRev & Trim$(gLongToStrDec(tlMatrixInfo.lDirect(ilLoop), 2)) & "," 'TTP:9992-Gross Direct
                End If
                'gross split
                ilRemainder = tlMatrixInfo.lGross(ilLoop) Mod 100
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gLongToStrDec(tlMatrixInfo.lGross(ilLoop), 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & ","
                    slStrCustomRev = slStrCustomRev & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & "," 'TTP:9992-Gross Split
                Else
                    slStr = slStr & Trim$(gLongToStrDec(tlMatrixInfo.lGross(ilLoop), 2)) & ","
                    slStrCustomRev = slStrCustomRev & Trim$(gLongToStrDec(tlMatrixInfo.lGross(ilLoop), 2)) & ","  'TTP:9992-Gross Split
                End If
                'net split
                ilRemainder = tlMatrixInfo.lNet(ilLoop) Mod 100
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gLongToStrDec(tlMatrixInfo.lNet(ilLoop), 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
                    slStrCustomRev = slStrCustomRev & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & "," 'TTP:9992-Net Split
                Else
                    slStr = slStr & Trim$(gLongToStrDec(tlMatrixInfo.lNet(ilLoop), 2))
                    slStrCustomRev = slStrCustomRev & Trim$(gLongToStrDec(tlMatrixInfo.lNet(ilLoop), 2)) & "," 'TTP:9992-Net Split
                End If
                
            Else            'tnet
                llAdjPromoMerch = 0
                If tlMatrixInfo.sCashTrade = "P" Or tlMatrixInfo.sCashTrade = "M" Then      'promo or merch, do not show under gross split column
                    llAdjPromoMerch = tlMatrixInfo.lGross(ilLoop)       'save gross split in case not a promo/merch amt, need to show it in gross split column.  Otherwise, blank the gross split column and
                                                                    'calc the Tnet by subtracting out Promo/merch amt
                    'tlMatrixINfo.lGross(ilLoop) & .lDirect should show zero for a merch or promo in the Gross split column, only need to subtract the amt for Tnet
                    tlMatrixInfo.lDirect(ilLoop) = 0
                    tlMatrixInfo.lGross(ilLoop) = 0
                    tlMatrixInfo.lNet(ilLoop) = -tlMatrixInfo.lNet(ilLoop)      'promo & merch must be subtracted
                Else
                    llAdjPromoMerch = llAdjPromoMerch
                End If
   
                '1-22-12 the whole amt goes into the first slsp, as well as its split amt
                'Gross Direct (entire amount of month for all slsp splits; could be same as net field if no splits)
                ilRemainder = tlMatrixInfo.lDirect(ilLoop) Mod 100
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gLongToStrDec(tlMatrixInfo.lDirect(ilLoop), 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & ","
                    slStrCustomRev = slStrCustomRev & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & "," 'TTP:9992-Gross Direct
                Else
                    slStr = slStr & Trim$(gLongToStrDec(tlMatrixInfo.lDirect(ilLoop), 2)) & ","
                    slStrCustomRev = slStrCustomRev & Trim$(gLongToStrDec(tlMatrixInfo.lDirect(ilLoop), 2)) & "," 'TTP:9992-Gross Direct
                End If
                
                ilRemainder = tlMatrixInfo.lGross(ilLoop) Mod 100           'gross split
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gLongToStrDec(tlMatrixInfo.lGross(ilLoop), 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & ","
                    slStrCustomRev = slStrCustomRev & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & "," 'TTP:9992-Gross Split
                Else
                    slStr = slStr & Trim$(gLongToStrDec(tlMatrixInfo.lGross(ilLoop), 2)) & ","
                    slStrCustomRev = slStrCustomRev & Trim$(gLongToStrDec(tlMatrixInfo.lGross(ilLoop), 2)) & "," 'TTP:9992-Gross Split
                End If
                
                'compute the final TNet value:  gross minus comm minus acquisition minus promo/merch
                llTNet = tlMatrixInfo.lNet(ilLoop) - tlMatrixInfo.lAcquisition(ilLoop)     'net minus acquisition costs
                'get the triple net
                ilRemainder = llTNet Mod 100
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gLongToStrDec(llTNet, 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
                    slStrCustomRev = slStrCustomRev & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & "," 'TTP:9992-Net Split
                Else
                    slStr = slStr & Trim$(gLongToStrDec(llTNet, 2))
                    slStrCustomRev = slStrCustomRev & slStr & Trim$(gLongToStrDec(llTNet, 2)) & "," 'TTP:9992-Net Split
                End If
            End If
            
            '1-28-20 add ntr type and contract #
            'slStr = slStr & ","
            If Trim$(smExportOptionName) = "RAB" Then           '1-28-20
                slStr = slStr & ","
                slStr = slStr & """" & Trim$(slNTRType) & """" & ","
                slStr = slStr & Trim$(str$(tlMatrixInfo.lCntrNo)) & ","
                            
                If tlMatrixInfo.lExtCntrNo = 0 Then 'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
                    slStr = slStr & ","
                Else
                    slStr = slStr & Trim$(str$(tlMatrixInfo.lExtCntrNo)) & ","
                End If

                slStr = slStr & """" & Trim$(Format(tlMatrixInfo.lOHDPacingDate, "ddddd")) & """" & ","          'added pacing (ohddate) for rab 6-17-20
                slStr = slStr & """" & Trim$(tlMatrixInfo.sTransactionType) & """"         '11-03-20 for RAB, TTP 10004 - added Transaction Type
                
                'TTP 10838 - RAB export: include digital line number on automated export (Cal Spots and Broadcast)
                If rbcMonthBy(0).Value Or rbcMonthBy(1).Value Then  '0=Std Bcast cal or 1=Cal by Spots
                    slStr = slStr & "," & Trim$(tlMatrixInfo.iLineNo)
                End If
                
                'TTP 10666 - Show Comment on Export?
                If ckcInclCmmts.Value = vbChecked Then
                    'TTP 10838
                    'slStr = slStr & "," 'TTP 10743 - RAB export: add line numbers
                    'slStr = slStr & Trim$(tlMatrixInfo.iLineNo)
                    slStr = slStr & ","
                    slStr = slStr & """" & Trim$(tlMatrixInfo.sComment(ilLoop)) & """"
                    tlMatrixInfo.sComment(ilLoop) = ""
                    'TTP 10742 - RAB Cal Spots manual export: when "include digital avg comments" is checked on, show current month and next month averages in separate comment columns to assist troubleshooting
                    slStr = slStr & ","
                    slStr = slStr & Trim$(tlMatrixInfo.dCurrMoAvg(ilLoop))
                    slStr = slStr & ","
                    slStr = slStr & Trim$(tlMatrixInfo.dNextMoAvg(ilLoop))
                End If
            End If
            
            'TTP 9992 - Custom Rev Export
            slStrCustomRev = slStrCustomRev & """" & Trim$(tlMatrixInfo.sTransactionType) & """" & ","   'Transaction Type
            If tlMatrixInfo.lReceivablesDateEntered <> 0 Then
                slStrCustomRev = slStrCustomRev & """" & Trim$(Format(tlMatrixInfo.lReceivablesDateEntered, "ddddd")) & """"  'Receivables Date Entered
            Else
                slStrCustomRev = slStrCustomRev & """" & """"  'Receivables Date Entered
            End If
            
            On Error GoTo mWriteExportRecErr
            
            If Trim$(smExportOptionName) = "CustomRevenueExport" Then
                'TTP 9992 - Custom Rev Export
                Print #hmMatrix, slStrCustomRev
            Else
                Print #hmMatrix, slStr
            End If
            
            'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
            '02/27/2023 - fix counter so that it shows even number 100's progress (instead of 99)
            lgExportCount = lgExportCount + 1
            igDOE = igDOE + 1
            If igDOE >= 100 Then
                If ckcPacingRange.Value = vbChecked Then
                    lacInfo(0).Caption = "Exporting File " & imPacingDay + 1 & ": " & lgExportCount & " records..."
                Else
                    lacInfo(0).Caption = "Exporting " & lgExportCount & " records..."
                End If
                igDOE = 0
                lacInfo(0).Refresh
                DoEvents
            End If
            On Error GoTo 0
        End If
    Next ilLoop

    For ilLoop = 1 To 24            'init the monthly info for next one
        tlMatrixInfo.iYear(ilLoop) = 0
        tlMatrixInfo.iMonth(ilLoop) = 0
        tlMatrixInfo.lGross(ilLoop) = 0
        tlMatrixInfo.lNet(ilLoop) = 0
        tlMatrixInfo.lDirect(ilLoop) = 0            '1-22-12
        tlMatrixInfo.lAcquisition(ilLoop) = 0
        tlMatrixInfo.sTransactionType = ""          '11-03-20 for RAB, TTP 10004
    Next ilLoop

    mWriteExportRec = ilError
    Exit Function

mWriteExportRecErr:
    ilError = True
    Resume Next
End Function

'*****************************************************************************************
'
'                   gCrMatrixPast - Matrix export of monthly revenue from past& future for up to 24 months
'                            1st part builds actual data from PHF & RVF.
'                            Gathers all "I" and "A" transactions and places
'                            the $ in the appropriate standard month. (2nd part is
'                            building projected data from contracts -see mCrMatrixProj)
'                   <Input>  llStdStartDates - array of up to 25 start dates, denoting
'                                              start date of each period to gather
'                            llLastBilled - Date of last invoice period
'                            ilLastbilledInx - Index into llStdStartDates of period last
'                                           invoiced
'                   <Return> 0 = OK, <> 0= error
'
'Function mCrMatrixPast(llStdStartDates() As Long, llLastBilled As Long, ilLastBilledInx As Integer) As Integer
Function mCrMatrixPast() As Integer
    Dim llCurrentRecd As Long
    Dim ilRet As Integer
    Dim slCode As String
    Dim slStr As String
    Dim llNet As Long
    Dim llGross As Long
    Dim llAcquisition As Long
    Dim llDate As Long
    Dim tlTranType As TRANTYPES
    ReDim tlRvf(0 To 0) As RVF
    Dim ilLoopRecv As Integer                   '9-16-01
    Dim tlMatrixInfo As MATRIXINFO
    Dim ilError As Integer
    Dim ilValidTran As Integer
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim ilLoopOnSlsp As Integer
    Dim llTempGross As Long
    Dim llTempNet As Long
    Dim llTempAcquisition As Long
    Dim llSplitGrossAmt As Long
    Dim llSplitNetAmt As Long
    Dim llSplitAcquisitionAmt As Long
    Dim llTempPct As Long
    Dim slGrossAmount As String
    Dim slNetAmount As String
    Dim slAcqAmount As String
    Dim slSharePct As String
    Dim ilReverseSign As Integer
    Dim slSplitGross As String
    Dim slSplitNet As String
    Dim ilUseSlsComm As Integer
    Dim ilMnfSubCo As Integer
    Dim ilLoop As Integer
    Dim ilSaveLastBilledInx As Integer
    Dim blTestOnlyPromoMerch As Boolean
    Dim blValidVehicle As Boolean
    Dim ilAcqCommPct As Integer
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim llAcqNet As Long
    Dim llAcqComm As Long
    Dim blAcqOK As Boolean
    Dim ilLastBilledInx As Integer
    Dim llDateAdjust As Long 'Allows adj's to be collected for the month that was Invoiced when the end of Month is > than the billing date
    Dim llTransEnteredDate As Long 'Pacing
    
    ilError = False
    ilUseSlsComm = False                'used for subroutine parameter
    tlTranType.iAdj = True              'look only for adjustments in the History & Rec files
    If ExpMatrix!ckcInclAdj.Value = vbUnchecked Then   '2-27-27 include adjust,ments?
        tlTranType.iAdj = False
    End If
    tlTranType.iInv = True
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = True
    tlTranType.iMerch = False               'always exclude Merchandise & promotions
    tlTranType.iPromo = False
    If ExpMatrix!ckcNTR.Value = vbChecked Then                       'include NTR?
        tlTranType.iNTR = True
    Else
        tlTranType.iNTR = False
    End If
    
    If ExpMatrix!rbcNetBy(1).Value Then         'tNet?  if so, need to obtain the promo & merchandising transactions all the way to the end of requested period
        tlTranType.iMerch = True
        tlTranType.iPromo = True
        ilSaveLastBilledInx = tmBillCycle.lCalBillCycleLastBilled
        tmBillCycle.lCalBillCycleLastBilled = igPeriods             'tnet needs to get future Merch & Promo transactions from receivables
        
        If ExpMatrix!rbcMonthBy(1).Value = True Then        'calendar option, ignore NTR since they are handled in the spot processing;
                                                            'get only the promo/merch for entire period requested
            tlTranType.iNTR = False
            If ExpMatrix!ckcInclAdj.Value = vbUnchecked Then            '2-27-17 if excldung adjustments, get nothing
                tlTranType.iCash = False
                tlTranType.iTrade = False
            End If
        End If
    End If
    
    'Determine LastBilled Index
    If tmBillCycle.ilBillCycle = 0 Then 'Std
        ilLastBilledInx = tmBillCycle.iStdBillCycleLastBilledInx
    ElseIf tmBillCycle.ilBillCycle = 1 Then 'Monthly Cal
        ilLastBilledInx = tmBillCycle.iCalBillCycleLastBilledInx
    ElseIf tmBillCycle.ilBillCycle = 2 Then 'BillClycle
        ilLastBilledInx = IIF(tmBillCycle.iCalBillCycleLastBilledInx > tmBillCycle.iStdBillCycleLastBilledInx, tmBillCycle.iCalBillCycleLastBilledInx, tmBillCycle.iStdBillCycleLastBilledInx)
    End If
    'TTP 10870 - RAB Cal Spot: bypassing adjustments when "include adjustment" checkbox is checked on and the period the export is run for doesn't include any unbilled months
    If ilLastBilledInx = 0 Then ilLastBilledInx = 1
    '4/14/21 - 'Allows adj's to be collected for the month that was Invoiced if invoiced < end of month (To determine which invoice adjustments will be included on the calendar versions, when invoice adjustments are set to be included, it gets the end date of the calendar month, and if the end date of the calendar month is less than or equal to the last invoice date, then adjustments made in that month or prior to that month will be included, otherwise they will be part of a future month.  so, get these adjustments)
    llDateAdjust = 0
    If tmBillCycle.ilBillCycle = 1 Then   'Monthly Cal
        slStr = gObtainEndCal(Format(tmBillCycle.lStdBillCycleLastBilled, "ddddd"))
        llDateAdjust = gDateValue(slStr) - tmBillCycle.lStdBillCycleLastBilled
    End If

    'TTP 10163 - using Pacing Date?
    lmPacingDate = 0
    If edcPacing.Text <> "" Then 'pacing date
        lmPacingDate = gDateValue(edcPacing.Text)
        lmPacingDate = lmPacingDate + imPacingDay 'TTP 10596 - : Custom Revenue Export: add capability of running pacing version for a date range
        If lmPacingDate > tmBillCycle.lStdBillCycleLastBilled Then         '10-19-15 use last billed date or pacing date entered, whichever is earlier for AN transactions
            lmPacingDate = tmBillCycle.lStdBillCycleLastBilled
        End If
        tlTranType.iInv = False
        tlTranType.iTrade = False
    End If
    
    For ilLoopRecv = 1 To ilLastBilledInx       'loop on # months to process for phf & rvf by contract # & tran date
        If tmBillCycle.ilBillCycle = 0 Then     'Std Cal
            If (tmBillCycle.lStdBillCycleStartDates(ilLoopRecv + 1) - 1) > tmBillCycle.lStdBillCycleLastBilled + llDateAdjust And lmPacingDate = 0 Then   '4/14/21 - 'Allows adj's to be collected for the month that was Invoiced if invoiced < end of month, TTP 10163
                tlTranType.iAdj = False         'ignore AN for tran dates in the future
                tlTranType.iNTR = False
                tlTranType.iHardCost = False
            End If
            slStr = Format$(tmBillCycle.lStdBillCycleStartDates(ilLoopRecv), "m/d/yy")
            slCode = Format$(tmBillCycle.lStdBillCycleStartDates(ilLoopRecv + 1) - 1, "m/d/yy")
            
            'TTP 10163 - never get ANs past the user effective pacing date
            If lmPacingDate > 0 Then
                If lmPacingDate < tmBillCycle.lStdBillCycleStartDates(ilLoopRecv) Then       'pacing date is prior to the months start date, done with the AN pass
                    Exit For
                ElseIf lmPacingDate >= tmBillCycle.lStdBillCycleStartDates(ilLoopRecv) And lmPacingDate < tmBillCycle.lStdBillCycleStartDates(ilLoopRecv + 1) Then
                    'pacing date falls within this month, make sure the end date isnt beyond the users entered pacing date
                    slCode = Format$(lmPacingDate, "m/d/yy")
                End If
            End If
            
        ElseIf tmBillCycle.ilBillCycle = 1 Then 'Monthly Cal
            If (tmBillCycle.lCalBillCycleStartDates(ilLoopRecv + 1) - 1) > tmBillCycle.lStdBillCycleLastBilled + llDateAdjust Then '4/2/21 Changed to Std from Cal - Per DH; when running the export by calendar, the last date invoiced must be the Last Std Bdcst Date Invoiced,  not Last Cal Date invoiced, as we do not bill by the calendar and that date will probably always be incorrect.  '4/14/21 - 'Allows adj's to be collected for the month that was Invoiced if invoiced < end of month
                tlTranType.iAdj = False         'ignore AN for tran dates in the future
                tlTranType.iNTR = False
                tlTranType.iHardCost = False
            End If
            slStr = Format$(tmBillCycle.lCalBillCycleStartDates(ilLoopRecv), "m/d/yy") 'start date
            slCode = Format$(tmBillCycle.lCalBillCycleStartDates(ilLoopRecv + 1) - 1, "m/d/yy") 'end date
            
        ElseIf tmBillCycle.ilBillCycle = 2 Then 'Use Bill method, expand Date range to include min/max of Start and End dates from both Std Bcast Cal + monthly Cal, we will sort the extra's out later
            If (tmBillCycle.lStdBillCycleStartDates(ilLoopRecv + 1) - 1) > tmBillCycle.lStdBillCycleLastBilled + llDateAdjust Then '4/14/21 - 'Allows adj's to be collected for the month that was Invoiced if invoiced < end of month
                tlTranType.iAdj = False         'ignore AN for tran dates in the future
                tlTranType.iNTR = False
                tlTranType.iHardCost = False
            End If
            slStr = IIF(tmBillCycle.lCalBillCycleStartDates(ilLoopRecv) < tmBillCycle.lStdBillCycleStartDates(ilLoopRecv), Format$(tmBillCycle.lCalBillCycleStartDates(ilLoopRecv), "m/d/yy"), Format$(tmBillCycle.lStdBillCycleStartDates(ilLoopRecv), "m/d/yy"))
            slCode = IIF(tmBillCycle.lCalBillCycleStartDates(ilLoopRecv + 1) > tmBillCycle.lStdBillCycleStartDates(ilLoopRecv + 1), Format$(tmBillCycle.lCalBillCycleStartDates(ilLoopRecv + 1) - 1, "m/d/yy"), Format$(tmBillCycle.lStdBillCycleStartDates(ilLoopRecv + 1) - 1, "m/d/yy"))
        End If
       
        ilRet = gObtainPhfRvfbyCntr(ExpMatrix, lmCntrNo, slStr, slCode, tlTranType, tlRvf())
        If ilRet = 0 Then
            'Print #hmMsg, "** Error in reading History or Receivables- export aborted"
            gAutomationAlertAndLogHandler "** Error in reading History or Receivables- export aborted"
            ilError = True
            mCrMatrixPast = ilError
            Exit Function
        End If

        For llCurrentRecd = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
            tmRvf = tlRvf(llCurrentRecd)
            '1/27/21 ---------------------
            'get contract from history or rec file
            tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd

            '9-19-06 when there is no sch lines and need to process merchandising for t-net, the contract may not be scheduled (schstatus = "N");  need to process those contracts
            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M" And tmChf.sSchStatus <> "N")
                 ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            'if contract headr not found, setup fake header
            gFakeChf tmRvf, tmChf
            '----------------------

            ilValidTran = False
            'test for inclusion/exclusion of ntrs & air time transactions
            If (tlTranType.iNTR = True And tmRvf.iMnfItem > 0) Or (tmRvf.iMnfItem = 0) Then
                ilValidTran = True
            End If
            
            blValidVehicle = True
            If Not gFilterLists(tmRvf.iAirVefCode, imIncludeCodes, imUseCodes()) Then
                blValidVehicle = False
                ilValidTran = False           'not a selected vehicle; bypass
            End If
            
            '1/27/21 ---------------------------------------
            If tgSpf.sSEnterAgeDate = "E" Then      'use entered date or ageing date
                gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr
            Else
                slCode = Trim$(str$(tmRvf.iAgePeriod) & "/15/" & Trim$(str$(tmRvf.iAgingYear)))
                slStr = gObtainEndStd(slCode)
            End If
            
            llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
            If tmBillCycle.ilBillCycle = 2 Then
                If tmChf.sBillCycle = "C" Then 'this contract have been billed by calendar
                    If llDate < tmBillCycle.lCalBillCycleStartDates(ilLoopRecv) Or llDate >= tmBillCycle.lCalBillCycleStartDates(ilLoopRecv + 1) Then       'within month of cal  dates
                        ilValidTran = False
                    End If
                Else             'standard bill contract
                    'the date cannot be beyond the end of the std date for the month processing
                    If llDate < tmBillCycle.lStdBillCycleStartDates(ilLoopRecv) Or llDate >= tmBillCycle.lStdBillCycleStartDates(ilLoopRecv + 1) Then        'out month of std dates
                        ilValidTran = False
                    End If
                End If
            End If
            
            If ilValidTran Then
                If tmRvf.sCashTrade <> "P" And tmRvf.sCashTrade <> "M" Then
                    'cash/trade IN or AN, ignore in future for standard, ok to include for calendar
                    If ExpMatrix!rbcMonthBy(0).Value = True Or ExpMatrix!rbcMonthBy(3).Value = True Then               'Std or Billing Method
                        'IN/AN for airtime or NTR cant be in the future
                        If ExpMatrix!rbcMonthBy(3).Value = True Then
                            If tmChf.sBillCycle = "C" Then
                                If (tmBillCycle.lCalBillCycleStartDates(ilLoopRecv + 1) - 1) > tmBillCycle.lCalBillCycleLastBilled Then
                                    ilValidTran = False
                                End If
                            Else
                                If (tmBillCycle.lStdBillCycleStartDates(ilLoopRecv + 1) - 1) > tmBillCycle.lStdBillCycleLastBilled Then
                                    ilValidTran = False
                                End If
                            End If
                        Else
                            If (tmBillCycle.lStdBillCycleStartDates(ilLoopRecv + 1) - 1) > tmBillCycle.lStdBillCycleLastBilled Then
                                ilValidTran = False
                            End If
                        End If
                    Else                '2-27-17 calendar, only AN allowed if requested
                        If (tmRvf.sTranType = "AN") And (tlTranType.iAdj) Then
                            'ok, valid tran
                            ilValidTran = ilValidTran
                        Else
                            ilValidTran = False
                        End If
                    End If
                End If
            End If
            
            gPDNToLong tmRvf.sNet, llNet
            gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
            slCode = Format$(llDate, "m/d/yy")
                       
            'TTP 9992 - use RVF/PHF iDateEntrd for "Receivables Date Entered" column
            gUnpackDateLong tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), tlMatrixInfo.lReceivablesDateEntered
                       
            If bmStdExport Then
                slCode = gObtainEndStd(slCode)          '11-5-13 cannot assume the month is the proper month that should be in the exported.
                                                    'ie NTR may be billed at the end of a cal month, but within the start of the std bdcst
            Else
                slCode = gObtainEndCal(slCode)
            End If
            'Setup month and year to store in export
            gObtainYearMonthDayStr slCode, True, slYear, slMonth, slDay
            
            '5/11/21 fix Pacing Date per Jason teams [5/11/21 1:25 PM]
            If lmPacingDate > 0 Then
                gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slStr
                llTransEnteredDate = gDateValue(slStr)      'date entered of Merch/promo must be equal/prior to pacing date
                If llTransEnteredDate > lmPacingDate Then
                    ilValidTran = False
                End If
            End If
    
            'PHFRVF routine has filtered only "I" & "HI" and "AN", along with the trans dates
            'see if selective contract for debugging
            'ignore Installment types of "I", which is billing, not revenue
            '11-17-17 allow $0 net but tnet with acquisition to process
'            If (llNet <> 0 And ilValidTran) And ((lmCntrNo = 0) Or (lmCntrNo <> 0 And lmCntrNo = tmRvf.lCntrNo)) And (tmRvf.sType <> "I") Then               'dont write out zero records
            If ((llNet <> 0 And ilValidTran) Or (ExpMatrix!rbcNetBy(1).Value = True And tmRvf.lAcquisitionCost <> 0)) And ((lmCntrNo = 0) Or (lmCntrNo <> 0 And lmCntrNo = tmRvf.lCntrNo)) And (tmRvf.sType <> "I") Then               'dont write out zero records
                'get contract from history or rec file if different than previous read
                'If tmRvf.lCntrNo <> tmChf.lCntrNo Then
                
'load chf is now done above, chfBillCycle is as part of valid checks
'                    tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
'                    tmChfSrchKey1.iCntRevNo = 32000
'                    tmChfSrchKey1.iPropVer = 32000
'                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
'                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
'                         ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                    Loop
'
'                    gFakeChf tmRvf, tmChf
                    ReDim lmSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
                    ReDim imSlfCode(0 To 9) As Integer             '4-20-00
                    ReDim imslfcomm(0 To 9) As Integer             'slsp under comm %
                    ReDim imslfremnant(0 To 9) As Integer          'slsp under remnant %
                    ReDim lmSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)

                    ilMnfSubCo = gGetSubCmpy(tmChf, imSlfCode(), lmSlfSplit(), tmRvf.iAirVefCode, ilUseSlsComm, lmSlfSplitRev())

                    'slsp, agency & advt & vehicles are in memory
                'End If

                If ExpMatrix!rbcNetBy(0).Value Then             'net (vs tnet)
                    gPDNToLong tmRvf.sGross, llGross
                    gPDNToLong tmRvf.sNet, llNet
                Else        'tnet

                    gPDNToLong tmRvf.sGross, llGross
                    gPDNToLong tmRvf.sNet, llNet
                    llAcquisition = tmRvf.lAcquisitionCost
'                    If tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P" Then
'                        llGross = -llGross
'                        llNet = -llNet
'                    End If
                    '7/31/15 implement acq commission  if applicable
                    If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then
                        ilAcqCommPct = 0
                        blAcqOK = gGetAcqCommInfoByVehicle(tmRvf.iBillVefCode, ilAcqLoInx, ilAcqHiInx)
                        ilAcqCommPct = gGetEffectiveAcqComm(llDate, ilAcqLoInx, ilAcqHiInx)
                        gCalcAcqComm ilAcqCommPct, llAcquisition, llAcqNet, llAcqComm
                        llAcquisition = llAcqNet
                    End If
                End If

                tlMatrixInfo.sAirNTR = "A"          'assume Air time
                'if NTR, get that commission instead
                If tmRvf.iMnfItem > 0 Then          'this indicates NTR
                    tlMatrixInfo.sAirNTR = "N"
                    'dont need any info from the NTR record at this time
                    'retrieve the associated NTR record from SBF
                    'tmSbfSrchKey.lCode = tmRvf.lSbfCode  '12-16-02
                    'ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    Print #hmMsg, "** Error in NTR transaction, NTR code " & Str$(tmRvf.lSbfCode) & " is invalid"
                    '    ilError = True
                    '    mCrMatrixPast = ilError
                    '    Exit Function
                    'End If
                End If                                                'associated sales source
                
                '1-22-12 obtain primary and secondary competitive codes
                tlMatrixInfo.iMnfComp1 = tmChf.iMnfComp(0)
                tlMatrixInfo.iMnfComp2 = tmChf.iMnfComp(1)
                '4-3-13 Order type :   Standard, PI, DR, Reservation, PSA, Promo, etc
                tlMatrixInfo.sOrderType = tmChf.sType
                tlMatrixInfo.lCntrNo = tmChf.lCntrNo
                tlMatrixInfo.lExtCntrNo = tmChf.lExtCntrNo   'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
                tlMatrixInfo.iNTRType = tmRvf.iMnfItem                       '1-28-20, RAB export doesnt use receivables/history
                tlMatrixInfo.sTransactionType = tmRvf.sTranType              '11-03-20 for RAB, TTP 10004
                gUnpackDateLong tmChf.iOHDDate(0), tmChf.iOHDDate(1), tlMatrixInfo.lOHDPacingDate           '6-17-20 added for RAB
                'TTP 10838 - Fix "Digital Line ID" column on the Standard Broadcast export - 8/19/23
                tlMatrixInfo.iLineNo = 0
                If tmRvf.lPcfCode > 0 And tmRvf.sTranType <> "AN" Then '9/22/23 JW: Per Jason, exclude AN's (WWO doesnt need Line# on ANs)
                    tlMatrixInfo.iLineNo = gObtainPcfCPMID(tmRvf.lPcfCode)
                End If
                lmTempGross = llGross
                lmTempNet = llNet
                lmTempAcquisition = llAcquisition
                lmTempPct = 0
                'determine amt of revenue sharing; could exceed 100%
                For ilLoopOnSlsp = 0 To 9
                    lmTempPct = lmTempPct + lmSlfSplit(ilLoopOnSlsp)
                Next ilLoopOnSlsp
                ilReverseSign = False
                If llNet < 0 Then
                    ilReverseSign = True            'always work with positive #s
                    lmTempGross = -lmTempGross
                    lmTempNet = -lmTempNet
                    lmTempAcquisition = -lmTempAcquisition
                    '3-18-14 invoice that was UNDO didnt show correctly
                    llGross = -llGross
                    llNet = -llNet
                    llAcquisition = -llAcquisition
                End If
    
                For ilLoopOnSlsp = 0 To 9
                    If lmSlfSplit(ilLoopOnSlsp) > 0 Then
                        tlMatrixInfo.iVefCode = tmRvf.iAirVefCode
                        'tlMatrixInfo.iSlfCode = tmRvf.iSlfCode
                        tlMatrixInfo.iSlfCode = imSlfCode(ilLoopOnSlsp)
                        tlMatrixInfo.iAgfCode = tmRvf.iAgfCode
                        tlMatrixInfo.iAdfCode = tmRvf.iAdfCode
                        tlMatrixInfo.sCashTrade = tmRvf.sCashTrade
                        If ilLoopOnSlsp = 0 Then            '1-22-12 1st slsp gets total gross amt as well as split in its record
                            If ilReverseSign Then
                                tlMatrixInfo.lDirect(ilLoopRecv) = -llGross     'working in all position #s, need to negate it if it was negative trans amt
                            Else
                                tlMatrixInfo.lDirect(ilLoopRecv) = llGross
                            End If
                        End If

                        'tlMatrixInfo.lGross(1) = llGross
                        'tlMatrixInfo.lNet(1) = llNet
                        
                        mObtainSlsRevenueShare llGross, llNet, llAcquisition, ilLoopOnSlsp, tlMatrixInfo, ilLoopRecv, ilReverseSign

                        tmPrfSrchKey.lCode = tmRvf.lPrfCode     'Product
                        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            tmPrf.sName = ""
                        End If
                        tlMatrixInfo.sProduct = tmPrf.sName
                        If Trim$(tlMatrixInfo.sProduct) = "" Then
                            tlMatrixInfo.sProduct = Trim$(tmChf.sProduct)
                        End If
                        tlMatrixInfo.iYear(ilLoopRecv) = Val(slYear)
                        tlMatrixInfo.iMonth(ilLoopRecv) = Val(slMonth)
                        ilRet = mWriteExportRec(tlMatrixInfo)
                        If ilRet <> 0 Then   'error
                            'Print #hmMsg, "Error writing export record for contract # " & str$(tmRvf.lCntrNo) & " from Receivables/History"
                            gAutomationAlertAndLogHandler "Error writing export record for contract # " & str$(tmRvf.lCntrNo) & " from Receivables/History"
                            ilError = True
                            mCrMatrixPast = ilError
                            Exit Function
                        End If
                    End If
                    
                Next ilLoopOnSlsp       'for illooponslsp = 0 to 9
            End If
        Next llCurrentRecd
    Next ilLoopRecv

    If ExpMatrix!rbcNetBy(1).Value Then         'tnet
        'ilLastBilledInx = ilSaveLastBilledInx
        tmBillCycle.lCalBillCycleLastBilled = ilSaveLastBilledInx
    End If
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
Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slNTR  As String
    Dim slCntr As String
    Dim slMissed As String
    Dim slMonthType As String
    Dim slAdj As String

    ilRet = 0
    'On Error GoTo mOpenMsgFileErr:
    slToFile = sgDBPath & "\Messages\" & "Exp" & Trim$(smExportOptionName) & ".Txt"
    sgMessageFile = slToFile
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    If ckcNTR.Value = vbChecked Then
        slNTR = "Include NTR"
    Else
       slNTR = "Exclude NTR"
    End If
    
    '6-18-20 show missed response for all options
    If ckcInclMissed.Value = vbChecked Then
        slMissed = "Include Missed"
    Else
       slMissed = "Exclude Missed"
    End If
   
    If edcContract.Text = "" Then
        slCntr = "All contracts"
    Else
        slCntr = "Cntr # " & edcContract.Text
    End If
    '6-19-20
    If rbcMonthBy(0).Value = True Then          'standard
        slMonthType = "Std Bdcst"
    ElseIf rbcMonthBy(1).Value = True Then        'calendar by spots
        slMonthType = "Calendar by Spots"
    Else
        slMonthType = "Calendar by Contract"
    End If
    
    If ckcInclAdj.Value = vbChecked Then          'incl adjustments
        slAdj = "Incl Adj"
    Else
        slAdj = "Excl Adj"
    End If
    
    gAutomationAlertAndLogHandler "** Export " & Trim$(smExportOptionName) & " **"
    If edcPacing.Text <> "" Then
        'TTP 10163 - add pacing date option to Custom Revenue Export"
        lmPacingDate = gDateValue(edcPacing.Text)
        lmPacingDate = lmPacingDate + imPacingDay 'TTP 10596 - : Custom Revenue Export: add capability of running pacing version for a date range
    End If

    mOpenMsgFile = True
    Exit Function
End Function

Private Sub ckcAll_Click()
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    Dim ilIndex As Integer
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicle.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
End Sub

Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcAmazon_Click()
    If ckcAmazon.Value = vbChecked Then
        frcAmazon.Left = 120
        frcAmazon.Top = 2880
        frcAmazon.Visible = True
    Else
        frcAmazon.Visible = False
    End If
End Sub

'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
Private Sub ckcPacingRange_Click()
    If ckcPacingRange.Value = vbChecked Then
        lacPacingRange.Visible = True
        edcPacingEnd.Visible = True
    Else
        lacPacingRange.Visible = False
        edcPacingEnd.Visible = False
        edcPacingEnd.Text = ""
    End If
    tmcClick_Timer
End Sub

Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub

Private Sub cmcExport_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDateTime As String
    Dim slMonthHdr As String * 36
    Dim ilSaveMonth As Integer
    Dim ilYear As Integer
    Dim slStart As String
    Dim slTimeStamp As String
    Dim ilHowManyDefined As Integer
    Dim ilHowMany As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim slFNMonth As String
    tmBillCycle.ilBillCycle = 0         '1-27-21 assume pulling the RAB by std cal
    lgExportCount = 0
    igDOE = 0
    
    lacInfo(0).Visible = True
    lacInfo(1).Visible = False
    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    'TTP 9992
    If ckcAmazon.Value = vbChecked Then
        If edcBucketName.Text = "" Or edcRegion.Text = "" Or edcAccessKey.Text = "" Or edcPrivateKey.Text = "" Then ckcAmazon.Value = vbUnchecked
    End If
    
    'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
    Dim ilNoPacingDays As Integer
    ilNoPacingDays = 0
    If edcPacing.Text = "" And edcPacingEnd.Text = "" Then ckcPacingRange.Value = vbUnchecked
    If ckcPacingRange.Value = vbChecked Then
        ilNoPacingDays = DateDiff("d", edcPacing.Text, edcPacingEnd.Text)
        If ilNoPacingDays = 0 Then ckcPacingRange.Value = vbUnchecked
        If ilNoPacingDays + 1 > 1 Then
            If ilNoPacingDays + 1 > 31 Then
                MsgBox "Invalid Selection:  Please select a range that generates no more than 31 days.", vbExclamation + vbOkOnly, "Date Range Pacing"
                edcPacingEnd.SetFocus
                Exit Sub
            End If
            ilRet = MsgBox("Confirm: The selectiong will generate " & ilNoPacingDays + 1 & " export files." & vbCrLf & "Okay to continue?", vbQuestion + vbYesNo, "Date Range Pacing")
            If ilRet = vbNo Then Exit Sub
        End If
    End If
        
    'Verify data input
    slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    slStr = ExpMatrix!edcMonth.Text             'month in text form (jan..dec, or 1-12
    gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
    If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
        ilSaveMonth = Val(slStr)
        ilRet = gVerifyInt(slStr, 1, 12)
        If ilRet = -1 Then

            ExpMatrix!edcNoMonths.SetFocus                 'invalid # periods
            'MsgBox "Month is Not Valid", vbOkOnly + vbApplicationModal, "Start Month"
            gAutomationAlertAndLogHandler "Month is Not Valid", vbOkOnly + vbApplicationModal, "Start Month"
            Exit Sub
        End If
    End If

    slFNMonth = Mid$(slMonthHdr, (ilSaveMonth - 1) * 3 + 1, 3)          'get the text month (jan...dec)
    slStr = ExpMatrix!edcYear.Text
    ilYear = gVerifyYear(slStr)
    If ilYear = 0 Then
        ExpMatrix!edcYear.SetFocus                 'invalid year
        gAutomationAlertAndLogHandler "Year is Not Valid", vbOkOnly + vbApplicationModal, "Start Year"
        Exit Sub
    End If

    slStr = ExpMatrix!edcNoMonths.Text            '#periods
    igPeriods = Val(slStr)
    ilRet = gVerifyInt(slStr, 1, 24)
    If ilRet = -1 Then
        ExpMatrix!edcNoMonths.SetFocus
        gAutomationAlertAndLogHandler "# months must be between 1 and 24", vbOkOnly + vbApplicationModal, "Number Months"
        Exit Sub
    End If

    lmCntrNo = 0                'ths is for debugging on a single contract
    slStr = ExpMatrix!edcContract
    If slStr <> "" Then
        lmCntrNo = Val(slStr)
    End If

    smExportName = Trim$(edcTo.Text)
    If Len(smExportName) = 0 Then
        Beep
        edcTo.SetFocus
        Exit Sub
    End If
    
    If (InStr(smExportName, ":") = 0) And (Left$(smExportName, 2) <> "\\") Then
        smExportName = Trim$(sgExportPath) & smExportName
    End If


    ilRet = 0
    ilRet = gFileExist(smExportName)
    If ilRet = 0 Then
        gAutomationAlertAndLogHandler "Filename already exists, enter new name", vbOkOnly + vbApplicationModal, "Save In"
        Exit Sub
    End If

    If Not mOpenMsgFile() Then          'open message file
         cmcCancel.SetFocus
         Exit Sub
    End If
    On Error GoTo 0
    ilRet = 0
    
    'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
    For imPacingDay = 0 To ilNoPacingDays
        tmcClick_Timer
        lacInfo(0).Visible = True
        smExportName = Trim$(edcTo.Text)
        If (InStr(smExportName, ":") = 0) And (Left$(smExportName, 2) <> "\\") Then
            smExportName = Trim$(sgExportPath) & smExportName
        End If
        Debug.Print smExportName
        ilRet = gFileOpen(smExportName, "Output", hmMatrix)
        If ilRet <> 0 Then
            'Print #hmMsg, "** Terminated **"
            gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Close #hmMatrix
            imExporting = False
            Screen.MousePointer = vbDefault
            'TTP 10011 - Error.Numner prevents MsgBox.  Additionally the Error # is stored in ilRet.
            gAutomationAlertAndLogHandler "Open Error #" & str$(ilRet) & " - " & smExportName, vbOkOnly, "Open Error"
            Exit Sub
        End If
        'Print #hmMsg, "** Storing Output into " & smExportName & " **"
        gAutomationAlertAndLogHandler "* Storing Output into " & smExportName
        If rbcMonthBy(0).Value = True Then gAutomationAlertAndLogHandler "* Calendar = Standard"
        If rbcMonthBy(1).Value = True Then gAutomationAlertAndLogHandler "* Calendar = Calendar Spots"
        If rbcMonthBy(2).Value = True Then gAutomationAlertAndLogHandler "* Calendar = Calendar Contract"
        If rbcMonthBy(3).Value = True Then gAutomationAlertAndLogHandler "* Calendar = Bill Method"
        gAutomationAlertAndLogHandler "* StartMonth =  " & edcMonth.Text
        gAutomationAlertAndLogHandler "* StartYear =  " & edcYear.Text
        gAutomationAlertAndLogHandler "* # months =  " & edcNoMonths.Text
        If ckcNTR.Value = vbChecked Then
            gAutomationAlertAndLogHandler "* Include NTR Revenue = True"
        Else
            gAutomationAlertAndLogHandler "* Include NTR Revenue = False"
        End If
        If ckcInclMissed.Value = vbChecked Then
            gAutomationAlertAndLogHandler "* Include Missed Spots = True"
        Else
            gAutomationAlertAndLogHandler "* Include Missed Spots = False"
        End If
        If ckcInclAdj.Value = vbChecked Then
            gAutomationAlertAndLogHandler "* Include Adjustments = True"
        Else
            gAutomationAlertAndLogHandler "* Include Adjustments = False"
        End If
        If ckcAll.Value = vbChecked Then
            gAutomationAlertAndLogHandler "* Include All Vehicles = True"
        Else
            gAutomationAlertAndLogHandler "* Include All Vehicles = False"
        End If
        'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
        If ckcPacingRange.Value = vbChecked Then
            gAutomationAlertAndLogHandler "* Pacing =  " & DateAdd("d", imPacingDay, edcPacing.Text)
        Else
            gAutomationAlertAndLogHandler "* Pacing =  " & edcPacing.Text
        End If
        gAutomationAlertAndLogHandler "* Contract# =  " & edcContract.Text
        
        If ckcAmazon.Value = vbChecked Then
            gAutomationAlertAndLogHandler "* AmazonBucket=True"
        Else
            gAutomationAlertAndLogHandler "* AmazonBucket=False"
        End If
    
        gAutomationAlertAndLogHandler "Exporting..."
        Me.Enabled = False
        lacInfo(0).Caption = "Exporting...": lacInfo(0).Refresh
        Screen.MousePointer = vbHourglass
        imExporting = True
    
        bmStdExport = True                          'assume standard exporting
        If Not rbcMonthBy(0).Value Then             'not std
            bmStdExport = False
            tmBillCycle.ilBillCycle = 1             'use Monthly Cal
        End If
        If rbcMonthBy(3).Value Then
            tmBillCycle.ilBillCycle = 2             'use Bill Cycle
        End If
        
        If mOpenMatrixFiles() = 0 Then
            '9-23-11 Build array of the vehicle group codes and names
            ilRet = gObtainMnfForType("H", slTimeStamp, tmMnfGroups())
            'Sort by code so that binary search can be used
            If UBound(tmMnfGroups) > 1 Then
                ArraySortTyp fnAV(tmMnfGroups(), 0), UBound(tmMnfGroups), 0, LenB(tmMnfGroups(0)), 0, -1, 0
            End If
            ReDim imUseCodes(0 To 1) As Integer 'zero index is ignored and in gFilterLists
            ilHowManyDefined = lbcVehicle.ListCount
            ilHowMany = lbcVehicle.SelCount
            If ilHowMany > (ilHowManyDefined / 2) + 1 Then   'more than half selected
                imIncludeCodes = False
            Else
                imIncludeCodes = True
            End If
        
            For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                slNameCode = tgUserVehicle(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If lbcVehicle.Selected(ilLoop) And imIncludeCodes Then               'selected ?
                    imUseCodes(UBound(imUseCodes)) = Val(slCode)
                    ReDim Preserve imUseCodes(0 To UBound(imUseCodes) + 1)
                Else        'exclude these
                    If (Not lbcVehicle.Selected(ilLoop)) And (Not imIncludeCodes) Then
                        imUseCodes(UBound(imUseCodes)) = Val(slCode)
                        ReDim Preserve imUseCodes(0 To UBound(imUseCodes) + 1)
                    End If
                End If
            Next ilLoop
            
            'Get Last Billed Std
            gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr  'convert last bdcst billing date to string
            tmBillCycle.lStdBillCycleLastBilled = gDateValue(slStr)            'convert last month billed to long
            slStart = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(str$(ilYear))
            'Get Std Bcast Cal Billing periods
            gBuildStartDates slStart, 1, igPeriods + 1, tmBillCycle.lStdBillCycleStartDates() 'build array of std start & end dates
            
            'Get Last Billed Cal
            gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slStr
            tmBillCycle.lCalBillCycleLastBilled = gDateValue(slStr)
            slStart = Trim$(str$(ilSaveMonth)) & "/01/" & Trim$(str(ilYear))
            'Get Calendar Billing periods
            gBuildStartDates slStart, 4, igPeriods + 1, tmBillCycle.lCalBillCycleStartDates() 'build array of std start & end dates
            
            'determine Last Billed indexes (Std)
            If tmBillCycle.lStdBillCycleLastBilled >= tmBillCycle.lStdBillCycleStartDates(igPeriods + 1) Then  'all in past
                tmBillCycle.iStdBillCycleLastBilledInx = igPeriods
            End If
            For ilLoop = 1 To igPeriods Step 1
                If tmBillCycle.lStdBillCycleLastBilled > tmBillCycle.lStdBillCycleStartDates(ilLoop) And tmBillCycle.lStdBillCycleLastBilled < tmBillCycle.lStdBillCycleStartDates(ilLoop + 1) Then
                    tmBillCycle.iStdBillCycleLastBilledInx = ilLoop
                    Exit For
                End If
            Next ilLoop
            
            'determine Last Billed indexes (Cal)
            'TTP 10870 - RAB Cal Spot: bypassing adjustments when "include adjustment" checkbox is checked on and the period the export is run for doesn't include any unbilled months
            'If tmBillCycle.lCalBillCycleLastBilled >= tmBillCycle.lCalBillCycleStartDates(igPeriods + 1) Then  'all in past
            If tmBillCycle.lStdBillCycleLastBilled >= tmBillCycle.lCalBillCycleStartDates(igPeriods + 1) Then  'all in past
                tmBillCycle.iCalBillCycleLastBilledInx = igPeriods
            End If
            For ilLoop = 1 To igPeriods Step 1
                If tmBillCycle.lStdBillCycleLastBilled > tmBillCycle.lCalBillCycleStartDates(ilLoop) And tmBillCycle.lStdBillCycleLastBilled < tmBillCycle.lCalBillCycleStartDates(ilLoop + 1) Then
                    tmBillCycle.iCalBillCycleLastBilledInx = ilLoop
                    Exit For
                End If
            Next ilLoop
            
            If bmStdExport Then         'standard exporting
                If tmBillCycle.lStdBillCycleStartDates(1) > tmBillCycle.lStdBillCycleLastBilled And (Not rbcNetBy(1).Value) Then         'projection only if dates all in future and by net values (not tnet)
                    ilRet = mCrMatrixProj()
                Else
                    ilRet = mCrMatrixPast()
                    If ilRet = 0 Then              '0 = ok
                        If tmBillCycle.lStdBillCycleLastBilled < tmBillCycle.lStdBillCycleStartDates(igPeriods) Or lmPacingDate > 0 Then                    'past only or past & projection
                            ilRet = mCrMatrixProj()
                        End If
                    End If
                End If
            Else                        'calendar exporting.
                If imExportOption = EXP_RAB Then        '6-17-20 rab has both cal spots and cal contract exports (originally just cal contract)
                    If rbcMonthBy(1).Value Then         'cal by spots
                        ilRet = mCrMatrixPast()
                        mCrCalendarMatrix tmBillCycle.lCalBillCycleStartDates() 'get the spots for export
                        ilRet = 0                       'indicates successful export message
                    ElseIf rbcMonthBy(2).Value Then     'cal by contract
                        ilRet = mCrMatrixPast()         'adjustment
                        ilRet = mCrMatrixProj()
                    Else                                'By Bill Method
                        ilRet = mCrMatrixPast()         'adjustment
                        ilRet = mCrMatrixProj()
                    End If
                    
                ElseIf imExportOption = EXP_CUST_REV Then        'TTP 9992
                    If rbcMonthBy(1).Value Then         'cal by spots
                        ilRet = mCrMatrixPast()
                        mCrCalendarMatrix tmBillCycle.lCalBillCycleStartDates() 'get the spots for export
                        ilRet = 0                       'indicates successful export message
                    End If
                
                Else '(EXP_MATRIX,EXP_TABLEAU)
                    ilRet = mCrMatrixPast()
                    mCrCalendarMatrix tmBillCycle.lCalBillCycleStartDates() 'get the spots for export
                    ilRet = 0                           'indicates successful export message
                End If
            End If
            
            Close #hmMatrix
            mCloseMatrixFiles
            'Erase llStdStartDates
            Screen.MousePointer = vbDefault
        Else
            lacInfo(0).Caption = "Open Error: Export Failed": lacInfo(0).Refresh
            'Print #hmMsg, "** Export Open error : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            gAutomationAlertAndLogHandler "** Export Open error : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        End If
    
        If ilRet = 0 Then           'true is successful
            'lacInfo(0).Caption = "Export Matrix Successfully Completed"
            lacInfo(0).Caption = "Export " & Trim$(smExportOptionName) & " Successfully Completed"
    
            'Print #hmMsg, "** Export Matrix Successfully completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            
            'TTP 9992 - AMAZON BUCKET SUPPORT:
            If ckcAmazon.Value = vbChecked And edcBucketName.Text <> "" And edcRegion.Text <> "" And edcAccessKey.Text <> "" And edcPrivateKey.Text <> "" Then
                If lgExportCount > 0 Then
                    'Print #hmMsg, "** Uploading " & smExportFilename & " to " & edcBucketName.Text & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                    'TTP 10504 - Amazon web bucket upload cleanup
                    gAutomationAlertAndLogHandler "** Uploading " & smExportFilename & " to " & AmazonBucketFolder(edcBucketName.Text, edcAmazonSubfolder.Text) & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                    lacInfo(0).Caption = "Uploading " & smExportFilename
                    lacInfo(0).Refresh
                    DoEvents
                    Set myBucket = New CsiToAmazonS3.ApiCaller
                    On Error Resume Next
                    err = 0
                    'TTP 10504 - Amazon web bucket upload cleanup
                    myBucket.UploadAmazonBucketFile AmazonBucketFolder(edcBucketName.Text, edcAmazonSubfolder.Text), edcRegion.Text, edcAccessKey.Text, edcPrivateKey.Text, edcTo.Text, False
                    
                    If err <> 0 Then
                        lacInfo(0).Caption = "Error Uploading " & smExportFilename & " - " & err & " - " & Error(err)
                        'Print #hmMsg, "** Error Uploading " & smExportFilename & " - " & err & " - " & Error(err) & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                        gAutomationAlertAndLogHandler "** Error Uploading " & smExportFilename & " - " & err & " - " & Error(err) & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                    Else
                        If myBucket.ErrorMessage <> "" Then
                            lacInfo(0).Caption = "Error Uploading " & smExportFilename
                            'Print #hmMsg, "** Error Uploading " & smExportFilename & " - " & Replace(myBucket.ErrorMessage, vbCrLf, ";") & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                            gAutomationAlertAndLogHandler "** Error Uploading " & smExportFilename & " - " & Replace(myBucket.ErrorMessage, vbCrLf, ";") & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                        Else
                            lacInfo(0).Caption = "Sucess Uploading " & smExportFilename
                            'Print #hmMsg, "** Finished Uploading " & smExportFilename & " - " & Replace(myBucket.Message, vbCrLf, ";") & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                            gAutomationAlertAndLogHandler "** Finished Uploading " & smExportFilename & " - " & Replace(myBucket.Message, vbCrLf, ";") & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                            If ckcKeepLocalFile.Value = vbUnchecked Then
                                'We want to remove the Local File
                                Kill edcTo.Text
                                'Print #hmMsg, "** Deleted Local Export File : " & smExportName & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                gAutomationAlertAndLogHandler "** Deleted Local Export File : " & smExportName & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                            End If
                        End If
                    End If
                    Set myBucket = Nothing
                Else
                    'Print #hmMsg, "** Nothing to Upload to Amazon, Record Count : " & lgExportCount & " **"
                    gAutomationAlertAndLogHandler "** Nothing to Upload to Amazon, Record Count : " & lgExportCount & " **"
                    lacInfo(0).Caption = "Nothing to Upload, Record Count : " & lgExportCount
                End If
                'Print #hmMsg, "** Export " & Trim$(smExportOptionName) & " Successfully completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " , Record Count : " & lgExportCount & " **"
                gAutomationAlertAndLogHandler "** Export " & Trim$(smExportOptionName) & " Successfully completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " , Record Count : " & lgExportCount & " **"
            End If
        Else
            lacInfo(0).Caption = "Export Failed"
            'Print #hmMsg, "** Export Failed **"
            gAutomationAlertAndLogHandler "** Export Failed **"
        End If
            
        gAutomationAlertAndLogHandler "** Export Procedure complete **"
    Next imPacingDay 'End of TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
    lacInfo(0).Visible = True
    Close #hmMsg
    
    Me.Enabled = True
    cmcExport.Enabled = False
    cmcCancel.Caption = "&Done"
    If igExportType <= 1 Then       'ok to set focus if manual mode
        cmcCancel.SetFocus
    End If
    Screen.MousePointer = vbDefault
    imExporting = False
    Exit Sub

ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    Me.Enabled = True
End Sub

Private Sub cmcTo_Click()
    CMDialogBox.DialogTitle = "Export To File"
    CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
    CMDialogBox.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
    CMDialogBox.DefaultExt = ".Csv"
    CMDialogBox.flags = cdlOFNCreatePrompt
    CMDialogBox.Action = 1 'Open dialog
    edcTo.Text = CMDialogBox.fileName
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
    If edcTo.Text = "" Then
        edcTo.Text = smExportName
    End If
End Sub

Private Sub edcContract_GotFocus()
    gCtrlGotFocus edcContract
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcMonth_Change()
    Dim slStr As String
    If Len(edcMonth) = 3 Then
        gCtrlGotFocus edcMonth
        If igExportType <= 1 Then
            tmcClick_Timer
        End If
    End If
End Sub

Private Sub edcMonth_Click()
    gCtrlGotFocus edcMonth
End Sub

Private Sub edcMonth_GotFocus()
    gCtrlGotFocus edcMonth
End Sub

Private Sub edcMonth_LostFocus()
    If igExportType <= 1 Then
        tmcClick_Timer
    End If
End Sub

Private Sub edcNoMonths_GotFocus()
    gCtrlGotFocus edcNoMonths
End Sub

Private Sub edcNoMonths_Change()
    If Val(edcNoMonths.Text) > 0 And Val(edcNoMonths.Text) < 25 Then
        'gCtrlGotFocus edcNoMonths
        If igExportType <= 1 Then
            tmcClick_Timer
        End If
    End If
End Sub

Private Sub edcPacing_Change()
'    'TTP 10163 - Include Adj, is Irrelevant for Pacing...
'    If imExportOption = EXP_CUST_REV And rbcMonthBy(0).Value = True And edcPacing.Text <> "" Then
'        ckcInclAdj.Enabled = False
'    Else
'        ckcInclAdj.Enabled = True
'        ckcInclAdj.Value = True
'    End If
End Sub

Private Sub edcPacing_GotFocus()
    gCtrlGotFocus edcPacing
End Sub

Private Sub edcPacing_LostFocus()
    If edcPacing.Text = "" Then Exit Sub
    If Not gValidDate(edcPacing.Text) Then
        edcPacing.SetFocus
    End If
End Sub

'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
Private Sub edcPacingEnd_GotFocus()
    gCtrlGotFocus edcPacingEnd
End Sub

Private Sub edcPacingEnd_LostFocus()
    If edcPacingEnd.Text = "" Then Exit Sub
    If Not gValidDate(edcPacingEnd.Text) Then
        edcPacingEnd.SetFocus
    End If
End Sub

Private Sub edcTo_Change()
    'get Filename from the full path and filename (used as the object name to upload to Amazon)
    Dim lsFilename As String
    Dim liSeparator As Integer
    lsFilename = edcTo.Text
    If lsFilename = "" Then Exit Sub
    liSeparator = InStrRev(lsFilename, "\")
    smExportFilename = Mid(lsFilename, liSeparator + 1)
    lacExportFilename.Caption = smExportFilename
    cmcExport.Enabled = True
End Sub

Private Sub edcTo_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
    End If

    lacInfo(0).Visible = False
    lacInfo(1).Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcYear_Change()
    If Len(edcYear.Text) = 4 Then
        gCtrlGotFocus edcYear
        If igExportType <= 1 Then
            tmcClick_Timer
        End If
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
    DoEvents    'Process events so pending keys are not sent to this
    plcMonthBy.Visible = False
    plcMonthBy.Visible = True
    PlcNetBy.Visible = False
    PlcNetBy.Visible = True
    rbcMonthBy(3).Visible = False
    plcMonthBy.Height = 240
    If imExportOption = EXP_RAB Then
        PlcNetBy.Visible = False
        ckcNTR.Value = vbChecked
        rbcMonthBy(3).Visible = True
        plcMonthBy.Height = 480
        
    ElseIf imExportOption = EXP_CUST_REV Then
        'TTP 9992
        PlcNetBy.Visible = False
        ckcNTR.Value = vbChecked
        'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
        ckcPacingRange.Visible = True
    Else
        If (Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER Then
            'show Tnet, and default to TNet is acquisition used
            If rbcNetBy(1).Value Then
                rbcNetBy_Click 1
            Else
                rbcNetBy(1).Value = True
            End If
        Else
            rbcNetBy(0).Value = True
            PlcNetBy.Visible = False
        End If
        rbcMonthBy(2).Visible = False
    End If
            

    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub

Private Sub Form_Load()
    mInit
    'igExportType As Integer  '0=Manual; 1=From Traffic, 2=Auto-Efficio Projection; 3=Auto-Efficio Revenue; 4=Auto-Matrix, 6 = auto rab
    If igExportType <= 1 Then                       'manual from exports or manual from traffic
        Me.WindowState = vbNormal
        If imExportOption = EXP_MATRIX Then
            If ((Asc(tgSpf.sUsingFeatures) And MATRIXEXPORT) <> MATRIXEXPORT) And ((Asc(tgSaf(0).sFeatures1) And MATRIXCAL) <> MATRIXCAL) Then
                gAutomationAlertAndLogHandler "** Matrix Export Disabled:  " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                If igExportType <= 1 Then           'manual mode, show disallowed on screen
                    lacInfo(0).AddItem "Matrix Export Disabled"
                End If
                imTerminate = True
                Exit Sub
            Else
                cmcExport.Enabled = True
            End If
        ElseIf imExportOption = EXP_TABLEAU Then
            If ((Asc(tgSaf(0).sFeatures2) And TABLEAUEXPORT) <> TABLEAUEXPORT) And ((Asc(tgSaf(0).sFeatures2) And TABLEAUCAL) <> TABLEAUCAL) Then
                gAutomationAlertAndLogHandler "** Tableau Export Disabled:  " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                If igExportType <= 1 Then           'manual mode, show disallowed on screen
                    lacInfo(0).AddItem "Tableau Export Disabled"
                End If
                imTerminate = True
                Exit Sub
            Else
                cmcExport.Enabled = True
            End If
        ElseIf imExportOption = EXP_RAB Then                                 '1-23-20
            If ((Asc(tgSaf(0).sFeatures6) And RABCALENDAR) <> RABCALENDAR) And ((Asc(tgSaf(0).sFeatures7) And RABSTD) <> RABSTD) And ((Asc(tgSaf(0).sFeatures7) And RABCALSPOTS) <> RABCALSPOTS) Then
                gAutomationAlertAndLogHandler "** RAB Export Disabled:  " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                If igExportType <= 1 Then           'manual mode, show disallowed on screen
                    lacInfo(0).AddItem "RAB Export Disabled"
                End If
                imTerminate = True
                Exit Sub
            Else
                cmcExport.Enabled = True
            End If
        ElseIf imExportOption = EXP_CUST_REV Then                                 'TTP 9992
            If (Asc(tgSaf(0).sFeatures7) And CUSTOMEXPORT) <> CUSTOMEXPORT Then
                gAutomationAlertAndLogHandler "** Custom Revenue Export Disabled:  " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                If igExportType <= 1 Then           'manual mode, show disallowed on screen
                    lacInfo(0).AddItem "Custom Revenue Export Disabled"
                End If
                imTerminate = True
                Exit Sub
            Else
                cmcExport.Enabled = True
            End If
        End If
    Else
        Me.WindowState = vbMinimized
        If imExportOption = EXP_RAB Then                                 '1-23-20
            If ((rbcMonthBy(0).Value = True) And ((Asc(tgSaf(0).sFeatures7) And RABSTD) <> RABSTD)) Or ((rbcMonthBy(1).Value = True) And ((Asc(tgSaf(0).sFeatures7) And RABCALSPOTS) <> RABCALSPOTS)) Or ((rbcMonthBy(2).Value = True) And ((Asc(tgSaf(0).sFeatures6) And RABCALENDAR) <> RABCALENDAR)) Then
                If ((Asc(tgSaf(0).sFeatures7) And RABSTD) <> RABSTD Or (Asc(tgSaf(0).sFeatures7) And RABSTD) <> RABSTD Or (Asc(tgSaf(0).sFeatures7) And RABCALSPOTS) <> RABCALSPOTS) Then              'std allowed?
                    gLogMsg "** " & Trim$(smExportOptionName) & " Export option disabled", "Exp" & Trim$(smExportOptionName) & ".txt", False   'exprab.txt
                    imTerminate = True
                End If
            End If
        Else
            If (rbcMonthBy(0).Value = True) Then
                If ((imExportOption = EXP_MATRIX) And ((Asc(tgSpf.sUsingFeatures) And MATRIXEXPORT) <> MATRIXEXPORT)) Or ((imExportOption = EXP_TABLEAU) And ((Asc(tgSaf(0).sFeatures2) And TABLEAUEXPORT) <> TABLEAUEXPORT)) Then
                    gLogMsg "** " & Trim$(smExportOptionName) & " Standard Export option disabled", "Exp" & Trim$(smExportOptionName) & ".txt", False   'expmatrix.txt or exptableau.txt
                    imTerminate = True
                End If
            End If
            If (rbcMonthBy(1).Value = True) Then        'cal matrix selected
                If ((imExportOption = EXP_MATRIX) And ((Asc(tgSaf(0).sFeatures1) And MATRIXCAL) <> MATRIXCAL)) Or ((imExportOption = EXP_TABLEAU) And ((Asc(tgSaf(0).sFeatures2) And TABLEAUCAL) <> TABLEAUCAL)) Then     'allowed?
                    gLogMsg "** " & Trim$(smExportOptionName) & " Calendar Export option disabled", "Exp" & Trim$(smExportOptionName) & ".txt", False   'expmatrix.txt or exptableau.txt
                    imTerminate = True
                End If
            End If
        End If
        
        If Not imTerminate Then     'if any errors and terminate set, do not export
            If igExportType = 4 Then                       'auto(4) from exports or manual from traffic
                gOpenTmf
                tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
                tmcSetTime.Enabled = True
                gUpdateTaskMonitor 1, "ME"
                cmcExport_Click
                gUpdateTaskMonitor 2, "ME"
            End If
            If igExportType = 5 Then                       'tableau auto(4) from exports or manual from traffic
                gOpenTmf
                tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
                tmcSetTime.Enabled = True
                gUpdateTaskMonitor 1, "TE"
                cmcExport_Click
                gUpdateTaskMonitor 2, "TE"
            End If
            If igExportType = 6 Then                       'RAB auto from exports or manual from traffic
                gOpenTmf
                tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
                tmcSetTime.Enabled = True
                gUpdateTaskMonitor 1, "RE"
                cmcExport_Click
                gUpdateTaskMonitor 2, "RE"
            End If
            If igExportType = 7 Then                       'CustomRevenueExport auto from exports or manual from traffic
                gOpenTmf
                tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
                tmcSetTime.Enabled = True
                gUpdateTaskMonitor 1, "RE"
                cmcExport_Click
                gUpdateTaskMonitor 2, "RE"
            End If
        End If
        imTerminate = True
    End If
    tmcClick.Interval = 2000    '2 seconds
    tmcClick.Enabled = True
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tgAcqComm
    Erase tgAcqCommInx
    
    If igExportType = 4 Or igExportType = 5 Or igExportType = 6 Then         '1-29-20 unload if matrix , tableau,  rab exports
        tmcSetTime.Enabled = False
        gCloseTmf
    End If
    
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmPrf)
    ilRet = btrClose(hmVff)
    ilRet = btrClose(hmCef)
    ilRet = btrClose(hmPcf)
    
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmVef
    btrDestroy hmMnf
    btrDestroy hmSof
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmAgf
    btrDestroy hmSbf
    btrDestroy hmPrf
    btrDestroy hmVff
    btrDestroy hmCef
    btrDestroy hmPcf
    
    Set ExpMatrix = Nothing   'Remve data segment
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim slDay As String
    Dim slMonth As String
    Dim slYear As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim slMonthStr As String * 36
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilVff As Integer
    Dim ilLoop As Integer
    Dim slLocation As String
    Dim slReturn As String * 130
    Dim slFileName As String
    Dim ilVefInx As Integer
    Dim StartPeriod As Integer
    Dim StartMonth As Integer
    
    slMonthStr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmNowDate = gDateValue(Format$(gNow(), "m/d/yy"))

    Dim slDateTime As String
    slDateTime = Format$(gNow(), "m/d/yy") & " " & Format(Now, "hh:mm:ss AMPM")
    
    gCenterStdAlone ExpMatrix
    smMonthBy = "Month By"
    '7-9-15 implement Tableau option which has same format as Matrix
    imExportOption = ExportList!lbcExport.ItemData(ExportList!lbcExport.ListIndex)
    If imExportOption = EXP_MATRIX Then
        smExportOptionName = "Matrix"
    ElseIf imExportOption = EXP_TABLEAU Then
        smExportOptionName = "Tableau"
    ElseIf imExportOption = EXP_RAB Then
        smExportOptionName = "RAB"                  '1-23-20
    ElseIf imExportOption = EXP_CUST_REV Then
        smExportOptionName = "CustomRevenueExport"
    Else
        smExportOptionName = ""
    End If
    
    gAutomationAlertAndLogHandler "* OptionName=" & smExportOptionName 'log
    
    ilRet = gObtainAdvt()   'Build into tgCommAdf
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gObtainAgency() 'Build into tgCommAgf
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gObtainVef() 'Build into tgMVef
    If ilRet = False Then
        imTerminate = True
    End If
    
    ilRet = gObtainSalesperson() 'Build into tgMSlf
    If ilRet = False Then
        imTerminate = True
    End If
    
    ilRet = gBuildAcqCommInfo(ExpMatrix)
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf)", ExpMatrix
    imMnfRecLen = Len(tmMnf)

    hmVff = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVff, "", sgDBPath & "Vff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vff)", ExpMatrix
    imVffRecLen = Len(tmVff)
    
    If igExportType >= 4 And igExportType <= 7 Then
                                        'as a timing issue prevents the filename from showing in the text box.
                                        'igExportType As Integer  '0=Manual; 1=From Traffic, 2=Auto-Efficio Projection; 3=Auto-Efficio Revenue; 4=Auto-Matrix, 6 = rab, 7=CustomRevenueExport
        smClientName = Trim$(tgSpf.sGClient)
        If tgSpf.iMnfClientAbbr > 0 Then
            tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                smClientName = Trim$(tmMnf.sName)
            End If
        End If
    End If
    
    'determine default month year
    slDate = Format$(lmNowDate, "m/d/yy")
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    'Default to last month, based on today's date
    If Val(slMonth) = 1 Then
        ilMonth = 12
        ilYear = Val(slYear) - 1
    Else
        ilMonth = Val(slMonth) - 1
        ilYear = Val(slYear)
    End If

    edcMonth.Text = Mid$(slMonthStr, (ilMonth - 1) * 3 + 1, 3)
    edcYear.Text = Trim$(str$(ilYear))
    rbcMonthBy(3).Visible = False
    If igExportType <= 1 Then                    'igExportType As Integer  '0=Manual; 1=From Traffic, 2=Auto-Efficio Projection; 3=Auto-Efficio Revenue; 4=Auto-Matrix, 5 = tableau, 6 = rab
        If Trim$(smExportOptionName) = "Matrix" Then
            If (Asc(tgSpf.sUsingFeatures) And MATRIXEXPORT) = MATRIXEXPORT Then     'std
                rbcMonthBy(0).Value = True
                If (Asc(tgSaf(0).sFeatures1) And MATRIXCAL) <> MATRIXCAL Then
                    rbcMonthBy(1).Enabled = False
                End If
            End If
            If (Asc(tgSaf(0).sFeatures1) And MATRIXCAL) = MATRIXCAL Then
                rbcMonthBy(1).Value = True
                If (Asc(tgSpf.sUsingFeatures) And MATRIXEXPORT) <> MATRIXEXPORT Then
                    rbcMonthBy(0).Enabled = False
                End If
            End If
        ElseIf Trim$(smExportOptionName) = "Tableau" Then
            '6-9-15 implement TAbleau export
            If (Asc(tgSaf(0).sFeatures2) And TABLEAUEXPORT) = TABLEAUEXPORT Then     'std
                rbcMonthBy(0).Value = True
                If (Asc(tgSaf(0).sFeatures2) And TABLEAUCAL) <> TABLEAUCAL Then
                    rbcMonthBy(1).Enabled = False
                End If
            End If
            If (Asc(tgSaf(0).sFeatures2) And TABLEAUCAL) = TABLEAUCAL Then     'std
                rbcMonthBy(1).Value = True
                If (Asc(tgSaf(0).sFeatures2) And TABLEAUEXPORT) <> TABLEAUEXPORT Then
                    rbcMonthBy(0).Enabled = False
                End If
            End If
        ElseIf Trim$(smExportOptionName) = "RAB" Then                                         '1-23-20
            PlcNetBy.Visible = False
            rbcNetBy(0).Value = True
            rbcMonthBy(0).Enabled = False
            rbcMonthBy(1).Enabled = False
            rbcMonthBy(2).Enabled = False
            rbcMonthBy(3).Enabled = False
            If (Asc(tgSaf(0).sFeatures6) And RABCALENDAR) = RABCALENDAR Then
                ckcInclMissed.Visible = False
                rbcMonthBy(2).Enabled = True
            End If
            If (Asc(tgSaf(0).sFeatures7) And RABSTD) = RABSTD Then                          '6-19-20 added std
                ckcInclMissed.Visible = False
                rbcMonthBy(0).Enabled = True
            End If
            If (Asc(tgSaf(0).sFeatures7) And RABCALSPOTS) = RABCALSPOTS Then
                ckcInclMissed.Visible = False
                rbcMonthBy(1).Enabled = True
            End If
            'Check to see if std and cal spots are check on, then show "Bill Method"
            If (Asc(tgSaf(0).sFeatures7) And RABCALSPOTS) = RABCALSPOTS And (Asc(tgSaf(0).sFeatures7) And RABSTD) = RABSTD Then
                rbcMonthBy(3).Enabled = True
                rbcMonthBy(3).Visible = True
            End If
            'hierachy of what month type gets defaulted
            If rbcMonthBy(0).Enabled Then       'std
                rbcMonthBy(0).Value = True
            ElseIf rbcMonthBy(2).Enabled = True Then        ' cal spots
                rbcMonthBy(2).Value = True
            Else
                rbcMonthBy(1).Value = True              'cal spts
            End If
       ElseIf Trim$(smExportOptionName) = "CustomRevenueExport" Then
            PlcNetBy.Visible = False
            rbcNetBy(0).Value = True
            rbcMonthBy(0).Enabled = True
            rbcMonthBy(1).Enabled = True
            rbcMonthBy(2).Enabled = False
            rbcMonthBy(2).Visible = False
            
            'hierachy of what month type gets defaulted
            If rbcMonthBy(0).Enabled Then       'std
                rbcMonthBy(0).Value = True
            Else
                rbcMonthBy(1).Value = True              'cal spts
            End If
        End If
    Else
        On Error GoTo mObtainIniValuesErr
       'determine if coming from auto mode and to get parameters from Exports.ini
       'find exports.ini
        sgIniPath = gSetPathEndSlash(sgIniPath, True)
        If igDirectCall = -1 Then
            slFileName = sgIniPath & "Exports.Ini"
        Else
            slFileName = CurDir$ & "\Exports.Ini"
        End If
        'look for matching sections which has the parameter options
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Calendar", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            If smExportOptionName = "RAB" Then
                '6-17-20 auto export - determine the month options for RAB
                If InStr(1, slReturn, "Std", vbTextCompare) > 0 Then
                    rbcMonthBy(0).Value = True                              'std
                ElseIf InStr(1, slReturn, "Cal", vbTextCompare) > 0 Or InStr(1, slReturn, "Spot", vbTextCompare) > 0 Then
                    rbcMonthBy(1).Value = True                              'cal month by spots
                ElseIf InStr(1, slReturn, "CCnt", vbTextCompare) > 0 Then
                    rbcMonthBy(2).Value = True                              'CCnt" cal by contract
                ElseIf InStr(1, slReturn, "BillMethod", vbTextCompare) > 0 Then
                    rbcMonthBy(3).Value = True                              'Bill Method
                End If
            End If
        Else
            If InStr(1, slReturn, "Std", vbTextCompare) > 0 Then
                rbcMonthBy(0).Value = True
                '6-17-20 3 different month methods std = standard month (use history and contracts), Cal = cal spots, CCnt = Cal month by contracts
            ElseIf InStr(1, slReturn, "Cal", vbTextCompare) > 0 Then     'cal month by spots
                rbcMonthBy(1).Value = True
            ElseIf InStr(1, slReturn, "CCnt", vbTextCompare) > 0 Then
                If smExportOptionName = "CustomRevenueExport" Then
                    rbcMonthBy(0).Value = True                              'Standard BCast Cal
                Else
                    rbcMonthBy(2).Value = True                              'cal month by contract
                End If
            ElseIf InStr(1, slReturn, "BillMethod", vbTextCompare) > 0 Then
                rbcMonthBy(3).Value = True                              'Bill Method
            End If
        End If
        
        If rbcMonthBy(0).Value = True Then
            gAutomationAlertAndLogHandler "* Option Calendar=" & rbcMonthBy(0).Caption 'log
        ElseIf rbcMonthBy(1).Value = True Then
            gAutomationAlertAndLogHandler "* Option Calendar=" & rbcMonthBy(1).Caption 'log
        ElseIf rbcMonthBy(2).Value = True Then
            gAutomationAlertAndLogHandler "* Option Calendar=" & rbcMonthBy(2).Caption  'log
        ElseIf rbcMonthBy(3).Value = True Then
            gAutomationAlertAndLogHandler "* Option Calendar=" & rbcMonthBy(4).Caption  'log
        End If
        
        '3/23/2021 - TTP 10121 - add support for new parameter "StartMonth" on Exports.ini, however it should override/circumvent the StartPeriod logic, therefore TTP 9949 is Moved below in StartMonth=0 condition
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "StartMonth", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'Use Default Period, or perhaps use StartPeriod (below)
            StartMonth = 0
        Else
            StartMonth = Val(Trim$(gStripChr0(slReturn)))
            If StartMonth > 0 And StartMonth < 13 Then
               edcMonth.Text = MonthName(StartMonth, True)
            Else
                'Not Valid StartMonth
                StartMonth = 0
            End If
        End If
        
        '-------------------------------------
        If StartMonth = 0 Then
            '11/2/2020 - TTP 9949 - allow 2 month start period (automation), add # of months to filename. Exports.ini "StartPeriod=#" - if StartPeriod=2 then the Start Month / Year is set to 2 Months [StdBC months or Cal months] Prior to current date / passed in date
            ilRet = GetPrivateProfileString(sgExportIniSectionName, "StartPeriod", "Not Found", slReturn, 128, slFileName)
            If Left$(slReturn, ilRet) = "Not Found" Then
                'Use the Start Month/Year that's already defaulted (1 month prior to today)
            Else
                StartPeriod = Trim$(gStripChr0(slReturn))
                If StartPeriod > 1 Then
                    If rbcMonthBy(0).Value = True Then 'Standard
                       edcMonth.Text = MonthName(Month(DateAdd("m", -StartPeriod, gObtainEndStd(Format$(lmNowDate, "m/d/yy")))), True)
                       edcYear.Text = Year(DateAdd("m", -StartPeriod, gObtainEndStd(Format$(lmNowDate, "m/d/yy"))))
                    End If
                    If rbcMonthBy(1).Value = True Or rbcMonthBy(2).Value = True Then 'Calendar Spots 'Calendar Contract
                       edcMonth.Text = MonthName(Month(DateAdd("m", -StartPeriod, Format$(lmNowDate, "m/d/yy"))), True)
                       edcYear.Text = Year(DateAdd("m", -StartPeriod, Format$(lmNowDate, "m/d/yy")))
                    End If
                End If
            End If
        End If
        
        gAutomationAlertAndLogHandler "* Option Month=" & edcMonth.Text 'log
        gAutomationAlertAndLogHandler "* Option Year=" & edcYear.Text 'log
        
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Dollars", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            rbcNetBy(0).Value = True            'default to Net
        Else
            If InStr(1, slReturn, "TNet", vbTextCompare) > 0 Then
                rbcNetBy(1).Value = True
            Else
                rbcNetBy(0).Value = True
            End If
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Months", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcNoMonths.Text = 24                   '11-3-14 change default from 13 to 24 if not found
        Else
            slCode = Trim$(gStripChr0(slReturn))
            If Val(slCode) = 0 Or Val(slCode) > 24 Then     'max 24 months
                edcNoMonths.Text = 24                       'invalid input , take default of 24 weeks
            Else
                edcNoMonths.Text = Trim$(gStripChr0(slReturn))
            End If
        End If
        
        gAutomationAlertAndLogHandler "* Option Months=" & edcNoMonths.Text 'log
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "NTR", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            ckcNTR.Value = vbChecked
        Else
            If InStr(1, slReturn, "No", vbTextCompare) > 0 Then  'found No to exclude NTR
                ckcNTR.Value = vbUnchecked
            Else
                ckcNTR.Value = vbChecked
            End If
        End If
        gAutomationAlertAndLogHandler "* Option NTR=" & ckcNTR.Value  'log
        
        '6-23-15 Option to include/exclude missed spots
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Missed", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            ckcInclMissed.Value = vbUnchecked           'default to exclude Missed
        Else
            If InStr(1, slReturn, "No", vbTextCompare) > 0 Then  'found No to exclude Missed
                ckcInclMissed.Value = vbUnchecked
            Else
                ckcInclMissed.Value = vbChecked
            End If
        End If
        gAutomationAlertAndLogHandler "* Option InclMissed=" & ckcInclMissed.Value 'log
        
        '6-30-20 Option to use adjustments
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Adjustments", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            ckcInclAdj.Value = vbUnchecked           'default to exclude adjustments
        Else
            If InStr(1, slReturn, "No", vbTextCompare) > 0 Then  'found No to exclude Missed
                ckcInclAdj.Value = vbUnchecked
            Else
                ckcInclAdj.Value = vbChecked
            End If
        End If
        gAutomationAlertAndLogHandler "* Option InclAdj=" & ckcInclAdj.Value 'log

        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Export", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'default to the export path
            sgExportPath = sgExportPath
        Else
            sgExportPath = Trim$(gStripChr0(slReturn))
        End If
        sgExportPath = gSetPathEndSlash(sgExportPath, True)
        gAutomationAlertAndLogHandler "* Option ExportPath=" & sgExportPath 'log

        'TTP 9992 - Amazon Support
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "BucketName", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcBucketName.Text = ""
        Else
            slCode = Trim$(gStripChr0(slReturn))
            edcBucketName.Text = slCode
        End If
        If ckcAmazon.Value = vbChecked Then
            gAutomationAlertAndLogHandler "* AmazonBucket=True"
        Else
            gAutomationAlertAndLogHandler "* AmazonBucket=False"
        End If

        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "BucketFolder", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcAmazonSubfolder.Text = ""
        Else
            slCode = Trim$(gStripChr0(slReturn))
            edcAmazonSubfolder.Text = slCode
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Region", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcRegion.Text = ""
        Else
            slCode = Trim$(gStripChr0(slReturn))
            edcRegion.Text = slCode
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "AccessKey", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcAccessKey.Text = ""
        Else
            slCode = Trim$(gStripChr0(slReturn))
            edcAccessKey.Text = slCode
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "PrivateKey", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcPrivateKey.Text = ""
        Else
            slCode = Trim$(gStripChr0(slReturn))
            edcPrivateKey.Text = slCode
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "KeepLocalFile", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            ckcKeepLocalFile.Value = vbUnchecked           'delete local file after upload to Amazon
        Else
            If InStr(1, slReturn, "Yes", vbTextCompare) > 0 Then
                ckcKeepLocalFile.Value = vbChecked         'Keep local file
            Else
                ckcKeepLocalFile.Value = vbUnchecked       'delete local file after upload to Amazon
            End If
        End If
        'TTP 9992
        If edcBucketName.Text <> "" And edcRegion.Text <> "" And edcAccessKey.Text <> "" And edcPrivateKey.Text <> "" Then
            'If INI provides all 4 AWS values then Check Amazon
            ckcAmazon.Value = vbChecked
        Else
            ckcAmazon.Value = vbUnchecked
        End If
    End If
    
    imSetAll = True
    If igExportType >= 4 And igExportType <= 7 Then     'need to setup the filename if background mode
        tmcClick_Timer
    End If
    
    smClientName = Trim$(tgSpf.sGClient)
    If tgSpf.iMnfClientAbbr > 0 Then
        tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smClientName = Trim$(tmMnf.sName)
        End If
    End If

    Screen.MousePointer = vbDefault
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    Exit Sub

mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub

mObtainIniValuesErr:
    Resume Next

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
    Unload ExpMatrix
    igManUnload = NO
End Sub

Private Sub lbcVehicle_Click()
  If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
    End If
   ' mSetCommands
End Sub

Private Sub plcMonthBy_Paint()
    plcMonthBy.Cls
    plcMonthBy.CurrentX = 0
    plcMonthBy.CurrentY = 0
    plcMonthBy.Print smMonthBy
End Sub

Private Sub PlcNetBy_Paint()
    PlcNetBy.CurrentX = 0
    PlcNetBy.CurrentY = 0
    PlcNetBy.Print "Export Net"
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    'TTP 9992
    If smExportOptionName = "CustomRevenueExport" Then
        plcScreen.Print "Custom Revenue Export"
    Else
        plcScreen.Print smExportOptionName & " Export"
    End If
End Sub

Private Sub rbcMonthBy_Click(Index As Integer)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVefCode As Integer
    Dim ilVefInx As Integer
    Dim slNameCode As String
    Dim slCode As String

    If igExportType <= 1 Then
        tmcClick_Timer
        If Index = 0 Then
            ckcInclMissed.Visible = False           '03-24-15  disallow missed option for std bdcst, past from rvf & future from contracts
            ckcInclMissed.Value = vbUnchecked
            ckcInclCmmts.Visible = False: ckcInclCmmts.Value = vbUnchecked
        ElseIf Index = 1 Then                       '6-19-20 cal spots allow missed to be included
            ckcInclMissed.Visible = True
            ckcInclCmmts.Visible = True             'TTP 10666
        ElseIf Index = 2 Then                       'cal by contract
            ckcInclMissed.Visible = False           '6-19-20 cal contract doesnt allow missed to be included
            ckcInclMissed.Value = False
            ckcInclCmmts.Visible = False: ckcInclCmmts.Value = vbUnchecked
        Else 'by Bill Cycle
            ckcInclMissed.Visible = True
            ckcInclCmmts.Visible = False: ckcInclCmmts.Value = vbUnchecked
        End If
    End If
    
    tmcClick.Enabled = True
    imSetAll = True
    lbcVehicle.Clear
    
    If rbcMonthBy(0).Value Then             'std, allow vehicle types:  conv, rep, ntr , selling, & game
        ilRet = gPopUserVehicleBox(ExptGen, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSPORT + VEHSELLING + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + VEHSPORT + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    Else                                    'for Cal, spots exists only for conv & selling , game vehicles
        ilRet = gPopUserVehicleBox(ExptGen, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSPORT + VEHSELLING + VEHNTR + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    End If
    imVehCount = 0
    For ilLoop = 0 To ExpMatrix!lbcVehicle.ListCount - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        
        '7-9-15 check if tableau export and should be included in vehicle list
        tmSrchVffKey.iCode = ilVefCode
        ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, ilVefCode, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            '11-7-18 change test so its the same as v70.  if nothing set to Y, then all vehicles get checked on
             ilVefInx = gBinarySearchVef(tmVff.iVefCode)
             If ilVefInx > 0 Then           '6-20-20 RAB-test for inclusion
                If (imExportOption = EXP_MATRIX And tmVff.sExportMatrix = "Y") _
                   Or (imExportOption = EXP_TABLEAU And tmVff.sExportTableau = "Y") _
                   Or (imExportOption = EXP_RAB And tgMVef(ilVefInx).sExportRAB = "Y") _
                   Or (imExportOption = EXP_CUST_REV And tmVff.sExportCustom = "Y") Then 'TTP 9992 Custom Rev export, check Veh Options for ExportCustom ="Y"
                        lbcVehicle.Selected(ilLoop) = True
                        imVehCount = imVehCount + 1
                End If
            Else
                ilRet = ilRet
            End If
        End If
    Next ilLoop
    
    'TTP 10163 - hide pacing, unless on Custom Rev...availible for BC Cal only..
    If imExportOption = EXP_CUST_REV And rbcMonthBy(0).Value = True Then
        lacPacing.Visible = True
        edcPacing.Visible = True
        'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
        ckcPacingRange.Visible = True
    Else
        lacPacing.Visible = False
        edcPacing.Visible = False
        edcPacing.Text = ""
        'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
        ckcPacingRange.Value = vbUnchecked
        ckcPacingRange.Visible = False
        edcPacingEnd.Text = ""
    End If
        
    If lbcVehicle.SelCount <= 0 Then
        imVehCount = lbcVehicle.ListCount
        ckcAll.Value = vbUnchecked
        ckcAll.Value = vbChecked
    End If
End Sub

Private Sub rbcNetBy_Click(Index As Integer)
    If igExportType <= 1 Then
        tmcClick_Timer
    End If
    Exit Sub
End Sub

Private Sub tmcClick_Timer()
    Dim slRepeat As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slMonthBy As String
    Dim slMonthHdr As String
    Dim slStr As String
    Dim ilSaveMonth As Integer
    Dim slFNMonth As String
    Dim ilYear As Integer
    Dim slExtension As String * 4
    smExportFilename = ""
    tmcClick.Enabled = False
    'Determine name of export (.txt file)
    'for Matrix and Tableau - .txt; for RAB - .csv
    slExtension = ".txt"
    slRepeat = "A"
    If imExportOption = EXP_RAB Then
        slExtension = ".csv"
        ckcInclMissed.Visible = False
        slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
        slStr = ExpMatrix!edcMonth.Text             'month in text form (jan..dec, or 1-12
        gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
        If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
            ilSaveMonth = Val(slStr)
            ilRet = gVerifyInt(slStr, 1, 12)
            If ilRet = -1 Then
                Exit Sub
            End If
        End If
    
        DoEvents
        slFNMonth = Mid$(slMonthHdr, (ilSaveMonth - 1) * 3 + 1, 3)          'get the text month (jan...dec)
        slStr = ExpMatrix!edcYear.Text
        ilYear = gVerifyYear(slStr)
        If ilYear = 0 Then
            Exit Sub
        End If
        '6-17-20 alter the filename
        If rbcMonthBy(0).Value Then  'Std Bcast cal
            slMonthBy = "Std-"
            ckcInclMissed.Visible = False           '03-24-15  disallow missed option for std bdcst, past from rvf & future from contracts
            ckcInclMissed.Value = vbUnchecked
        ElseIf rbcMonthBy(1).Value Then  'cal by Spots
            slMonthBy = "CalSpots-"
            ckcInclMissed.Visible = True
        ElseIf rbcMonthBy(2).Value Then  'cal by contract
            slMonthBy = "CalCnt-"
            ckcInclMissed.Visible = False
            ckcInclMissed.Value = False
        Else                             'by Bill Method
            slMonthBy = "BillMethod-"
            ckcInclMissed.Visible = True
            ckcInclMissed.Value = False
        End If
        slMonthBy = slMonthBy & Trim$(slFNMonth) & Trim$(str$(ilYear)) & "-"
        '11/3/20 - TTP # 9993 - Add number of months after the start month
        slMonthBy = slMonthBy & Trim$(edcNoMonths.Text) & "-"
        
    ElseIf imExportOption = EXP_CUST_REV Then
        'TTP 9992 - Custom Revenue Export
        slExtension = ".csv"
        ckcInclMissed.Visible = False
        slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
        slStr = ExpMatrix!edcMonth.Text             'month in text form (jan..dec, or 1-12
        gGetMonthNoFromString slStr, ilSaveMonth    'getmonth #
        If ilSaveMonth = 0 Then                     'input isn't text month name, try month #
            ilSaveMonth = Val(slStr)
            ilRet = gVerifyInt(slStr, 1, 12)
            If ilRet = -1 Then
                Exit Sub
            End If
        End If
    
        DoEvents
        slFNMonth = Mid$(slMonthHdr, (ilSaveMonth - 1) * 3 + 1, 3)          'get the text month (jan...dec)
        slStr = ExpMatrix!edcYear.Text
        ilYear = gVerifyYear(slStr)
        If ilYear = 0 Then
            Exit Sub
        End If
        
        If rbcMonthBy(0).Value Then
            slMonthBy = "Std-"
            ckcInclMissed.Visible = False           '03-24-15  disallow missed option for std bdcst, past from rvf & future from contracts
            ckcInclMissed.Value = vbUnchecked
        ElseIf rbcMonthBy(1).Value Then
            slMonthBy = "CalSpots-"
            ckcInclMissed.Visible = True
        Else                'cal by contract
            slMonthBy = "CalCnt-"
            ckcInclMissed.Visible = False
            ckcInclMissed.Value = False
        End If
        slMonthBy = slMonthBy & Trim$(slFNMonth) & Trim$(str$(ilYear)) & "-"
        slMonthBy = slMonthBy & Trim$(edcNoMonths.Text) & "-"

    Else
        If rbcMonthBy(0).Value = True Then
            slMonthBy = "Brd"
            ckcInclMissed.Visible = False           '03-24-15  disallow missed option for std bdcst, past from rvf & future from contracts
            ckcInclMissed.Value = vbUnchecked
        Else
            slMonthBy = "Cal"                       '03-24-15 allow missed/cancel to be included or excluded
            ckcInclMissed.Visible = True
        End If
        If rbcNetBy(0).Value = True Then
            slMonthBy = Trim$(slMonthBy) & "-Net"
        Else
            slMonthBy = Trim$(slMonthBy) & "-TNet"
        End If
    End If
    
    Do
        ilRet = 0
        smExportFilename = Trim$(smExportOptionName) & " " & slMonthBy & Format(gNow, "mmddyy") & slRepeat & " " & gFileNameFilter(Trim$(smClientName)) & slExtension
        smExportName = Trim$(sgExportPath) & Trim$(smExportOptionName) & " " & slMonthBy & Format(gNow, "mmddyy")
        smExportName = Trim$(smExportName) & slRepeat & " " & gFileNameFilter(Trim$(smClientName))
        'TTP 10596 - Custom Revenue Export: add capability of running pacing version for a date range
        If ckcPacingRange.Value = vbChecked Then
            lmPacingDate = gDateValue(Trim(edcPacing.Text))
            lmPacingDate = lmPacingDate + imPacingDay 'TTP 10596 - : Custom Revenue Export: add capability of running pacing version for a date range
            smExportName = Trim$(smExportName) & "-Pacing_" & Format(lmPacingDate, "YYYYMMDD")
        End If
        smExportName = Trim$(smExportName) & slExtension            '2-27-14
        'slDateTime = FileDateTime(smExportName)
        ilRet = gFileExist(smExportName)
        If ilRet = 0 Then       'if went to mOpenErr , there was a filename that existed with same name. Increment the letter
            slRepeat = Chr(Asc(slRepeat) + 1)
            If slRepeat = "[" Then MsgBox "Cleanup Exports folder.  Too many duplicate Exports found.", vbCritical, "Exports"
        End If
    Loop While ilRet = 0
    edcTo.Text = smExportName
    edcTo.Visible = True
    Exit Sub
End Sub

'
'               Search the array of vehicle groups (tmMnfGroups)
'               <input> ilMnfCode = Multiname code
'               Return : -1 if not found, else index to the vehicle group item
Private Function mBinarySearchMnfVehicleGroup(ilMnfCode As Integer)
    Dim ilMiddle As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    ilMin = LBound(tmMnfGroups)
    ilMax = UBound(tmMnfGroups) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilMnfCode = tmMnfGroups(ilMiddle).iCode Then
            'found the match
            mBinarySearchMnfVehicleGroup = ilMiddle
            Exit Function
        ElseIf ilMnfCode < tmMnfGroups(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchMnfVehicleGroup = -1
End Function

'               get the vehicle group description whether
'               it be format, subcompany, subtotals, research or market
'               <input> ilMnfCode - VEhicle group code
'               <output>  slName - description if one exists
Private Sub mGetVGName(tlVef As VEF, ilMnfCode As Integer, slName As String)
    Dim ilVGIndex As Integer
    slName = ""         'init if one not found
    'TTP 10586 - Custom Revenue Export: vehicle group values may not be shown even when defined
    'If tlVef.iMnfVehGp2 > 0 Then
    If ilMnfCode > 0 Then
        ilVGIndex = mBinarySearchMnfVehicleGroup(ilMnfCode)
        If ilVGIndex >= 0 Then
            slName = Trim$(tmMnfGroups(ilVGIndex).sName)
        End If
    End If
End Sub

Private Sub mObtainSlsRevenueShare(llGross As Long, llNet As Long, llAcquisition As Long, ilLoopOnSlsp As Integer, tlMatrixInfo As MATRIXINFO, ilMonthInx As Integer, ilReverseSign As Integer)
    Dim slStr As String
    Dim slSharePct As String
    Dim slGrossAmount As String
    Dim slNetAmount As String
    Dim llSplitNetAmt As Long
    Dim llSplitGrossAmt As Long
    Dim llSplitAcquisitionAmt As Long
    Dim slAcqAmount As String
    
    If lmSlfSplit(ilLoopOnSlsp) = 0 Then
        tlMatrixInfo.lGross(ilMonthInx) = 0
        tlMatrixInfo.lNet(ilMonthInx) = 0
        tlMatrixInfo.lAcquisition(ilMonthInx) = 0
        Exit Sub
    End If
    lmTempPct = lmTempPct - lmSlfSplit(ilLoopOnSlsp)            'orignally the total % of all slsp (could exceed 100%)
                                                                'as each slsp is processed, that % is subt from original.
                                                                'the last gets the extra pennies
    slSharePct = gLongToStrDec(lmSlfSplit(ilLoopOnSlsp), 4)       'slsp share

    If lmTempPct <= 0 Then           '100 exhausted, last slsp gets extra pennies
        slSharePct = "100.0000"
        llSplitNetAmt = lmTempNet           'remainder of $ left to split, last one gets extra pennies
        llSplitGrossAmt = lmTempGross
        llSplitAcquisitionAmt = lmTempAcquisition
    Else
        slGrossAmount = gLongToStrDec(llGross, 2)
        slStr = gMulStr(slSharePct, slGrossAmount)                       ' gross portion of possible split
        llSplitGrossAmt = Val(gRoundStr(slStr, "1", 0))
        lmTempGross = lmTempGross - llSplitGrossAmt
        
        slNetAmount = gLongToStrDec(llNet, 2)
        slStr = gMulStr(slSharePct, slNetAmount)                       ' net portion of possible split
        llSplitNetAmt = Val(gRoundStr(slStr, "1", 0))
        lmTempNet = lmTempNet - llSplitNetAmt
        
        slAcqAmount = gLongToStrDec(llAcquisition, 2)
        slStr = gMulStr(slSharePct, slAcqAmount)                       ' acquisition portion of possible split
        llSplitAcquisitionAmt = Val(gRoundStr(slStr, "1", 0))
        lmTempAcquisition = lmTempAcquisition - llSplitAcquisitionAmt
    End If
    
    If ilReverseSign Then
        tlMatrixInfo.lGross(ilMonthInx) = -llSplitGrossAmt         '9-23-11 put $ in the month they belong in
        tlMatrixInfo.lNet(ilMonthInx) = -llSplitNetAmt
        tlMatrixInfo.lAcquisition(ilMonthInx) = -llSplitAcquisitionAmt
    Else
        tlMatrixInfo.lGross(ilMonthInx) = llSplitGrossAmt           '9-23-11 put $ in the month they belong in
        tlMatrixInfo.lNet(ilMonthInx) = llSplitNetAmt
        tlMatrixInfo.lAcquisition(ilMonthInx) = llSplitAcquisitionAmt
    End If
'Debug.Print " - mObtainSlsRevenueShare, " & ilMonthInx & "=" & llSplitGrossAmt
End Sub

'
'           Create Calendar Month export with spot file
'
Public Sub mCrCalendarMatrix(llStartDates() As Long)
    Dim ilRet As Integer
    Dim slCntrTypes As String
    Dim slCntrStatus As String
    Dim ilHOState As Integer
    Dim slStart As String
    Dim slEnd As String
    Dim llContrCode As Long
    Dim ilCurrentRecd As Integer
    Dim llLoopOnSpots As Long
    Dim blFirstTime  As Boolean
    Dim slPrevKey As String
    Dim slCurrKey As String
    Dim ilCurrLine As Integer
    Dim ilPrevLine As Integer
    Dim ilCurrVefCode As Integer
    Dim ilPrevVefCode As Integer
    Dim slTempKey As String
    Dim slTempLine As String
    Dim slTempVehicle As String
    Dim ilMnfSubCo As Integer
    Dim ilLoop As Integer
    Dim ilVefCode As Integer
    Dim llSpotInx As Long
    Dim llAirDate As Long
    Dim ilLoopOnMonth As Integer
    Dim ilFoundMonth As Integer
    Dim ilClf As Integer
    Dim slCashAgyComm  As String
    Dim tlMatrixInfo As MATRIXINFO
    Dim tlSBFTypes As SBFTypes
    Dim blValidVehicle As Boolean
    Dim blLineExists As Boolean     '6-23-15  prevent wrong line association if ssf errors and no clf for sdf
    'TTP 10666 - Podcast Ad Server for RAB Cal Spots
    Dim lDIGITALLINEAVERAGE As DIGITALLINEAVERAGE
    Dim slRvfStart As String
    Dim slRvfEnd As String
    Dim tmTranTypes As TRANTYPES
    tmTranTypes.iInv = True
    tmTranTypes.iWriteOff = False
    tmTranTypes.iPymt = False
    tmTranTypes.iCash = True
    tmTranTypes.iTrade = False
    tmTranTypes.iMerch = False
    tmTranTypes.iPromo = False
    tmTranTypes.iNTR = False
    tmTranTypes.iAirTime = True
    'TTP 10844 - RAB and Billed and Booked report: digital line invoice adjustments were being included even if the "include adjustments" checkbox was not checked
    'tmTranTypes.iAdj = True
    tmTranTypes.iAdj = False
    ReDim tlRvf(0 To 0) As RVF
    Dim tlPcf As PCF
    Dim blFound As Boolean
    Dim ilAdjust As Integer
    Dim llExtAmount As Long
    Dim llCPMStartDate As Long
    Dim llCPMEndDate As Long
    Dim slTemp As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim ilLoop2 As Integer
    Dim tmPcf2() As PCF         'All PCF records for all versions of the contract TTP 10955
    
    'setup type statement as to which type of SBF records to retrieve (only NTR)
    tlSBFTypes.iNTR = True          'include NTR billing
    tlSBFTypes.iInstallment = False      'exclude Installment billing
    tlSBFTypes.iImport = False           'exclude rep import billing

    slStart = Format$(llStartDates(1) - 30, "m/d/yy")
    slEnd = Format$(llStartDates(igPeriods + 1) - 1, "m/d/yy")           'end of last calendar month
    If lmCntrNo > 0 Then
        ReDim tmChfAdvtExt(0 To 1) As CHFADVTEXT
        tmChfSrchKey1.lCntrNo = lmCntrNo
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, Len(tmChf), tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
           ReDim tmChfAdvtExt(0 To 0) As CHFADVTEXT
        Else
            'setup 1 entry in the active contract array for processing single contract
            tmChfAdvtExt(0).lCntrNo = tmChf.lCntrNo
            tmChfAdvtExt(0).lCode = tmChf.lCode
            tmChfAdvtExt(0).iSlfCode(0) = tmChf.iSlfCode(0)
            tmChfAdvtExt(0).iAdfCode = tmChf.iAdfCode
            'JW 04/19/23 - found Bill Cycle wasnt populated when using Debug Contract #
            tmChfAdvtExt(0).sBillCycle = tmChf.sBillCycle
        End If
    Else
        'Gather all contracts for previous year and current year whose effective date entered
        'is prior to the effective date that affects either previous year or current year
        'slCntrTypes = gBuildCntTypes() '3-6-19 remove testing user for allowable contract typesd
        slCntrTypes = "CTRQ"         'all types: regular contracts (C), PI, DR, Remnants,reservations (V:  ignore since they are not invoiced) Ignore PSA(p) and Promo(m)
        slCntrStatus = "HO"               'Sched Holds, orders
        
        ilHOState = 1                       'Sched holds or orders only, no unsch contracts since spots need to be retrieved
        ilRet = gObtainCntrForDate(ExpMatrix, slStart, slEnd, slCntrStatus, slCntrTypes, ilHOState, tmChfAdvtExt())
    End If
    
    'readjust the start date to only pick up spots from the user requested period.  Backing it up to find the contracts was necessary in case
    'a mg/outside spot was sched after the contracts expiration date
    slStart = Format$(llStartDates(1), "m/d/yy")
    For ilCurrentRecd = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1     'loop on contracts
        
        'obtain the contract & lines and save the common header info
        llContrCode = tmChfAdvtExt(ilCurrentRecd).lCode
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())

        'Get RVF Date range - TTP 10666
        gUnpackDate tgChfCT.iStartDate(0), tgChfCT.iStartDate(1), slRvfStart
        If tmChfAdvtExt(ilCurrentRecd).sBillCycle = "C" Then 'Cal
            'v81 TTP 10666 - new issue Wed 5/3/23 4:01 PM
            slRvfEnd = Format(tmBillCycle.lCalBillCycleLastBilled, "ddddd")
        Else
            slRvfEnd = Format(tmBillCycle.lStdBillCycleLastBilled, "ddddd")
        End If

        'initialize Matrix info common to contract
        tlMatrixInfo.iAgfCode = tgChfCT.iAgfCode
        tlMatrixInfo.iAdfCode = tgChfCT.iAdfCode
        tlMatrixInfo.sProduct = tgChfCT.sProduct
        tlMatrixInfo.iMnfComp1 = tgChfCT.iMnfComp(0)
        tlMatrixInfo.iMnfComp2 = tgChfCT.iMnfComp(1)
        tlMatrixInfo.sOrderType = tgChfCT.sType
        tlMatrixInfo.lCntrNo = tgChfCT.lCntrNo
        tlMatrixInfo.lExtCntrNo = tgChfCT.lExtCntrNo   'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
        gUnpackDateLong tgChfCT.iOHDDate(0), tgChfCT.iOHDDate(1), tlMatrixInfo.lOHDPacingDate           '6-17-20 added for RAB
        tlMatrixInfo.iNTRType = 0
        
        For ilLoop = 1 To 24                '1-22-12
            tlMatrixInfo.lDirect(ilLoop) = 0
            lmProject(ilLoop) = 0
            lmAcquisition(ilLoop) = 0
        Next ilLoop
        'obtain agency for commission
        If tgChfCT.iAgfCode > 0 Then
            slCashAgyComm = ".00"
            tmAgfSrchKey.iCode = tgChfCT.iAgfCode
            ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)
            End If
        Else
            slCashAgyComm = ".00"
        End If              'iagfcode > 0
        
        '-----------------------------------------------------------
        'Get NTR
        If tgChfCT.sNTRDefined = "Y" And ExpMatrix!ckcNTR.Value = vbChecked Then        'this has NTR billing
            mMatrixNTR tlSBFTypes, llStartDates(), tlMatrixInfo, 1, slCashAgyComm
        End If
        
        '-----------------------------------------------------------
        'Get Spots
        tlMatrixInfo.iNTRType = 0                   '1-28-20 init ntr description for air time data
        'search for all spots by for this contract
        ReDim tmSdfExtSort(0 To 0) As SDFEXTSORT
        'ReDim tmSdfExt(1 To 1) As SDFEXT
        ReDim tmSdfExt(0 To 0) As SDFEXT
        llContrCode = tmChfAdvtExt(ilCurrentRecd).lCode
        ilRet = gObtainCntrSpot(-1, False, llContrCode, -1, "S", slStart, slEnd, tmSdfExtSort(), tmSdfExt(), 0, False, True) 'search for spots between requested user dates
        blFirstTime = True
        'loop thru the sorted array to get spots by line, vehicle, date
        For llLoopOnSpots = LBound(tmSdfExtSort) To UBound(tmSdfExtSort) - 1
            llSpotInx = tmSdfExtSort(llLoopOnSpots).lSdfExtIndex
            'process only scheduled, makegood and outsides spots (ignore missed, cancelled hidden)
            blValidVehicle = True
            If Not gFilterLists(tmSdfExt(llSpotInx).iVefCode, imIncludeCodes, imUseCodes()) Then      'filter vehicle if selected
                blValidVehicle = False
            End If

'6-23-15 add option to check if sch status = "M" (missed) or "C" (cancelled) and user requests to Include Missed
            If ((tmSdfExt(llSpotInx).sSchStatus = "S" Or tmSdfExt(llSpotInx).sSchStatus = "G" Or tmSdfExt(llSpotInx).sSchStatus = "O") Or ((tmSdfExt(llSpotInx).sSchStatus = "M" Or tmSdfExt(llSpotInx).sSchStatus = "C") And (ckcInclMissed.Value = vbChecked))) And (blValidVehicle) Then
                If blFirstTime Then
                    slTempKey = Trim$(tmSdfExtSort(llLoopOnSpots).sKey)
                    ilRet = gParseItem(slTempKey, 1, "|", slTempLine)
                    ilPrevLine = Val(slTempLine)
                    ilRet = gParseItem(slTempKey, 2, "|", slTempVehicle)
                    slPrevKey = Trim$(slTempLine) & "|" & Trim$(slTempVehicle)
                    slCurrKey = slPrevKey
                    ilCurrLine = ilPrevLine

                    ilCurrVefCode = tmSdfExt(llSpotInx).iVefCode
                    ilPrevVefCode = ilCurrVefCode
                    
                    blFirstTime = False
                    'get the line index
                    blLineExists = False
                    For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                        If ilCurrLine = tgClfCT(ilClf).ClfRec.iLine Then
                            blLineExists = True
                            tmClf = tgClfCT(ilClf).ClfRec
                            Exit For
                        End If
                    Next ilClf
                Else
                    slTempKey = Trim$(tmSdfExtSort(llLoopOnSpots).sKey)
                    ilRet = gParseItem(slTempKey, 1, "|", slTempLine)
                    ilCurrLine = Val(slTempLine)
                    ilRet = gParseItem(slTempKey, 2, "|", slTempVehicle)
                    slCurrKey = Trim$(slTempLine) & "|" & Trim$(slTempVehicle)
                    ilCurrVefCode = tmSdfExt(llSpotInx).iVefCode
                End If

                If (StrComp(slPrevKey, slCurrKey, vbBinaryCompare) = 0) Then
                    If blLineExists Then
                        'equal line & vehicle
                        'determine the month this spot goes into, and accumulate the $
                        mGetRateAndAddToArray tmSdfExtSort(llLoopOnSpots).lSdfExtIndex, llStartDates()
                    Else
                        gUnpackDateLong tmSdfExt(llSpotInx).iDate(0), tmSdfExt(llSpotInx).iDate(1), llAirDate
                        'Print #hmMsg, "** Run SSFCheck for " + slTempVehicle + " for Spot ID" + str$(tmSdfExt(llSpotInx).lCode) + " on " + Format$(llAirDate, "m/d/yy")
                        gAutomationAlertAndLogHandler "** Run SSFCheck for " + slTempVehicle + " for Spot ID" + str$(tmSdfExt(llSpotInx).lCode) + " on " + Format$(llAirDate, "m/d/yy")
                    End If
                Else
                    
                    'tlMatrixInfo.iLineNo = tmSdfExt(llSpotInx).iLineNo 'TTP 10743 - RAB export: add line numbers
                    tlMatrixInfo.iLineNo = 0 '5/26/23 - Per Jason suppress the line number on the RAB export for the air time and NTR records
                    
                    'different line or vehicle, create an output line
                    ilRet = mSplitAndCreate(llStartDates(), tlMatrixInfo, 1, slCashAgyComm, ilPrevVefCode)
                    slPrevKey = slCurrKey
                    ilPrevLine = ilCurrLine
                    ilPrevVefCode = ilCurrVefCode
                    'initialize for next line/vehicle
                    For ilLoop = 1 To 24
                        lmProject(ilLoop) = 0
                        lmAcquisition(ilLoop) = 0
                    Next ilLoop
                    blLineExists = False                'compensate for ssf errors and no line exists
                    For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                        If ilCurrLine = tgClfCT(ilClf).ClfRec.iLine Then
                            blLineExists = True
                            tmClf = tgClfCT(ilClf).ClfRec
                            mGetRateAndAddToArray tmSdfExtSort(llLoopOnSpots).lSdfExtIndex, llStartDates()
                            Exit For
                        End If
                    Next ilClf
                End If
            End If                      'endif SchStatus
        Next llLoopOnSpots
        ilRet = mSplitAndCreate(llStartDates(), tlMatrixInfo, 1, slCashAgyComm, ilPrevVefCode)

        '-----------------------------------------------------------
        'TTP 10666 - daily avg Digital lines in RAB & B&B Cal Spots
        'Get Ad Server
        tlMatrixInfo.iNTRType = 0
        ilRet = Asc(tgSaf(0).sFeatures8)
        If ((ilRet And PODCASTCPMTAG) = PODCASTCPMTAG) And (tgChfCT.sAdServerDefined = "Y") And (imExportOption = EXP_RAB Or imExportOption = EXP_MATRIX Or imExportOption = EXP_EFFICIOPROJ Or imExportOption = EXP_TABLEAU Or imExportOption = EXP_CUST_REV) Then
            ReDim tmSBFAdjust(0 To 0) As ADJUSTLIST  'build new for every contract
            'Insure the common monthly buckets are initialized for the schedule lines
            For ilLoop = 1 To 24
                lmProject(ilLoop) = 0
                lmAcquisition(ilLoop) = 0
            Next ilLoop
            '-------------------------------------
            'Get PCF (pcf_Pod_CPM_Cntr)
            ilRet = gObtainPcf(hmPcf, tgChfCT.lCode, tmPcf())       'obtain all pcm for matching contract code
            If ilRet Then
                '-------------------------------------
                'gather all the receivables that exist for this contract
'Debug.Print "gObtainPhfRvfbyCntr:" & slRvfStart & " - " & slRvfEnd
                ilRet = gObtainPhfRvfbyCntr(Me, tgChfCT.lCntrNo, slRvfStart, slRvfEnd, tmTranTypes, tlRvf())
                If ilRet Then
                    '-------------------------------------
                    'create entry into tmSBFAdjust, for all Lines ORDERED from PCF
                    For ilPcfLoop = LBound(tmPcf) To UBound(tmPcf) - 1
                        tlPcf = tmPcf(ilPcfLoop)
                        blValidVehicle = True
                        If Not gFilterLists(tlPcf.iVefCode, imIncludeCodes, imUseCodes()) Then
                            blValidVehicle = False        'not a selected vehicle; bypass
                        End If
                        If blValidVehicle And tlPcf.sType <> "P" Then            'get only the hidden lines or standard line IDs, ignore Pkg
                            blFound = False
                            For ilLoop = LBound(tmSBFAdjust) To UBound(tmSBFAdjust) - 1 Step 1
                                If tmSBFAdjust(ilLoop).lPodCode = tlPcf.lCode Then
                                    tmSBFAdjust(ilLoop).lOrderedCPMCost = tmSBFAdjust(ilLoop).lOrderedCPMCost + tlPcf.lTotalCost
                                    blFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not (blFound) Then
                                ilLoop = UBound(tmSBFAdjust)
                                tmSBFAdjust(ilLoop).iVefCode = tlPcf.iVefCode
                                tmSBFAdjust(ilLoop).lOrderedCPMCost = tlPcf.lTotalCost
                                tmSBFAdjust(ilLoop).lPodCode = tlPcf.lCode
                                'JW - 5/18/23 - Fixed RAB and B&B for V81 TTP 10725 – new issue 5-17-23.zip
                                tmSBFAdjust(ilLoop).iPodCPMID = tlPcf.iPodCPMID
                                ReDim Preserve tmSBFAdjust(0 To UBound(tmSBFAdjust) + 1) As ADJUSTLIST
                            End If
                        End If
                    Next ilPcfLoop

                    '-------------------------------------
                    'get the amounts of adserver Ordered (PCF)
                    'TTP 10955 updated gBuildDigitalLineAverage
                    gObtainPcfByCntrNo tgChfCT.lCntrNo, tmPcf2
                    For ilPcfLoop = LBound(tmPcf) To UBound(tmPcf) - 1
                        tlPcf = tmPcf(ilPcfLoop)
                        
                        lDIGITALLINEAVERAGE = gBuildDigitalLineAverage(tmChfAdvtExt(ilCurrentRecd).sBillCycle, tmBillCycle.lStdBillCycleLastBilled, tmBillCycle.lCalBillCycleLastBilled, tlPcf, tlRvf(), tmPcf2)
                        For ilAdjust = LBound(tmSBFAdjust) To UBound(tmSBFAdjust) - 1
                            If tmSBFAdjust(ilAdjust).lPodCode = lDIGITALLINEAVERAGE.lPcfCode Then
                                'Match the period Amounts from lDIGITALLINEAVERAGE into the reported periods
                                For ilLoop = 1 To 24
                                    tlMatrixInfo.dCurrMoAvg(ilLoop) = 0
                                    tlMatrixInfo.dNextMoAvg(ilLoop) = 0
                        
                                    slTemp = gObtainEndStd(Format(tmBillCycle.lCalBillCycleStartDates(ilLoop), "ddddd"))
                                    gObtainYearMonthDayStr slTemp, True, slYear, slMonth, slDay
                                    For ilLoop2 = 0 To lDIGITALLINEAVERAGE.iLastInx
                                        If lDIGITALLINEAVERAGE.iMonths(ilLoop2) = Val(slMonth) And lDIGITALLINEAVERAGE.iYears(ilLoop2) = Val(slYear) Then
                                            llExtAmount = CLng(lDIGITALLINEAVERAGE.dExtAmountGross(ilLoop2) * 100)
                                            tlMatrixInfo.sComment(ilLoop) = lDIGITALLINEAVERAGE.sComment(ilLoop2) 'TTP 10666
                                            'TTP 10725 - Billed and Booked Cal Spots and RAB Cal Spots: not including digital line contract that should be included
                                            tlMatrixInfo.lGross(ilLoop) = llExtAmount
                                            
                                            'TTP 10743 - RAB export: add line numbers
                                            tlMatrixInfo.iLineNo = lDIGITALLINEAVERAGE.iLineNo
                                            
                                            'TTP 10742 - RAB Cal Spots manual export: when "include digital avg comments" is checked on, show current month and next month averages in separate comment columns to assist troubleshooting
                                            'TTP 10822 - include Current Month if 1st Month is Index 0
                                            'If ilLoop2 > 0 Then
                                            If ilLoop2 > 0 Or lDIGITALLINEAVERAGE.iFirstInx = 0 Then
                                                tlMatrixInfo.dCurrMoAvg(ilLoop) = lDIGITALLINEAVERAGE.dDailyAmt(ilLoop2)
                                            End If
                                            If tmChfAdvtExt(ilCurrentRecd).sBillCycle = "S" And ilLoop < 24 Then
                                                tlMatrixInfo.dNextMoAvg(ilLoop) = lDIGITALLINEAVERAGE.dDailyAmt(ilLoop2 + 1)
                                            End If
                                            tmSBFAdjust(ilAdjust).lProject(ilLoop) = llExtAmount
                                            lmProject(ilLoop) = lmProject(ilLoop) + llExtAmount
                                            Exit For
                                        End If
                                    Next ilLoop2
                                Next ilLoop
                            End If
                        Next ilAdjust
                        
                        '-------------------------------------
                        'Apply Amounts and Write line
                        ilRet = mSplitAndCreate(tmBillCycle.lCalBillCycleStartDates, tlMatrixInfo, 1, slCashAgyComm, tlPcf.iVefCode)
                        '-------------------------------------
                        'init the projected gross & net values
                        For ilLoop = 1 To 24
                            lmProject(ilLoop) = 0
                            lmAcquisition(ilLoop) = 0
                            tlMatrixInfo.dCurrMoAvg(ilLoop) = 0
                            tlMatrixInfo.dNextMoAvg(ilLoop) = 0
                        Next ilLoop
                        tlMatrixInfo.iLineNo = 0
                    Next ilPcfLoop
                End If
            End If
        End If
    Next ilCurrentRecd 'End of Contract Loop
    
    Erase tmSdfExtSort, tmSdfExt
    Erase tmChfAdvtExt
    Exit Sub
End Sub

'                   mSplitAndCreate - obtain all $ obtained from spots, create export records for each
'                   split salesperson
'
'                   Contract header and lines are in memory
Private Function mSplitAndCreate(llStartDates() As Long, tlMatrixInfo As MATRIXINFO, ilFirstProjInx As Integer, slCashAgyComm As String, ilVefCode As Integer) As Integer
    Dim ilMnfSubCo As Integer
    Dim ilCorT As Integer
    Dim ilStartCorT As Integer
    Dim ilEndCorT As Integer
    Dim ilTemp As Integer
    Dim llTempGross(0 To 24) As Long    'index zero ignored
    Dim llTempNet(0 To 24) As Long      'index zero ignored
    Dim llTempAcquisition(0 To 24) As Long  'index zero ignored
    Dim slPctTrade As String
    Dim ilLoopOnMonth As Integer
    Dim ilLoopOnSlsp As Integer
    Dim ilReverseSign As Integer
    Dim slStr As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim ilRet As Integer
    Dim ilError As Integer
    Dim blGotRevenue As Boolean
    
    ilError = False                        'error return from function
    slPctTrade = gIntToStrDec(tgChfCT.iPctTrade, 0)
    
    If tgChfCT.iPctTrade = 0 Then                     'setup loop to do cash & trade
        ilStartCorT = 1
        ilEndCorT = 1
    ElseIf tgChfCT.iPctTrade = 100 Then
        ilStartCorT = 1
        ilEndCorT = 1
    Else
        ilStartCorT = 1     'split cash/trade
        ilEndCorT = 2
    End If

    'create an output line
    ReDim lmSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
    ReDim imSlfCode(0 To 9) As Integer             '4-20-00
    ReDim imslfcomm(0 To 9) As Integer             'slsp under comm %
    ReDim imslfremnant(0 To 9) As Integer          'slsp under remnant %
    ReDim lmSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)

    ilMnfSubCo = gGetSubCmpy(tgChfCT, imSlfCode(), lmSlfSplit(), ilVefCode, False, lmSlfSplitRev())

    For ilCorT = ilStartCorT To ilEndCorT Step 1        '2 passes if split cash/trade
        'move the projected values into temporary work area, as they might be modified
        'due to split cash/trade calculations
        blGotRevenue = False
        For ilTemp = 1 To 24
            llTempGross(ilTemp) = lmProject(ilTemp)
            llTempAcquisition(ilTemp) = lmAcquisition(ilTemp)
            If lmProject(ilTemp) <> 0 Or lmAcquisition(ilTemp) <> 0 Then
                blGotRevenue = True
'Debug.Print " mSplitAndCreate, GotRevenue:" & lmProject(ilTemp)
            End If
        Next ilTemp
        
        If blGotRevenue Then
            gCalcMonthAmt llTempGross(), llTempNet(), llTempAcquisition(), ilFirstProjInx, igPeriods, ilCorT, slCashAgyComm, tgChfCT
            'Create the export records.  The
            'Setup month and year to store in export
            'use next months start date minus 1 to get the end date of the current month
            For ilLoopOnMonth = ilFirstProjInx To igPeriods
                If llTempNet(ilLoopOnMonth) <> 0 Or llTempAcquisition(ilLoopOnMonth) <> 0 Then           'bypass $0
                    lmTempGross = llTempGross(ilLoopOnMonth)
                    lmTempNet = llTempNet(ilLoopOnMonth)
                    lmTempAcquisition = llTempAcquisition(ilLoopOnMonth)
                    lmTempPct = 0
                    'determine amt of revenue sharing; could exceed 100%
                    For ilLoopOnSlsp = 0 To 9
                        lmTempPct = lmTempPct + lmSlfSplit(ilLoopOnSlsp)
                    Next ilLoopOnSlsp
                    ilReverseSign = False
                       
                    For ilLoopOnSlsp = 0 To 9
                        slStr = Format$(llStartDates(ilLoopOnMonth + 1) - 1, "m/d/yy")
                        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                        tlMatrixInfo.iYear(ilLoopOnMonth) = Val(slYear)
                        tlMatrixInfo.iMonth(ilLoopOnMonth) = Val(slMonth)
                        tlMatrixInfo.sAirNTR = "A"          'assume Air time
                        If ilCorT = 1 Then          'cash
                            tlMatrixInfo.sCashTrade = "C"
                        Else
                            tlMatrixInfo.sCashTrade = "T"
                        End If
                                                            
                        If ilLoopOnSlsp = 0 Then            '1-22-12 1st slsp gets total gross amt as well as split in its record
                            tlMatrixInfo.lDirect(ilLoopOnMonth) = lmTempGross         'llGross
                        End If

                        mObtainSlsRevenueShare llTempGross(ilLoopOnMonth), llTempNet(ilLoopOnMonth), llTempAcquisition(ilLoopOnMonth), ilLoopOnSlsp, tlMatrixInfo, ilLoopOnMonth, False
'                                    tlMatrixInfo.lGross(ilLoopOnMonth) = llTempGross(ilLoopOnMonth)
'                                    tlMatrixInfo.lNet(ilLoopOnMonth) = llTempNet(ilLoopOnMonth)
                        tlMatrixInfo.iSlfCode = imSlfCode(ilLoopOnSlsp)

                        tlMatrixInfo.iVefCode = ilVefCode
                        ilRet = mWriteExportRec(tlMatrixInfo)
                        If ilRet <> 0 Then   'error
                            'Print #hmMsg, "Error writing export record for contract # " & str$(tgChfCT.lCntrNo) & ", Line # " & str$(tmClf.iLine)
                            gAutomationAlertAndLogHandler "Error writing export record for contract # " & str$(tgChfCT.lCntrNo) & ", Line # " & str$(tmClf.iLine)
                            ilError = True
                            mSplitAndCreate = ilError
                            Exit Function
                        End If
                    Next ilLoopOnSlsp
                End If              'llTempNet(ilLoopOnMonth) <> 0
            Next ilLoopOnMonth
        End If                          'blGotRevenue
    Next ilCorT
    mSplitAndCreate = ilError
    Exit Function
End Function

'
'               gGetRateAndAddToArray
'               <input> llSpotInx - index from the sorted array of spots (tmSdfExtSort), that points to the actual spot data array (tmSdfExt)
'               pass the index from the sorted spot array of the spot information
'               to be able to retrieve the spot cost from the flight
'
'
Public Sub mGetRateAndAddToArray(llSpotInx As Long, llStartDates() As Long)
    Dim ilLoopOnMonth As Integer
    Dim llAirDate As Long
    Dim ilRet As Integer
    Dim slSpotRate As String
    Dim llSpotRate As Long
    Dim ilAcqCommPct As Integer
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim llAcqNet As Long
    Dim llAcqComm As Long
    Dim blAcqOK As Boolean

    gUnpackDateLong tmSdfExt(llSpotInx).iDate(0), tmSdfExt(llSpotInx).iDate(1), llAirDate
    '3-7-19 A = regular contract, T = Remnant,Q = PI, V = Reserv, R = DR
    'NOTE:  Reservation contracts are changed to type A
    If tmSdfExt(llSpotInx).sSpotType = "A" Or tmSdfExt(llSpotInx).sSpotType = "T" Or tmSdfExt(llSpotInx).sSpotType = "Q" Or tmSdfExt(llSpotInx).sSpotType = "R" Then   'include regular sched spots , PI, DR
        'make up spot info for spot price routine
        tmSdf.iAdfCode = tmSdfExt(llSpotInx).iAdfCode
        tmSdf.iDate(0) = tmSdfExt(llSpotInx).iDate(0)
        tmSdf.iDate(1) = tmSdfExt(llSpotInx).iDate(1)
        tmSdf.iVefCode = tmSdfExt(llSpotInx).iVefCode
        tmSdf.lChfCode = tmSdfExt(llSpotInx).lChfCode
        tmSdf.lCode = tmSdfExt(llSpotInx).lCode
        tmSdf.sSchStatus = tmSdfExt(llSpotInx).sSchStatus
        tmSdf.sSpotType = tmSdfExt(llSpotInx).sSpotType
        tmSdf.sPriceType = tmSdfExt(llSpotInx).sPriceType
        
        For ilLoopOnMonth = 1 To igPeriods + 1 Step 1      'loop thru months to find the match
            If llAirDate >= llStartDates(ilLoopOnMonth) And llAirDate < llStartDates(ilLoopOnMonth + 1) Then
                'process the spot $
                ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slSpotRate)
                If (InStr(slSpotRate, ".") <> 0) Then        'found spot cost
                    'is it a .00?
                    If gCompNumberStr(slSpotRate, "0.00") = 0 Then       'its a .00 spot
                        llSpotRate = 0
                    Else
                        llSpotRate = gStrDecToLong(slSpotRate, 2)
                    End If
                Else
                    'its a bonus, recap, n/c, etc. which is still $0
                    llSpotRate = 0
                End If
                lmProject(ilLoopOnMonth) = lmProject(ilLoopOnMonth) + llSpotRate
                                    
                '7/31/15 implement acq commission  if applicable
                If ExpMatrix!rbcNetBy(1).Value Then     't-net?
                    If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then    'acq comm are applicable
                        ilAcqCommPct = 0
                        blAcqOK = gGetAcqCommInfoByVehicle(tmSdf.iVefCode, ilAcqLoInx, ilAcqHiInx)
                        ilAcqCommPct = gGetEffectiveAcqComm(llAirDate, ilAcqLoInx, ilAcqHiInx)
                        gCalcAcqComm ilAcqCommPct, tmClf.lAcquisitionCost, llAcqNet, llAcqComm
                        lmAcquisition(ilLoopOnMonth) = lmAcquisition(ilLoopOnMonth) + llAcqNet
                    Else                        'no acq comm; both net and gross are the same
                        lmAcquisition(ilLoopOnMonth) = lmAcquisition(ilLoopOnMonth) + tmClf.lAcquisitionCost
                    End If
                End If
                Exit For
            End If
        Next ilLoopOnMonth
    End If
    Exit Sub
End Sub

'
'                       mMatrixNTR - include NTR into Matrix export by option
'
Private Function mMatrixNTR(tlSBFTypes As SBFTypes, llStartDates() As Long, tlMatrixInfo As MATRIXINFO, ilFirstProjInx As Integer, slCashAgyComm As String) As Integer
    Dim ilRet As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim ilFoundMonth As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim ilLoopOnMonth As Integer
    Dim slPctTrade As String
    Dim ilMnfSubCo As Integer
    Dim ilCorT As Integer
    Dim ilStartCorT As Integer
    Dim ilEndCorT As Integer
    Dim llAmt As Long
    Dim llAcquisitionAmt As Long
    Dim slGross As String
    Dim slGrossPct As String
    Dim slNet As String
    Dim slAcquisition As String
    Dim llNet As Long
    Dim llGross As Long
    Dim llAcquisition As Long
    Dim ilLoopOnSlsp As Integer
    Dim ilReverseSign As Integer
    Dim slStr As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim ilError As Integer
    Dim slStart As String
    Dim slEnd As String
    Dim slCashAgyCommPct As String
    Dim blValidVehicle As Boolean
    ReDim tlSbf(0 To 0) As SBF

    ilError = False                         'error return
    
    slStart = Format$(llStartDates(1), "m/d/yy")
    slEnd = Format$(llStartDates(igPeriods + 1) - 1, "m/d/yy")
    If tgChfCT.iPctTrade = 0 Then                     'setup loop to do cash & trade
        ilStartCorT = 1
        ilEndCorT = 1
    ElseIf tgChfCT.iPctTrade = 100 Then
        ilStartCorT = 2
        ilEndCorT = 2
    Else
        ilStartCorT = 1     'split cash/trade
        ilEndCorT = 2
    End If
     ilRet = gObtainSBF(ExpMatrix, hmSbf, tgChfCT.lCode, slStart, slEnd, tlSBFTypes, tlSbf(), 0)   '11-28-06 add last parm to indicate which key to use

    For llSbf = LBound(tlSbf) To UBound(tlSbf) - 1
        tmSbf = tlSbf(llSbf)
        ilFoundMonth = False
        gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
        llDate = gDateValue(slDate)
        For ilLoopOnMonth = ilFirstProjInx To igPeriods Step 1       '11-5-14 (remove +1 from igPeriods ) loop thru months to find the match
            If llDate >= llStartDates(ilLoopOnMonth) And llDate < llStartDates(ilLoopOnMonth + 1) Then
                ilFoundMonth = True
                Exit For
            End If
        Next ilLoopOnMonth
        
        
        If Not gFilterLists(tmSbf.iAirVefCode, imIncludeCodes, imUseCodes()) Then
            blValidVehicle = False
            ilFoundMonth = False            'not a selected vehicle; bypass
        End If

        If ilFoundMonth Then
            If tmSbf.sAgyComm = "N" Then        'ntr comm flag overrides the contract
                slCashAgyCommPct = ".00"
            Else
                slCashAgyCommPct = slCashAgyComm        'agy comm determine from the contracts agency info
            End If

            ReDim lmSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
            ReDim imSlfCode(0 To 9) As Integer             '4-20-00
            ReDim imslfcomm(0 To 9) As Integer             'slsp under comm %
            ReDim imslfremnant(0 To 9) As Integer          'slsp under remnant %
            ReDim lmSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)

            ilMnfSubCo = gGetSubCmpy(tgChfCT, imSlfCode(), lmSlfSplit(), tmSbf.iAirVefCode, False, lmSlfSplitRev())

            For ilCorT = ilStartCorT To ilEndCorT
                If ilCorT = 1 Then              'cash
                    slPctTrade = gSubStr("100.", gIntToStrDec(tgChfCT.iPctTrade, 0))
                    tlMatrixInfo.sCashTrade = "C"
                Else            'trade portion
                    slPctTrade = gIntToStrDec(tgChfCT.iPctTrade, 0)
                    tlMatrixInfo.sCashTrade = "T"
                End If
                tlMatrixInfo.iNTRType = tmSbf.iMnfItem                 '1-28-20
                'convert the $ to gross & net strings
                llAmt = tmSbf.lGross * tmSbf.iNoItems
                slGross = gLongToStrDec(llAmt, 2)       'convert to xxxx.xx
                slGrossPct = gSubStr("100.00", slCashAgyCommPct)        'determine  % to client (normally 85%)
                slNet = gDivStr(gMulStr(slGrossPct, slGross), "100")    'net value
                
                llAcquisitionAmt = tmSbf.lAcquisitionCost * tmSbf.iNoItems
                slAcquisition = gLongToStrDec(llAcquisitionAmt, 2)      'convert to xxxx.xx
                
                'calculate the new gross & net if split cash/trade
                slNet = gDivStr(gMulStr(slNet, slPctTrade), "100")
                llNet = gStrDecToLong(slNet, 2)
                slGross = gDivStr(gMulStr(slGross, slPctTrade), "100")
                llGross = gStrDecToLong(slGross, 2)
                
                slAcquisition = gDivStr(gMulStr(slAcquisition, slPctTrade), "100")
                llAcquisition = gStrDecToLong(slAcquisition, 2)

                If llNet <> 0 Or llAcquisition <> 0 Then           'bypass $0
                    lmTempGross = llGross
                    lmTempNet = llNet
                    lmTempAcquisition = llAcquisition
                    'determine amt of revenue sharing; could exceed 100%
                    For ilLoopOnSlsp = 0 To 9
                        lmTempPct = lmTempPct + lmSlfSplit(ilLoopOnSlsp)
                    Next ilLoopOnSlsp
                    ilReverseSign = False
                    
                    For ilLoopOnSlsp = 0 To 9
                        slStr = Format$(llStartDates(ilLoopOnMonth + 1) - 1, "m/d/yy")
                        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                        tlMatrixInfo.iYear(ilLoopOnMonth) = Val(slYear)
                        tlMatrixInfo.iMonth(ilLoopOnMonth) = Val(slMonth)
                        tlMatrixInfo.sAirNTR = "N"          'NTR flag
                        If ilCorT = 1 Then          'cash
                            tlMatrixInfo.sCashTrade = "C"
                        Else
                            tlMatrixInfo.sCashTrade = "T"
                        End If
                        
                        If ilLoopOnSlsp = 0 Then            '1-22-12 1st slsp gets total gross amt as well as split in its record
                            tlMatrixInfo.lDirect(ilLoopOnMonth) = llGross
                        End If
                        
                        mObtainSlsRevenueShare llGross, llNet, llAcquisition, ilLoopOnSlsp, tlMatrixInfo, ilLoopOnMonth, False
                        tlMatrixInfo.iSlfCode = imSlfCode(ilLoopOnSlsp)
                        tlMatrixInfo.iVefCode = tmSbf.iBillVefCode
                        'tlMatrixInfo.lGross(ilLoopOnMonth) = llGross
                        'tlMatrixInfo.lNet(ilLoopOnMonth) = llNet
                        ilRet = mWriteExportRec(tlMatrixInfo)
                        If ilRet <> 0 Then   'error
                            'Print #hmMsg, "Error writing export record for NTR, contract # " & str$(tgChfCT.lCntrNo) & " Contract file"
                            gAutomationAlertAndLogHandler "Error writing export record for NTR, contract # " & str$(tgChfCT.lCntrNo) & " Contract file"
                            ilError = True
                            mMatrixNTR = ilError
                            Exit Function
                        End If
                    Next ilLoopOnSlsp
                End If
            Next ilCorT
        End If
    Next llSbf
    mMatrixNTR = ilError
    Exit Function
End Function

Private Sub tmcSetTime_Timer()
    If imExportOption = EXP_MATRIX Then
        gUpdateTaskMonitor 0, "ME"
    ElseIf imExportOption = EXP_TABLEAU Then
        gUpdateTaskMonitor 0, "TE"
    Else
        gUpdateTaskMonitor 0, "RE"              '2-3-20 RAB CRM export
    End If
End Sub

'TTP 10599 - Fix TTP 10205 & TTP 10503 / JW - 11/28/22 - slight Performance fix with ADFX/AGFX lookups (External Agency ID, External Advertiser ID, Advertiser CRMID, and Agency CRMID)
Sub mGetAdfxCodes(ilAdfID As Integer, smADFXRefId As String, lmADFXCRMID As Long)
    Dim slSql As String
    Dim myRsQuery As ADODB.Recordset
    'mGetAdfxRefID = ""
    If ilAdfID = 0 Then Exit Sub
    If ilAdfID = imLastAdfCode Then
        'mGetAdfxRefID = smLastAdfName
        smADFXRefId = smLastADFXRefId
        lmADFXCRMID = lmLastADFXCRMID
        Exit Sub
    End If
    slSql = "select adfxRefId as Code, adfxCRMID as CRMIDCode from ADFX_Advertisers where adfxCode = " & ilAdfID
    Set myRsQuery = gSQLSelectCall(slSql)
    If Not myRsQuery.EOF Then
        'mGetAdfxRefID = myRsQuery!Code
        'imLastAdfCode = ilAdfID
        'smLastAdfName = myRsQuery!Code
        imLastAdfCode = ilAdfID
        smADFXRefId = Trim(myRsQuery!code)
        lmADFXCRMID = myRsQuery!CRMIDCode
        smLastADFXRefId = smADFXRefId
        lmLastADFXCRMID = lmADFXCRMID
    End If
End Sub

'TTP 10599 - Fix TTP 10205 & TTP 10503 / JW - 11/28/22 - slight Performance fix with ADFX/AGFX lookups (External Agency ID, External Advertiser ID, Advertiser CRMID, and Agency CRMID)
Sub mGetAgfxCodes(ilAgfID As Integer, smAGFXRefId As String, lmAGFXCRMID As Long)
    Dim slSql As String
    Dim myRsQuery As ADODB.Recordset
    'mGetAgfxRefID = ""
    If ilAgfID = 0 Then Exit Sub
    If ilAgfID = imLastAgyCode Then
        smAGFXRefId = smLastAGFXRefId
        lmAGFXCRMID = lmLastAGFXCRMID
        Exit Sub
    End If
    smAGFXRefId = ""
    lmAGFXCRMID = 0
    slSql = "select agfxRefId as Code, agfxCRMID as CRMIDCode from AGFX_Agencies where agfxCode = " & ilAgfID
    Set myRsQuery = gSQLSelectCall(slSql)
    If Not myRsQuery.EOF Then
        imLastAgyCode = ilAgfID
        smAGFXRefId = Trim(myRsQuery!code)
        lmAGFXCRMID = myRsQuery!CRMIDCode
        smLastAGFXRefId = smAGFXRefId
        lmLastAGFXCRMID = lmAGFXCRMID
    End If
End Sub

'Function mGetOwnerID
'TTP 10447 - RAB Export: add VefCode, Participant Name, and participant code
'JW 4/15/21
Function mGetOwnerID(ilVehicleID As Long) As Integer
    If lmLastVehicleId = ilVehicleID Then
        mGetOwnerID = imLastOwnerId
        Exit Function
    Else
        Dim slSql As String
        Dim myRsQuery As ADODB.Recordset
        slSql = "SELECT top(1) pifMnfGroup From PIF_Participant_Info WHERE "
        slSql = slSql & "PifVefCode = " & ilVehicleID
        slSql = slSql & " AND pifSeqNo = 1  ORDER BY pifEndDate DESC"
        Set myRsQuery = gSQLSelectCall(slSql)
        If Not myRsQuery.EOF Then
            mGetOwnerID = myRsQuery!pifMnfGroup
            lmLastVehicleId = ilVehicleID
        End If
    End If
End Function

'Function mGetOwnerName
'TTP 10447 - RAB Export: add VefCode, Participant Name, and participant code
'JW 4/15/21
Function mGetOwnerName(ilOwnerID As Integer) As String
    If imLastOwnerId = ilOwnerID Then
        mGetOwnerName = smLastOwner
        Exit Function
    Else
        Dim slSql As String
        Dim myRsQuery As ADODB.Recordset
        slSql = "select mnfName from MNF_Multi_Names where mnfCode = " & ilOwnerID
        Set myRsQuery = gSQLSelectCall(slSql)
        If Not myRsQuery.EOF Then
            mGetOwnerName = Trim(myRsQuery!mnfname)
            smLastOwner = mGetOwnerName
            imLastOwnerId = ilOwnerID
        End If
    End If
End Function

Function EnableForm(trueFalse As Boolean)
    plcMonthBy.Enabled = trueFalse
    PlcNetBy.Enabled = trueFalse
    edcMonth.Enabled = trueFalse
    edcYear.Enabled = trueFalse
    edcNoMonths.Enabled = trueFalse
    ckcNTR.Enabled = trueFalse
    ckcInclMissed.Enabled = trueFalse
    ckcInclAdj.Enabled = trueFalse
    edcPacing.Enabled = trueFalse
    edcContract.Enabled = trueFalse
    edcTo.Enabled = trueFalse
    ckcAmazon.Enabled = trueFalse
    cmcExport.Enabled = trueFalse
    cmcCancel.Enabled = trueFalse
    cmcTo.Enabled = trueFalse
    ckcAll.Enabled = trueFalse
    lbcVehicle.Enabled = trueFalse
    DoEvents
End Function


'Boostr Phase 2: Billed and Booked std broadcast: update to use new daily or monthly method depending on Site setting for future periods
'mDeterminePeriodAmountByDaily
'   Inputs:
'       slPeriodStart   - Start date of the month being reported (string Date MM/DD/YYYY)
'       slPeriodEnd     - End date of the month being reported (string Date MM/DD/YYYY)
'       slLineStartDate - Start Date of the PCF Line (string Date MM/DD/YYYY)
'       slLineEndDate   - End Date of the PCF Line (string Date MM/DD/YYYY)
'       slRemainingAmount - The Remaining amount (Line Total minus whats already been Invoiced) (string Amount ####.##)
'   Output:
'       Month Amount (double precision number ####.##)
Public Function mDeterminePeriodAmountByDaily(slLastBilledDate As String, slPeriodStart As String, slPeriodEnd As String, slLineStartDate As String, slLineEndDate As String, llBilledAmount, llTotalAmount As Long) As Double
    Dim dlDailyAmount As Double 'The daily $ Amount
    Dim ilNumberOfDaysRemaining As Integer 'How many days from Invoice Start Date to Line EndDate
    Dim ilNumberOfDaysInPeriod As Integer 'How many days are in this period
    Dim dStartDate As Date 'Temp Start Date
    Dim dEndDate As Date 'Temp End
    
    'Determine how many days remain of this line (beyond what's been billed)
    If DateValue(slLastBilledDate) + 1 > DateValue(slLineStartDate) Then
        dStartDate = DateValue(slLastBilledDate) + 1
    Else
        dStartDate = DateValue(slLineStartDate)
    End If
    dEndDate = DateValue(slLineEndDate)
    ilNumberOfDaysRemaining = DateDiff("d", dStartDate, dEndDate) + 1
    If ilNumberOfDaysRemaining <= 0 Then Exit Function
    
    'Determine how many days of this Line are being invoiced
    dStartDate = IIF(DateValue(slLineStartDate) > gDateValue(slPeriodStart), gDateValue(slLineStartDate), gDateValue(slPeriodStart))
    dEndDate = IIF(DateValue(slLineEndDate) < gDateValue(slPeriodEnd), gDateValue(slLineEndDate), gDateValue(slPeriodEnd))
    ilNumberOfDaysInPeriod = DateDiff("d", dStartDate, dEndDate) + 1
    If ilNumberOfDaysInPeriod <= 0 Then Exit Function
    
    'Determine the daily amount
    dlDailyAmount = ((llTotalAmount - llBilledAmount) / 100) / ilNumberOfDaysRemaining
    
    'Determine the amount to apply to this Period (slPeriodStart - slPeriodEnd)
    mDeterminePeriodAmountByDaily = dlDailyAmount * ilNumberOfDaysInPeriod
    
    Debug.Print "mDeterminePeriodAmountByDaily: "
    Debug.Print " -> Line Dates: " & slLineStartDate & " to " & slLineEndDate
    Debug.Print " -> RemainingAmount: " & Format((llTotalAmount - llBilledAmount) / 100, "#.00")
    Debug.Print " -> NumberOfDaysRemaining: " & ilNumberOfDaysRemaining
    Debug.Print " -> Period: " & slPeriodStart & " to " & slPeriodEnd & " = " & ilNumberOfDaysInPeriod
    Debug.Print " -> DailyAmount: " & dlDailyAmount
    Debug.Print " -> Month Amount: " & mDeterminePeriodAmountByDaily
End Function

