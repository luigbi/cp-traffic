VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExpCashOrInv 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3690
   ClientLeft      =   825
   ClientTop       =   2400
   ClientWidth     =   8295
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
   ScaleHeight     =   3690
   ScaleWidth      =   8295
   Begin V81TrafficExports.CSI_Calendar CSI_CalEnd 
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Text            =   "01/10/2023"
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
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
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin V81TrafficExports.CSI_Calendar CSI_CalStart 
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Text            =   "01/10/2023"
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
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
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin VB.Frame frcAmazon 
      Height          =   1455
      Left            =   6240
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CheckBox ckcKeepLocalFile 
         Caption         =   "Keep Local File"
         Height          =   195
         Left            =   4080
         TabIndex        =   23
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox edcBucketName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         ToolTipText     =   "Bucket Name"
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox edcAccessKey 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5280
         PasswordChar    =   "*"
         TabIndex        =   21
         ToolTipText     =   "The Access Key Assigned by AWS"
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox edcPrivateKey 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5280
         PasswordChar    =   "*"
         TabIndex        =   22
         ToolTipText     =   "The Private Key Assigned by AWS"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox edcRegion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         ToolTipText     =   "Region/Endpoint - Example: USEast1, USEast2, USWest1 or USWest2"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox edcAmazonSubfolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         ToolTipText     =   "(Optional) Amazon Web Bucket Subfolder Name.   Example: Counterpoint"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lacExportFilename 
         Caption         =   "lacExportFilename"
         Height          =   255
         Left            =   6000
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "BucketName"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "AccessKey"
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "PrivateKey"
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Region"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1095
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
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.OptionButton optFormat 
      Caption         =   "Excel (XLS)"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   33
      Top             =   2280
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton optFormat 
      Caption         =   "CSV"
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   32
      Top             =   2280
      Width           =   1335
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
      Left            =   6720
      TabIndex        =   13
      Top             =   2640
      Width           =   1365
   End
   Begin VB.PictureBox plcTo 
      Height          =   375
      Left            =   1080
      ScaleHeight     =   315
      ScaleWidth      =   5445
      TabIndex        =   30
      Top             =   2640
      Width           =   5505
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   30
         Width           =   5505
      End
   End
   Begin VB.CheckBox ckcAmazon 
      Caption         =   "Upload to Amazon Web bucket"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5595
      Top             =   360
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1305
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5115
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5715
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1245
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
      Left            =   3360
      TabIndex        =   15
      Top             =   3195
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
      Left            =   4680
      TabIndex        =   16
      Top             =   3195
      Width           =   1050
   End
   Begin VB.CheckBox ckcInclTrade 
      Caption         =   "Include Trade"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1815
   End
   Begin VB.CheckBox ckcInclAdj 
      Caption         =   "Include Adjustments"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   6480
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.CheckBox ckcSummary 
      Caption         =   "Summary Version"
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame frcSummary 
      Caption         =   "Summary Format"
      Height          =   495
      Left            =   2640
      TabIndex        =   36
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
      Begin VB.OptionButton optSummaryFormat2 
         Caption         =   "Sage Intacct"
         Height          =   195
         Left            =   1680
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optSummaryFormat1 
         Caption         =   "22-Column"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label lacFileType 
      Appearance      =   0  'Flat
      Caption         =   "File Type"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   34
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label lacSaveIn 
      Appearance      =   0  'Flat
      Caption         =   "Save In"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   31
      Top             =   2670
      Width           =   810
   End
   Begin VB.Label lacSelCFrom 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   885
   End
   Begin VB.Label lacTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Export"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   75
      Width           =   555
   End
   Begin VB.Label lacSelCFrom1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2760
      TabIndex        =   6
      Top             =   480
      Width           =   810
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   420
      TabIndex        =   4
      Top             =   1650
      Visible         =   0   'False
      Width           =   7740
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   2550
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   420
      TabIndex        =   3
      Top             =   1365
      Visible         =   0   'False
      Width           =   7710
   End
End
Attribute VB_Name = "ExpCashOrInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim hmTo As Integer   'From file hanle
Dim hmCashInv As Integer    'file handle
Dim hmMsg As Integer        'error message logging file
Dim imTerminate As Integer
Dim imExporting As Integer
Dim smStart As String       'Starting date entered without slashes for filename
Dim smEnd As String         'Ending date entered without slashes for filename
Dim smFullStartDate As String   'start date entered with slashes
Dim smFullEndDate As String     'end date entered with slashes
Dim lmStart As Long         'Starting date
Dim lmEnd As Long           'Ending date

Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlfSrchKey0 As INTKEY0     'SLF key image
Dim tmSlf As SLF

Dim hmSof As Integer            'Sale Office file handle
Dim imSofRecLen As Integer      'SOF record length
Dim tmSof As SOF

Dim hmAdf As Integer 'Advertiser file handle
Dim tmAdf As ADF        'ADF record image
Dim imAdfRecLen As Integer        'ADF record length

Dim hmAgf As Integer
Dim tmAgf As AGF
Dim imAgfRecLen As Integer
Dim tmSrchKey As INTKEY0

Dim tmChf As CHF
Dim imCHFRecLen As Integer
Dim hmCHF As Integer
Dim tmChfSrchKey1 As CHFKEY1

Dim tmMnf As MNF
Dim imMnfRecLen As Integer
Dim hmMnf As Integer
Dim tmMnfSrchKey As INTKEY0

Dim tmPrf As PRF
Dim imPrfRecLen As Integer
Dim hmPrf As Integer
Dim tmPrfSrchKey As LONGKEY0

Dim tmRvf As RVF
Dim imRvfRecLen As Integer
Dim hmRvf As Integer

Dim hmVef As Integer

Dim tmExport_TranInfo As EXPORT_TRANINFO
Dim tmExport_TranSummary() As EXPORT_TRANSUMMARY
Dim tmSofList() As SOF
Dim smSSMnfStamp As String
Dim tmSSMnfList() As MNF
Dim tmNTRList() As MNF
Dim smNTRMNFStamp As String
Dim imExportOption As Integer
Dim smExportOptionName As String
Dim smExportName As String
Dim lmNowDate As Long
'TTP 10208 - overflow error on Invoice Register export and Cash Receipts export
'control totals at end of export
'Dim lmGross As Long
'Dim lmComm As Long
'Dim lmNet As Long
'control totals at end of export
Dim dmGross As Double
Dim dmComm As Double
Dim dmNet As Double

Private Type EXPORT_TRANINFO
    sAccountID As String * 10   'Agency or direct advertiser station code
    sAgyName As String * 40     'Agency name
    iAgyCode As Integer         'Agency Code
    sAdvName As String * 30     'Advertiser name
    iAdvCode As Integer         'Advertiser Code
    sProduct As String * 35     'Product name
    sSlspName As String * 40    'Salesperson first/last name
    iSlspCode As Integer         'SalesmanID - TTP 10487
    sSlspStnCode As String * 40 'Salesperson Station Code
    sOffice As String * 20      'sales office
    sSalesSource As String * 20 'sales source
    sBusCat As String * 20      'Business category
    sNTRType As String * 20     'NTR Item Name
    sBillVehicle As String * 40 'Billing Vehicle Name
    iAirVehicleCode As Integer  'Airing Vehicle Code - TTP 10487
    sAirVehicle As String * 40  'Airing vehicle name
    sContract As String * 9     'contract#
    sTranDate As String * 10    'Transaction Date
    sInvNo As String * 10       'invoice #
    sTranType As String * 2     'Transaction type
    sAction As String * 1       'Action on Payment or journal entry
    sPostDate As String * 10    'posting transaction date
    sCheck As String * 10       'check number
    sNet As String * 10         'net amount
    sGross As String * 10       'gross amount
    sComm As String * 10        'commission
    sCashTrade As String * 1    'C = Cash ,T = trade
    sPolitical As String * 1    'Y = political, else N
End Type

Private Type EXPORT_TRANSUMMARY
    iAdvCode As Integer             'Advertiser Code
    iAgyCode As Integer             'Agency Code
    sAccountID As String * 10       'agency or direct advertiser station code (not used)
    sAction As String * 1           'Action on Payment or journal entry (not used)
    sAdvertiserRefId As String * 36 'AdvertiserRefID (Guid)
    sAdvName As String * 30         'Advertiser name
    sAgencyRefId As String * 36     'AgencyRefID (Guid)
    sAgyName As String * 40         'Agency name
    iAirVehicleCode As Integer      'Airing vehicle ID - Invoice Register Export - add summary version setting to current export
    sAirVehicle As String * 40      'Airing vehicle name (not used)
    sBalance As String * 10         'Balance
    sBillingGroup As String         'BillingGroup
    sBillVehicle As String * 40     'Billing Vehicle Name (not used)
    sBusCat As String * 20          'Business category (not used)
    sCashTrade As String * 1        'C = Cash ,T = trade (not used)
    sCheck As String * 10           'check number (not used)
    sComm As String * 10            'commission (not used)
    sContract As String * 9         'contract#
    sGross As String * 10           'gross amount (not used)
    sInvNo As String * 10           'invoice #
    sNet As String * 10             'net amount (not used)
    sNTRType As String * 20         'NTR Item Name (not used)
    sOffice As String * 20          'sales office (Not used)
    sOfficeCode As String * 20      'SalesOfficeCode
    sPolitical As String * 1        'Y = political, else N (not used)
    sPostDate As String * 10        'posting transaction date (not used)
    sProduct As String * 35         'Product name (Not used)
    sRevenueCode1 As String         'RevenueCode1
    sSalesSource As String * 20     'sales source (not used)
    iSlspCode   As Integer          'Salesperson Code - TTP 10519 - Invoice Register Export - add summary version setting to current export
    sSlspName As String * 40        'Salesperson first/last name (AE Full Name)
    sSlspStnCode As String * 40     'Salesperson Code
    sTranDate As String * 10        'Transaction Date (not used)
    sTranType As String * 2         'Transaction type (not used)
End Type

Dim myBucket As CsiToAmazonS3.ApiCaller
Dim omBook As Object
Dim omSheet As Object
Dim imExcelRow As Integer
Dim lmRecordsProcessed As Long
Dim lmRecordsExported As Long
Dim imLastAdfCode As Integer
Dim smLastAdfName As String
Dim smAdfxDirectRefID As String
Dim imLastAgyCode As Integer
Dim smLastAgyName As String
Dim tmVGMNF() As MNF            'vehicle groups - TTP 10487

Private Sub ckcAmazon_Click()
    If ckcAmazon.Value = vbChecked Then
        frcAmazon.Left = 120
        frcAmazon.Top = 1560
        frcAmazon.Visible = True
    Else
        frcAmazon.Visible = False
    End If
End Sub

Private Sub ckcSummary_Click()
    'TTP 10612 - Invoice Register Summary export: new 7 column option
    If ckcSummary.Value = vbChecked Then
        frcSummary.Visible = True
    Else
        frcSummary.Visible = False
    End If
    mGetExportFilename
End Sub

Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub

Private Sub cmcExport_Click()
    Dim slToFile As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slRepeat As String
    Dim ilFileExists As Integer
    Dim blShowMesg As Boolean       'true: error returned from gathering data
    'Dim slClientName As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim ilTemp As Integer
    Dim slFullStartDate As String
    Dim slFullEndDate As String

    If ckcAmazon.Value = vbChecked Then
        If edcBucketName.Text = "" Or edcRegion.Text = "" Or edcAccessKey.Text = "" Or edcPrivateKey.Text = "" Then ckcAmazon.Value = vbUnchecked
    End If
    frcAmazon.Visible = False

    lacInfo(0).Visible = True
    lacInfo(1).Visible = False
    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    
    smExportName = Trim$(edcTo.Text)

    lmStart = gDateValue(CSI_CalStart.Text)
    lmEnd = gDateValue(CSI_CalEnd.Text)
    smFullStartDate = Format(lmStart, "ddddd")
    smFullEndDate = Format(lmEnd, "ddddd")
    gObtainYearMonthDayStr CSI_CalStart.Text, True, slYear, slMonth, slDay
    smStart = Trim$(slMonth) & Trim$(slDay) & Mid(slYear, 3, 2)
    gObtainYearMonthDayStr CSI_CalEnd.Text, True, slYear, slMonth, slDay
    smEnd = Trim$(slMonth) & Trim$(slDay) & Mid(slYear, 3, 2)
    
    'calendar control doesnt test for lack of slashes (xx/xx/xx); returns blank if no slashes
    If lmStart = 0 Or lmEnd = 0 Or (lmStart > lmEnd) Then
        ''MsgBox "Invalid Start and/or End Date Requested"
        gAutomationAlertAndLogHandler "Invalid Start and/or End Date Requested", vbOkOnly + vbCritical + vbApplicationModal, "ExportCashOrInv"
        Exit Sub
    End If
    'FileName = "InvReg mmddyy-mmddyy ClientName.csv" or "Cash mmddyy-mmddyy ClientName.csv"
    If Trim(smExportName) = "" Then mGetExportFilename
    Select Case imExportOption
        Case EXP_CASH, Exp_INVREG
            If InStr(1, smExportName, ".csv") = 0 Then smExportName = smExportName & ".csv"
            
        Case EXP_AUDACYINV
            'TTP 10260 - JW - 7/28/21 - Support CSV
            If optFormat(1).Value = True Then
                If InStr(1, smExportName, ".xls") = 0 Then smExportName = smExportName & ".xls"
            Else
                If InStr(1, smExportName, ".csv") = 0 Then smExportName = smExportName & ".csv"
            End If
    End Select
    
    slToFile = Trim$(smExportName)
    If (InStr(slToFile, ":") = 0) And (Left$(slToFile, 2) <> "\\") Then
        slToFile = sgExportPath & slToFile
    End If
    ilRet = 0
    'On Error GoTo cmcExportErr:
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        ''MsgBox slToFile & " already exists, Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gAutomationAlertAndLogHandler slToFile & " already exists, Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        Exit Sub
    Else
        ilRet = 0
        'hmTo = FreeFile
        'Open slToFile For Output As hmTo
        Select Case imExportOption
            Case EXP_CASH, Exp_INVREG
                ilRet = gFileOpen(slToFile, "Output", hmCashInv)
                If ilRet <> 0 Then
                    ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                    gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                    Exit Sub
                End If
            Case EXP_AUDACYINV
                'TTP 10260 - JW - 7/28/21 - Support CSV
                If optFormat(1).Value = True Then
                    'We will open Excel later
                Else
                    ilRet = gFileOpen(slToFile, "Output", hmCashInv)
                    If ilRet <> 0 Then
                        ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                        gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                        Exit Sub
                    End If
                End If
        End Select
    End If

    smExportName = slToFile
    edcTo.Text = smExportName
    Screen.MousePointer = vbHourglass
    imExporting = True
    
    On Error GoTo 0
    ilRet = 0
    'On Error GoTo cmcExportErr:
    'hmCashInv = FreeFile
    'Open smExportName For Output As hmCashInv
'        ilRet = gFileOpen(smExportName, "Output", hmCashInv)
'        If ilRet <> 0 Then
'            Print #hmMsg, "** Terminated **"
'            Close #hmMsg
'            Close #hmCashInv
'            imExporting = False
'            Screen.MousePointer = vbDefault
'            MsgBox "Open Error #" & str$(Err.Numner) & smExportName, vbOkOnly, "Open Error"
'            Exit Sub
'        End If
'        Print #hmMsg, "** Storing Output into " & smExportName & " **"
    
    Select Case imExportOption
        Case EXP_CASH, Exp_INVREG
            sgMessageFile = sgDBPath & "Messages\" & "ExportCashOrInv.txt"
            'gLogMsg smExportOptionName & " for: " & smStart & "-" & smEnd, "ExportCashOrInv.txt", False
            'gLogMsg smExportOptionName & "* Storing Output into " & smExportName,  "ExportCashOrInv.txt", False
            gAutomationAlertAndLogHandler "** Export " & smExportOptionName & " **"
            gAutomationAlertAndLogHandler "* Storing Output into " & smExportName
            gAutomationAlertAndLogHandler "* StartDate= " & CSI_CalStart.Text
            gAutomationAlertAndLogHandler "* EndDate= " & CSI_CalEnd.Text
            
        Case EXP_AUDACYINV
            sgMessageFile = sgDBPath & "Messages\" & "WOINV.txt"
            'gLogMsg "WOINV" & " for: " & smStart & "-" & smEnd, "WOINV.txt", False
            'gLogMsg "WOINV" & "* Storing Output into " & smExportName, "WOINV.txt", False
            gAutomationAlertAndLogHandler "** Export WOINV **"
            gAutomationAlertAndLogHandler "* Storing Output into " & smExportName
            gAutomationAlertAndLogHandler "* StartDate= " & CSI_CalStart.Text
            gAutomationAlertAndLogHandler "* EndDate= " & CSI_CalEnd.Text
            If ckcInclTrade.Value = vbChecked Then
                gAutomationAlertAndLogHandler "* InclTrade = True"
            Else
                gAutomationAlertAndLogHandler "* InclTrade = False"
            End If
            If ckcInclAdj.Value = vbChecked Then
                gAutomationAlertAndLogHandler "* InclAdj = True"
            Else
                gAutomationAlertAndLogHandler "* InclAdj = False"
            End If
            If optFormat(0).Value = True Then gAutomationAlertAndLogHandler "* Format = CSV"
            If optFormat(1).Value = True Then gAutomationAlertAndLogHandler "* Format = Excel"
            If ckcAmazon.Value = vbChecked Then
                gAutomationAlertAndLogHandler "* AmazonBucket=True"
            Else
                gAutomationAlertAndLogHandler "* AmazonBucket=False"
            End If

    End Select
    cmcExport.Enabled = False
    lmRecordsProcessed = 0
    
    'Export Records....
    blShowMesg = mObtainTransAndWrite()
    
    lacInfo(1).Caption = ""
    If Not imTerminate Then
        If blShowMesg = True Then        'error will be an error code
            lacInfo(0).Caption = "Call Counterpoint: Export has errors, see ExportCashOrInv.txt"
            'gLogMsg "Export Errors: Did Not Complete, Export File: " & slToFile, "ExportCashOrInv.txt", False
            gAutomationAlertAndLogHandler "Export Errors: Did Not Complete, Export File: " & slToFile
        Else
            lacInfo(0).Caption = "Export Successfully Completed"
            'gLogMsg "Export Successfully Completed, Export File: " & slToFile, "ExportCashOrInv.txt", False
            gAutomationAlertAndLogHandler "Export Successfully Completed, Export File: " & slToFile
        End If
        lacInfo(1).Caption = "Export Stored in: " & slToFile
    
        lacInfo(0).Visible = True
        lacInfo(1).Visible = True
    
        'Finish exporting
        Select Case imExportOption
            Case EXP_CASH, Exp_INVREG
                Close hmCashInv
            Case EXP_AUDACYINV
                If optFormat(1).Value = True Then
                    mDecorateExcel
                    'TTP 10260 - JW - 7/28/21 - Audacy wants a Named Range
                    'JW - 7/30/21 - Name Manager: Property -> WOImport
                    ilRet = gExcelOutputGeneration("NM", omBook, omSheet, , "WOImport")
                    mSaveExcel smExportName
                Else
                    'TTP 10260 - JW - 7/28/21 - Support CSV
                    Close hmCashInv
                End If
        End Select
    Else
        lacInfo(1).Caption = "Export Canceled: " & Now()
        lacInfo(1).Visible = True
    End If
    
    '----------------------------------
    If blShowMesg = False And Not imTerminate Then
        'AMAZON BUCKET SUPPORT:
        If ckcAmazon.Value = vbChecked And edcBucketName.Text <> "" And edcRegion.Text <> "" And edcAccessKey.Text <> "" And edcPrivateKey.Text <> "" Then
            If lmRecordsExported > 0 Then
                'Print #hmMsg, "** Uploading to " & edcBucketName.Text & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                gAutomationAlertAndLogHandler "** Uploading " & smExportName & " to " & AmazonBucketFolder(edcBucketName.Text, edcAmazonSubfolder.Text) & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                lacInfo(0).Caption = "Uploading " & smExportName
                lacInfo(0).Refresh
                DoEvents
                Set myBucket = New CsiToAmazonS3.ApiCaller
                On Error Resume Next
                err = 0
                'TTP 10504 - Amazon web bucket upload cleanup
'                If edcAmazonSubfolder.Text <> "" Then
'                    edcAmazonSubfolder.Text = Trim(Replace(edcAmazonSubfolder.Text, "\", "/"))
'                    If right(edcAmazonSubfolder.Text, 1) <> "/" Then edcAmazonSubfolder.Text = edcAmazonSubfolder.Text & "/"
'                    If Left(edcAmazonSubfolder.Text, 1) = "/" Then edcAmazonSubfolder.Text = Mid(edcAmazonSubfolder.Text, 2)
'                    '3/1/21 - added Folder support: "|" to split Bucket Name and Subfolder
'                    myBucket.UploadAmazonBucketFile edcBucketName.Text + "|" + edcAmazonSubfolder.Text, edcRegion.Text, edcAccessKey.Text, edcPrivateKey.Text, smExportName, False
'                Else
'                    myBucket.UploadAmazonBucketFile edcBucketName.Text, edcRegion.Text, edcAccessKey.Text, edcPrivateKey.Text, smExportName, False
'                End If
                myBucket.UploadAmazonBucketFile AmazonBucketFolder(edcBucketName.Text, edcAmazonSubfolder.Text), edcRegion.Text, edcAccessKey.Text, edcPrivateKey.Text, edcTo.Text, False
                If err <> 0 Then
                    lacInfo(0).Caption = "Error Uploading " & smExportName & " - " & err & " - " & Error(err)
                    'Print #hmMsg, "** Error Uploading " & smExportName & " - " & err & " - " & Error(err) & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                Else
                    If myBucket.ErrorMessage <> "" Then
                        lacInfo(0).Caption = "Error Uploading " & smExportName
                        'Print #hmMsg, "** Error Uploading " & smExportName & " - " & Replace(myBucket.ErrorMessage, vbCrLf, ";") & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                    Else
                        lacInfo(0).Caption = "Sucess Uploading " & smExportName
                        'Print #hmMsg, "** Finished Uploading " & smExportName & " - " & Replace(myBucket.Message, vbCrLf, ";") & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                        If ckcKeepLocalFile.Value = vbUnchecked Then
                            'We want to remove the Local File
                            Kill smExportName
                            lacInfo(1).Caption = "Deleted Local Export File: " & smExportName
                            'Print #hmMsg, "** Deleted Local Export File : " & smExportName & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                        End If
                    End If
                End If
                Set myBucket = Nothing
            Else
                'Print #hmMsg, "** Nothing to Upload to Amazon, Record Count : " & lmRecordsExported & " **"
                lacInfo(0).Caption = "Nothing to Upload, Record Count : " & lmRecordsExported
            End If
            'Print #hmMsg, "** Export " & Trim$(smExportOptionName) & " Successfully completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " , Record Count : " & lmRecordsExported & " **"
        End If
    Else
        lacInfo(0).Caption = "Export Failed"
        'Print #hmMsg, "** Export Failed **"
    End If
        
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    cmcExport.Enabled = False
    Screen.MousePointer = vbDefault
    imExporting = False
    
    'close left open Excel, if Canceled by user
    If Not omBook Is Nothing Then
        ogExcel.Quit
        Set omBook = Nothing
    End If
    Exit Sub
'cmcExportErr:
'      ilRet = Err.Number
'       Resume Next

ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub

Private Sub cmcTo_Click()
    CMDialogBox.DialogTitle = "Export To File"
    Select Case imExportOption
        Case EXP_CASH
            smExportOptionName = "Cash Receipts"
            CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
            CMDialogBox.DefaultExt = ".Csv"
        Case Exp_INVREG
            smExportOptionName = "Invoice Register"
            CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
            CMDialogBox.DefaultExt = ".Csv"
        Case EXP_AUDACYINV
            smExportOptionName = "WO Invoice" 'TTP 10205 - 6/21/21 - JW - WO Invoice Export - create new WO Invoice Export
            CMDialogBox.Filter = "Excel|*.XLS"
            CMDialogBox.DefaultExt = ".XLS"
    End Select
    
    CMDialogBox.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
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

Private Sub CSI_CalEnd_CalendarChanged()
    mGetExportFilename
End Sub

Private Sub CSI_CalEnd_Change()
    mGetExportFilename
End Sub

Private Sub CSI_CalStart_CalendarChanged()
    mGetExportFilename
End Sub

Private Sub CSI_CalStart_Change()
    mGetExportFilename
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    'Me.Refresh          'calendar control doesnt handle refresh command correctly
    'seems to be timing issue; need to do this kludge to turn off and on if refresh statement is in
'       CSI_CalStart.Visible = False
'       CSI_CalStart.Visible = True
'       CSI_CalEnd.Visible = False
'       CSI_CalEnd.Visible = True
    CSI_CalStart.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
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
        'cmcCancel_Click
        tmcCancel.Enabled = True
        Me.Left = 2 * Screen.Width      'move off the screen so screen won't flash
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
  
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
   
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmSof)
    btrDestroy hmSof

    Erase tmSofList, tmSSMnfList, tmNTRList
    Set ExpCashOrInv = Nothing   'Remove data segment

End Sub

Private Sub optFormat_Click(Index As Integer)
    mGetExportFilename
End Sub

Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub
Private Sub mInit()
    Dim ilRet As Integer
    Dim ilTemp As Integer
    
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    imExporting = False
    
    ckcInclTrade.Visible = False
    ckcInclAdj.Visible = False
    optFormat(0).Visible = False
    optFormat(1).Visible = False
    lacFileType.Visible = False
    ckcSummary.Visible = False
    
    imExportOption = ExportList!lbcExport.ItemData(ExportList!lbcExport.ListIndex)
    Select Case imExportOption
        Case EXP_CASH
            smExportOptionName = "Cash Receipts"
            
        Case Exp_INVREG
            smExportOptionName = "Invoice Register"
            'TTP 10487
            ReDim tmVGMNF(0 To 0) As MNF
            ilRet = gObtainMnfForType("H", "", tmVGMNF())      'vehicle groups
            ckcSummary.Top = ckcInclTrade.Top
            ckcSummary.Visible = True
            
        Case EXP_AUDACYINV
            smExportOptionName = "WO Invoice" 'TTP 10205 - 6/21/21 - JW - WO Invoice Export
            ckcInclTrade.Visible = True
            ckcInclAdj.Visible = True
            ckcInclTrade.Value = 0 'Include Trade: unchecked by default. When checked on, Trade invoices will be included.
            ckcInclAdj.Value = 0 'Include Adj: unchecked by default.
            optFormat(0).Visible = True
            optFormat(1).Visible = True
            lacFileType.Visible = True
    End Select
    
    lacTitle.Caption = "Export " & smExportOptionName
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "AGF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: AGF.Btr)", ExpCashOrInv
    imAgfRecLen = Len(tmAgf)
    
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "ADF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: ADF.Btr)", ExpCashOrInv
    imAdfRecLen = Len(tmAdf)
    
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Slf)", ExpCashOrInv
    imSlfRecLen = Len(tmSlf)
    
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf)", ExpCashOrInv
    imCHFRecLen = Len(tmChf)
    
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef)", ExpCashOrInv
    
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf)", ExpCashOrInv
    imMnfRecLen = Len(tmMnf)
    
    
    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Prf)", ExpCashOrInv
    imPrfRecLen = Len(tmPrf)
    
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf)", ExpCashOrInv
    imSofRecLen = Len(tmSof)
    
    ReDim tmSSMnfList(0 To 0) As MNF
    ilRet = gObtainMnfForType("S", smSSMnfStamp, tmSSMnfList())        'sales source
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    
    'Build array of Sales offices
    ilTemp = 0
    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tmSofList(0 To ilTemp) As SOF
        tmSofList(ilTemp) = tmSof
        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop
    
    ReDim tmNTRList(0 To 0) As MNF
    ilRet = gObtainMnfForType("I", smNTRMNFStamp, tmNTRList())      'ntr types
    
    Select Case imExportOption
        Case EXP_CASH, Exp_INVREG
            plcTo.Visible = False
            cmcTo.Visible = False
            lacSaveIn.Visible = False
            ckcAmazon.Visible = False
            
        Case EXP_AUDACYINV
            plcTo.Visible = True
            cmcTo.Visible = True
            lacSaveIn.Visible = True
            ckcAmazon.Visible = True
            
    End Select
    
    gCenterStdAlone ExpCashOrInv
    Screen.MousePointer = vbDefault
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    
    Exit Sub

mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mTerminate()
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpCashOrInv
    igManUnload = NO
End Sub

'           mGetallTables - read all supporting files to build export record for spot
'           <input>  none
'           <output> tmExport_TranInfo record contains info from supporting files
'           Return : blAllOK
Private Function mGetAllTables() As Boolean
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slRecord As String
    Dim llAmt As Long
    Dim slPrice As String
    Dim slSharePct As String
    Dim ilCommPct As Integer
    Dim slStr As String
    Dim llNet As Long
    Dim llGross As Long
    Dim llComm As Long
    Dim ilRemainder As Integer
    Dim slStripCents As String
    Dim ilLoopOnEntry As Integer
    Dim ilLoopTemp As Integer
    Dim blAllOK As Boolean
    Dim llDate As Long
    Dim blFound As Boolean

    On Error GoTo mGetAllTablesErr

    On Error GoTo 0
    blAllOK = True                      'assume all OK with retrieval of supporting files
    tmExport_TranInfo.sAccountID = ""   'agency or direct advertiser station code
    tmExport_TranInfo.sAgyName = ""     'Agency name
    tmExport_TranInfo.iAgyCode = 0      'Agency Code
    tmExport_TranInfo.sAdvName = ""     'Advertiser name
    tmExport_TranInfo.iAdvCode = 0      'Advertiser Code
    tmExport_TranInfo.sProduct = ""     'Product name
    tmExport_TranInfo.sSlspName = ""    'Salesperson first/last name
    tmExport_TranInfo.iSlspCode = 0      'Salesperson Code - TTP 10487
    tmExport_TranInfo.sSlspStnCode = "" 'Salesperson Code
    tmExport_TranInfo.sOffice = ""      'sales office
    tmExport_TranInfo.sSalesSource = "" 'sales source
    tmExport_TranInfo.sBusCat = ""      'Business category
    tmExport_TranInfo.sNTRType = ""     'NTR Item Name
    tmExport_TranInfo.sBillVehicle = "" 'Billing Vehicle Name
    tmExport_TranInfo.iAirVehicleCode = 0 'Airing Vehicle Code - TTP 10487
    tmExport_TranInfo.sAirVehicle = ""  'Airing vehicle name
    tmExport_TranInfo.sContract = ""    'contract#
    tmExport_TranInfo.sTranDate = ""    'Transaction Date
    tmExport_TranInfo.sInvNo = ""       'Invoice #
    tmExport_TranInfo.sTranType = ""    'transaction type
    tmExport_TranInfo.sAction = ""      'Action on Payment or journal entry
    tmExport_TranInfo.sPostDate = ""    'posting transaction date
    tmExport_TranInfo.sCheck = ""       'check number
    tmExport_TranInfo.sNet = ""         'net amount
    tmExport_TranInfo.sGross = ""       'gross amount
    tmExport_TranInfo.sComm = ""        'commission
    tmExport_TranInfo.sCashTrade = ""   'C = Cash ,T = trade
    tmExport_TranInfo.sPolitical = "N"  'Y = political, else N
    
    'Common fields with Cash Receipts and Inv Register
    tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
    tmChfSrchKey1.iCntRevNo = 32000
    tmChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd

    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F")
         ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop

    gFakeChf tmRvf, tmChf       'create a bare-bone header if the contract doesnt exist
    tmExport_TranInfo.sContract = str$(tmRvf.lCntrNo)
    
    If tmChf.iMnfBus > 0 Then
        blFound = False
        For ilLoop = 0 To UBound(tgBusCatMnf) - 1
            If tgBusCatMnf(ilLoop).iCode = tmChf.iMnfBus Then
                tmExport_TranInfo.sBusCat = Trim$(tgBusCatMnf(ilLoop).sName)
                blFound = True
                Exit For
            End If
        Next ilLoop
        If Not blFound Then
            blAllOK = False
            mMissingID CLng(tmChf.iMnfBus), "Invalid Business Category ID"
        End If
    End If

    If tmRvf.lPrfCode > 0 Then
        tmPrfSrchKey.lCode = tmRvf.lPrfCode
        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) Then
            tmExport_TranInfo.sProduct = tmPrf.sName
        End If
    End If
    
    'get agency or advt name
    If tmRvf.iAdfCode > 0 Then
        ilLoop = gBinarySearchAdf(tmRvf.iAdfCode)
        If ilLoop <> -1 Then
            tmExport_TranInfo.sAdvName = Trim$(tgCommAdf(ilLoop).sName)
            tmExport_TranInfo.iAdvCode = tgCommAdf(ilLoop).iCode
            'is it political?
            If gIsItPolitical(tgCommAdf(ilLoop).iCode) Then            'its a political, include this contract?
                tmExport_TranInfo.sPolitical = "Y"
            Else
                tmExport_TranInfo.sPolitical = "N"
            End If
            If tmRvf.iAgfCode = 0 Then
                tmExport_TranInfo.sAccountID = Trim$(tgCommAdf(ilLoop).sCodeStn)              'need to get station code
            End If
        Else
            blAllOK = False
            mMissingID CLng(tmRvf.iAdfCode), "Invalid Advertiser ID"
          End If
    End If
    If tmRvf.iAgfCode > 0 Then
        ilLoop = gBinarySearchAgf(tmRvf.iAgfCode)
        If ilLoop <> -1 Then
            tmExport_TranInfo.sAgyName = Trim$(tgCommAgf(ilLoop).sName)
            tmExport_TranInfo.iAgyCode = Trim$(tgCommAgf(ilLoop).iCode)
            tmExport_TranInfo.sAccountID = Trim$(tgCommAgf(ilLoop).sCodeStn)
        Else
            blAllOK = False
            mMissingID CLng(tmRvf.iAgfCode), "Invalid Agency ID"
        End If
    End If
    tmExport_TranInfo.iSlspCode = 0 'TTP 10487
    If tmRvf.iSlfCode > 0 Then
        ilLoop = gBinarySearchSlf(tmRvf.iSlfCode)
        If ilLoop <> -1 Then
            tmExport_TranInfo.sSlspName = Trim$(tgMSlf(ilLoop).sFirstName) + " " + Trim$(tgMSlf(ilLoop).sLastName)
            tmExport_TranInfo.sSlspStnCode = Trim(tgMSlf(ilLoop).sCodeStn)
            tmExport_TranInfo.iSlspCode = tmRvf.iSlfCode 'TTP 10487
            'from slsp, get the sales office and source
            For ilLoopOnEntry = 0 To UBound(tmSofList)
                If tmSofList(ilLoopOnEntry).iCode = tgMSlf(ilLoop).iSofCode Then     'matching sales offices
                    tmExport_TranInfo.sOffice = tmSofList(ilLoopOnEntry).sName
                    For ilLoopTemp = 0 To UBound(tmSSMnfList) - 1
                        If tmSSMnfList(ilLoopTemp).iCode = tmSofList(ilLoopOnEntry).iMnfSSCode Then
                            tmExport_TranInfo.sSalesSource = tmSSMnfList(ilLoopTemp).sName
                            Exit For
                        End If
                    Next ilLoopTemp
                End If
            Next ilLoopOnEntry
        Else
            blAllOK = False
            mMissingID CLng(tmRvf.iSlfCode), "Invalid Salesperson ID"
        End If
    End If
    
    tmExport_TranInfo.iAirVehicleCode = 0 'TTP 10487
    If tmRvf.iAirVefCode > 0 Then
        ilLoop = gBinarySearchVef(tmRvf.iAirVefCode)
        If ilLoop <> -1 Then
            tmExport_TranInfo.sAirVehicle = Trim$(tgMVef(ilLoop).sName)
            tmExport_TranInfo.iAirVehicleCode = tmRvf.iAirVefCode 'TTP 10487
        Else
            blAllOK = False
            mMissingID CLng(tmRvf.iAirVefCode), "Invalid Airing Vehicle ID"
         End If
    End If
    
    If tmRvf.iBillVefCode > 0 Then
        ilLoop = gBinarySearchVef(tmRvf.iBillVefCode)
        If ilLoop <> -1 Then
            tmExport_TranInfo.sBillVehicle = Trim$(tgMVef(ilLoop).sName)
        Else
            blAllOK = False
            mMissingID CLng(tmRvf.iBillVefCode), "Invalid Billing Vehicle ID"
         End If
    End If
    
    If tmRvf.iMnfItem > 0 Then              'NTR exist?
        For ilLoopOnEntry = 0 To UBound(tmNTRList) - 1
            If tmNTRList(ilLoopOnEntry).iCode = tmRvf.iMnfItem Then
                tmExport_TranInfo.sNTRType = tmNTRList(ilLoopOnEntry).sName
                Exit For
            End If
        Next ilLoopOnEntry
    End If
        
    gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
    tmExport_TranInfo.sTranDate = Format$(llDate, "m/d/yyyy")
    tmExport_TranInfo.sInvNo = str$(tmRvf.lInvNo)
    tmExport_TranInfo.sTranType = tmRvf.sTranType
    
    gPDNToLong tmRvf.sNet, llNet
    
    'On Error GoTo mGetAllTablesErr
    dmNet = dmNet + (llNet / 100)             'accumulate Net control total for end of export / TTP 10208
    slStr = mRemoveNoCents(llNet)
    If Trim$(slStr) = "" Then slStr = "0"
    tmExport_TranInfo.sNet = Trim$(slStr)
    tmExport_TranInfo.sCashTrade = tmRvf.sCashTrade
               
    Select Case imExportOption
        Case EXP_CASH   'cash receipts
            tmExport_TranInfo.sAction = tmRvf.sAction
            gUnpackDateLong tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), llDate
            tmExport_TranInfo.sPostDate = Format$(llDate, "m/d/yyyy")
            tmExport_TranInfo.sCheck = Trim$(tmRvf.sCheckNo)
            
        Case Exp_INVREG
            gPDNToLong tmRvf.sGross, llGross
            'On Error GoTo mGetAllTablesErr
            dmGross = dmGross + (llGross / 100)             'accumulate Gross control total for end of export / TTP 10208
            slStr = mRemoveNoCents(llGross)
            If Trim$(slStr) = "" Then slStr = "0"
            tmExport_TranInfo.sGross = Trim$(slStr)
            llComm = llGross - llNet
            dmComm = dmComm + (llComm / 100)             'accumulate Comm control total for end of export / TTP 10208
            slStr = mRemoveNoCents(llComm)
            If Trim$(slStr) = "" Then slStr = "0"
            tmExport_TranInfo.sComm = Trim$(slStr)
            
        Case EXP_AUDACYINV
            
    End Select
           
    mGetAllTables = blAllOK     'return error flag
    Exit Function
mGetAllTablesErr:
        Resume Next
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
                Screen.MousePointer = vbDefault
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
    'Print #hmMsg, ""
    'Print #hmMsg, "** Export " & Trim$(smExportOptionName) & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " for " & smStart & "-" & smEnd
    gAutomationAlertAndLogHandler "** Export " & Trim$(smExportOptionName) & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " for " & smStart & "-" & smEnd
    mOpenMsgFile = True
    Exit Function
End Function

'
'           mGetPHFRVF -process phf and rvf to get either Cash/Journal Entry transactions or
'           Invoice/Adjustment transations
'           Return - true if error found, else false
Public Function mObtainTransAndWrite() As Boolean
    Dim slStr As String
    Dim blErrorFound As Boolean
    Dim llLoopOnTrans As Long
    Dim tlTranTypes As TRANTYPES
    ReDim tlRvf(0 To 0) As RVF
    Dim ilRet As Integer
    Dim blAllOK As Boolean      'true if everything found in getting tables
    Dim slRecord As String
    Dim slDelimiter As String
    ReDim tmExport_TranSummary(0) As EXPORT_TRANSUMMARY
    slDelimiter = Chr(30)
    blErrorFound = False
    Dim sSubCompany As String
    Dim sSubtotals As String
    Dim sMarket As String
    Dim sResearch As String
    Dim sFormat As String
    
    'create headers
    Select Case imExportOption
        Case EXP_CASH
            slStr = "Account ID,Agency,Advertiser,Product,Salesperson,Sales Office,Sales Source,Business Category,"
            slStr = slStr & "NTR Type,Billing Vehicle,Airing Vehicle,Contract #,Tran Date,Invoice #,Tran Type,"
            slStr = slStr & "Action,Posting Date,Check #,Net Amt,Cash/Trade,Political"
            
        Case Exp_INVREG
            'TTP 10487 - Invoice Register export: add several new fields
            'airing vehicle groups (sub-company, subtotals, market, research, format)
            'Advertiser Reference ID (adfxRefID
            'Agency Reference ID (agfxRefID)
            'Salesperson code (slfCode)
            'Station salesperson code (slfCodeStn)
            
            'TTP 10519 - Invoice Register Export - add summary version setting to current export
            If ckcSummary.Value = vbChecked Then
                'TTP 10612 - Invoice Register Summary export: new 7 column option
                If optSummaryFormat1.Value = True Then
                    '22-Column Format
                    slStr = "Account ID,Agency,Agency Reference ID,Advertiser,Advertiser Reference ID,Product,"
                    slStr = slStr & "Salesperson,Salesperson Code,Sales Office,Sales Source,Station Salesperson Code,Business Category,"
                    slStr = slStr & "NTR Type,Contract #,Tran Date,Invoice #,Tran Type,"
                    slStr = slStr & "Gross Amt,Comm Amt,Net Amt,Cash/Trade,Political"
                End If
                If optSummaryFormat2.Value = True Then
                    ''7-Column Format
                    'slStr = "Agency,Advertiser,Product,"
                    'slStr = slStr & "Contract #,Tran Date,"
                    'slStr = slStr & "Invoice #,Net Amt"
                    
                    'TTP 10626 - AURN Sage Intacct Export
                    slStr = "INVOICE_NO,PO_NO,CUSTOMER_ID,CREATED_DATE,TOTAL_DUE,TERM_NAME,DESCRIPTION,LINE_NO,ACCT_NO,LOCATION_ID,DEPT_ID,AMOUNT,ARINVOICEITEM_ARACCOUNT,ARINVOICEITEM_CUSTOMERID"
                End If
            Else
                slStr = "Account ID,Agency,Agency Reference ID,Advertiser,Advertiser Reference ID,Product,"
                slStr = slStr & "Salesperson,Salesperson Code,Sales Office,Sales Source,Station Salesperson Code,Business Category,"
                slStr = slStr & "NTR Type,Billing Vehicle,Airing Vehicle,Sub-company,Subtotals,Market,Research,Format,"
                slStr = slStr & "Contract #,Tran Date,Invoice #,Tran Type,"
                slStr = slStr & "Gross Amt,Comm Amt,Net Amt,Cash/Trade,Political"
            End If
            
        Case EXP_AUDACYINV
            slStr = "Account ID,Agency,Advertiser,Product,Salesperson,Sales Office,Sales Source,Business Category,"
            slStr = slStr & "NTR Type,Billing Vehicle,Airing Vehicle,Contract #,Tran Date,Invoice #,Tran Type,"
            slStr = "Property,Advertiser,AdvertiserGUID,Agency,AgencyGUID,AE Full Name,AECode,Sales Office,SalesOfficeCode,RevenueCode1,RevenueCode2,RevenueCode3,Date,Order#,Invoice#,Balance,BillingGroup"
            
    End Select
    
    'write header description
    Select Case imExportOption
        Case EXP_CASH, Exp_INVREG
            On Error GoTo mWriteErr
            Print #hmCashInv, slStr
                
        Case EXP_AUDACYINV
            'Open Excel
            If CreateExcel = False Then
                ''MsgBox "Export Canceled"
                gAutomationAlertAndLogHandler "Export Canceled", vbOkOnly, "ExpCashOrInv"
                blErrorFound = True
                Exit Function
            End If
            'write header
            'TTP 10260 - JW - 7/28/21 - Support CSV
            If optFormat(1).Value = True Then
                mWriteExcel slStr
            Else
                Print #hmCashInv, slStr
            End If
    End Select
    
    On Error GoTo 0
    tlTranTypes.iAdj = True
    tlTranTypes.iAirTime = True
    tlTranTypes.iCash = True
    tlTranTypes.iHardCost = True
    tlTranTypes.iInv = True
    tlTranTypes.iMerch = False
    tlTranTypes.iNTR = True
    tlTranTypes.iPromo = False
    tlTranTypes.iPymt = True
    tlTranTypes.iTrade = True
    tlTranTypes.iWriteOff = True
    Select Case imExportOption
        Case EXP_CASH 'cash, turn off invoices and adjustments
            tlTranTypes.iAdj = False
            tlTranTypes.iInv = False
        Case Exp_INVREG  'inv reg - turn off payments and journal entries
            tlTranTypes.iPymt = False
            tlTranTypes.iWriteOff = False
        Case EXP_AUDACYINV 'Includes: air time, NTR, Cash, Political, and Non-Political types.
                           'Optional: Trade (optional),Invoice Adjustments (AN Transactions)
                           'Excluded: TranType IN, HI=histInv, all else else...
            tlTranTypes.iPymt = False
            tlTranTypes.iWriteOff = False
            tlTranTypes.iHardCost = False
            If ckcInclTrade.Value = 0 Then tlTranTypes.iTrade = False
            If ckcInclAdj.Value = 0 Then tlTranTypes.iAdj = False
    End Select
    'TTP 10208
    dmGross = 0
    dmNet = 0
    dmComm = 0

    '-------------------------------------------
    '0= (last parm) indicates retrieve by tran date vs entered date
    ilRet = gObtainPhfRvf(ExpCashOrInv, smFullStartDate, smFullEndDate, tlTranTypes, tlRvf(), 0)
    
    For llLoopOnTrans = LBound(tlRvf) To UBound(tlRvf) - 1
        tmRvf = tlRvf(llLoopOnTrans)
        blAllOK = mGetAllTables()
        Select Case imExportOption
            Case EXP_CASH
                slRecord = """" & Trim$(tmExport_TranInfo.sAccountID) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sAgyName) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sAdvName) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sProduct) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sSlspName) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sOffice) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sSalesSource) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sBusCat) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sNTRType) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sBillVehicle) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sAirVehicle) & """" & ","
                slRecord = slRecord & Trim$(tmExport_TranInfo.sContract) & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sTranDate) & """" & ","
                slRecord = slRecord & Trim$(tmExport_TranInfo.sInvNo) & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sTranType) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sAction) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sPostDate) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sCheck) & """" & ","
                slRecord = slRecord & Trim$(tmExport_TranInfo.sNet) & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sCashTrade) & """" & ","
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sPolitical) & """"
            
            Case Exp_INVREG
                'TTP 10487 - Invoice Register export: add several new fields
                slRecord = """" & Trim$(tmExport_TranInfo.sAccountID) & """" & ","              'Account ID
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sAgyName) & """" & ","     'Agency
                slRecord = slRecord & """" & mGetAgfxRefID(tmExport_TranInfo.iAgyCode) & """" & "," 'Agency Reference ID - TTP 10487
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sAdvName) & """" & ","     'Advertiser
                slRecord = slRecord & """" & mGetAdfxRefID(tmExport_TranInfo.iAdvCode) & """" & "," 'Advertiser Reference ID - TTP 10487
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sProduct) & """" & ","     'Product
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sSlspName) & """" & ","    'Salesperson
                slRecord = slRecord & Trim$(tmExport_TranInfo.iSlspCode) & ","                  'Salesperson Code - TTP 10487
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sOffice) & """" & ","      'Sales Office
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sSalesSource) & """" & "," 'Sales Source
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sSlspStnCode) & """" & "," 'Station Salesperson Code - TTP 10487
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sBusCat) & """" & ","      'Business Category
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sNTRType) & """" & ","     'NTR Type
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sBillVehicle) & """" & "," 'Billing Vehicle
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sAirVehicle) & """" & ","  'Airing Vehicle
                mGetAllVefGroupNames tmExport_TranInfo.iAirVehicleCode, sSubCompany, sSubtotals, sMarket, sResearch, sFormat
                slRecord = slRecord & """" & Trim$(sSubCompany) & """" & ","                    'Sub-company - TTP 10487
                slRecord = slRecord & """" & Trim$(sSubtotals) & """" & ","                     'Subtotals - TTP 10487
                slRecord = slRecord & """" & Trim$(sMarket) & """" & ","                        'Market - TTP 10487
                slRecord = slRecord & """" & Trim$(sResearch) & """" & ","                      'Research - TTP 10487
                slRecord = slRecord & """" & Trim$(sFormat) & """" & ","                        'Format - TTP 10487
                slRecord = slRecord & Trim$(tmExport_TranInfo.sContract) & ","                  'Contract #
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sTranDate) & """" & ","    'Tran Date
                slRecord = slRecord & Trim$(tmExport_TranInfo.sInvNo) & ","                     'Invoice #
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sTranType) & """" & ","    'Tran Type
                slRecord = slRecord & Trim$(tmExport_TranInfo.sGross) & ","                     'Gross Amt
                slRecord = slRecord & Trim$(tmExport_TranInfo.sComm) & ","                      'Comm Amt
                slRecord = slRecord & Trim$(tmExport_TranInfo.sNet) & ","                       'Net Amt
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sCashTrade) & """" & ","   'Cash/Trade
                slRecord = slRecord & """" & Trim$(tmExport_TranInfo.sPolitical) & """"         'Political
            
            Case EXP_AUDACYINV
                slRecord = "" 'Lines Written after TransSummary built
        End Select

        'Write record
        On Error GoTo mWriteErr
        Select Case imExportOption
            Case EXP_CASH
                Print #hmCashInv, slRecord
                
            Case Exp_INVREG
                'TTP 10519 - Invoice Register Export - add summary version setting to current export
                If ckcSummary.Value = vbChecked Then
                    mUpdateTranSummary2 tmExport_TranInfo
                Else
                    Print #hmCashInv, slRecord
                End If
                
            Case EXP_AUDACYINV
                'Save info in Trans Summary
                mUpdateTranSummary tmExport_TranInfo
        End Select
        
        On Error GoTo 0
        If igDOE >= 500 Then
            lacInfo(0).Caption = "Processing " & lmRecordsProcessed & " records..."
            igDOE = 0
            DoEvents
        End If
        igDOE = igDOE + 1
        
        If imTerminate Then
            ''MsgBox "Export Canceled by user", vbInformation + vbOkOnly, "Export"
            gAutomationAlertAndLogHandler "Export Canceled by user", vbInformation, "Export"
            
            Set ogExcel = Nothing
            Set omBook = Nothing
            
            blErrorFound = True
            Exit For
            Exit Function
        End If
        lmRecordsProcessed = lmRecordsProcessed + 1
    Next llLoopOnTrans
    
    gAutomationAlertAndLogHandler "Processed " & lmRecordsProcessed & " records..."
    lacInfo(0).Caption = "Processed " & lmRecordsProcessed & " records..."
    lacInfo(0).Refresh
    DoEvents
    
    '-------------------------------------------
    Select Case imExportOption
        Case Exp_INVREG
            'TTP 10519 - Invoice Register Export - add summary version setting to current export
            If ckcSummary.Value = vbChecked Then
                If Not mExportTranSummary2 Then blErrorFound = True
            End If
        
        Case EXP_AUDACYINV
            If Not mExportTranSummary Then blErrorFound = True
    End Select
    
    '-------------------------------------------
    'output control total
    Select Case imExportOption
        Case EXP_CASH
            slRecord = """" & "Totals" & """" & "," 'Account ID,
            slRecord = slRecord & """" & """" & "," 'Agency,
            slRecord = slRecord & """" & """" & "," 'Advertiser,
            slRecord = slRecord & """" & """" & "," 'Product,
            slRecord = slRecord & """" & """" & "," 'Salesperson,
            slRecord = slRecord & """" & """" & "," 'Sales Office,
            slRecord = slRecord & """" & """" & "," 'Sales Source,
            slRecord = slRecord & """" & """" & "," 'Business Category,
            slRecord = slRecord & """" & """" & "," 'NTR Type,
            slRecord = slRecord & """" & """" & "," 'Billing Vehicle,
            slRecord = slRecord & """" & """" & "," 'Airing Vehicle,
            slRecord = slRecord & "," 'Contract #,
            slRecord = slRecord & """" & """" & "," 'Tran Date,
            slRecord = slRecord & "," 'Invoice #,
            slRecord = slRecord & """" & """" & "," 'Tran Type,
            slRecord = slRecord & """" & """" & "," 'Action,
            slRecord = slRecord & """" & """" & "," 'Posting Date,
            slRecord = slRecord & """" & """" & "," 'Check #,
            slStr = mRemoveNoCentsDbl(dmNet)
            If Trim$(slStr) = "" Then
                slStr = "0"
            End If
            tmExport_TranInfo.sNet = Trim$(slStr)
            slRecord = slRecord & Trim$(tmExport_TranInfo.sNet) & "," 'Net Amt,
            slRecord = slRecord & """" & """" & "," 'Cash/Trade,
            slRecord = slRecord & """" & """" 'Political
            On Error GoTo mWriteErr
            Print #hmCashInv, slRecord
            
        Case Exp_INVREG
            If ckcSummary.Value = vbChecked Then
                'TTP 10612 - Invoice Register Summary export: new 7 column option
                If optSummaryFormat1.Value = True Then
                    '22-Column format
                    slRecord = """" & "Totals" & """" & "," 'Account ID
                    slRecord = slRecord & """" & """" & "," 'Agency
                    slRecord = slRecord & """" & """" & "," 'Agency Reference ID
                    slRecord = slRecord & """" & """" & "," 'Advertiser
                    slRecord = slRecord & """" & """" & "," 'Advertiser Reference ID
                    slRecord = slRecord & """" & """" & "," 'Product
                    slRecord = slRecord & """" & """" & "," 'Salesperson
                    slRecord = slRecord & "," 'Salesperson Code
                    slRecord = slRecord & """" & """" & "," 'Sales Office
                    slRecord = slRecord & """" & """" & "," 'Sales Source
                    slRecord = slRecord & """" & """" & "," 'Station Salesperson Code
                    slRecord = slRecord & """" & """" & "," 'Business Category
                    slRecord = slRecord & """" & """" & "," 'NTR Type
                    slRecord = slRecord & "," 'Contract #
                    slRecord = slRecord & """" & """" & "," 'Tran Date
                    slRecord = slRecord & "," 'Invoice #
                    slRecord = slRecord & """" & """" & "," 'Tran Type
                    slStr = mRemoveNoCentsDbl(dmGross)
                    If Trim$(slStr) = "" Then
                        slStr = "0"
                    End If
                    tmExport_TranInfo.sGross = Trim$(slStr)
                    slRecord = slRecord & Trim$(tmExport_TranInfo.sGross) & "," 'Gross Amt
                    slStr = mRemoveNoCentsDbl(dmComm)
                    If Trim$(slStr) = "" Then
                        slStr = "0"
                    End If
                    tmExport_TranInfo.sComm = Trim$(slStr)
                    slRecord = slRecord & Trim$(tmExport_TranInfo.sComm) & "," 'Comm Amt
                    slStr = mRemoveNoCentsDbl(dmNet)
                    If Trim$(slStr) = "" Then
                        slStr = "0"
                    End If
                    tmExport_TranInfo.sNet = Trim$(slStr)
                    slRecord = slRecord & Trim$(tmExport_TranInfo.sNet) & "," 'Net Amt
                    slRecord = slRecord & """" & """" & "," 'Cash/Trade
                    slRecord = slRecord & """" & """" 'Political
                End If
                
                'TTP 10612 - Invoice Register Summary export: new 7 column option
                If optSummaryFormat2.Value = True Then
                    'TTP 10626 - AURN Sage Intacct Export: Suppress total from final row of export
                    ''7-Column format
                    'slRecord = """" & "Totals" & """" & "," 'Agency
                    'slRecord = slRecord & """" & """" & "," 'Advertiser
                    'slRecord = slRecord & """" & """" & "," 'Product
                    'slRecord = slRecord & "," 'Contract #
                    'slRecord = slRecord & """" & """" & "," 'Tran Date
                    'slRecord = slRecord & "," 'Invoice #
                    'slStr = mRemoveNoCentsDbl(dmNet)
                    'If Trim$(slStr) = "" Then
                    '    slStr = "0"
                    'End If
                    'tmExport_TranInfo.sNet = Trim$(slStr)
                    'slRecord = slRecord & Trim$(tmExport_TranInfo.sNet)  'Net Amt
                    slRecord = ""
                End If
            Else
                slRecord = """" & "Totals" & """" & "," 'Account ID
                slRecord = slRecord & """" & """" & "," 'Agency
                slRecord = slRecord & """" & """" & "," 'Agency Reference ID
                slRecord = slRecord & """" & """" & "," 'Advertiser
                slRecord = slRecord & """" & """" & "," 'Advertiser Reference ID
                slRecord = slRecord & """" & """" & "," 'Product
                slRecord = slRecord & """" & """" & "," 'Salesperson
                slRecord = slRecord & "," 'Salesperson Code
                slRecord = slRecord & """" & """" & "," 'Sales Office
                slRecord = slRecord & """" & """" & "," 'Sales Source
                slRecord = slRecord & """" & """" & "," 'Station Salesperson Code
                slRecord = slRecord & """" & """" & "," 'Business Category
                slRecord = slRecord & """" & """" & "," 'NTR Type
                slRecord = slRecord & """" & """" & "," 'Billing Vehicle
                slRecord = slRecord & """" & """" & "," 'Airing Vehicle
                slRecord = slRecord & """" & """" & "," 'Sub-company
                slRecord = slRecord & """" & """" & "," 'Subtotals
                slRecord = slRecord & """" & """" & "," 'Market
                slRecord = slRecord & """" & """" & "," 'Research
                slRecord = slRecord & """" & """" & "," 'Format
                slRecord = slRecord & "," 'Contract #
                slRecord = slRecord & """" & """" & "," 'Tran Date
                slRecord = slRecord & "," 'Invoice #
                slRecord = slRecord & """" & """" & "," 'Tran Type
                slStr = mRemoveNoCentsDbl(dmGross)
                If Trim$(slStr) = "" Then
                    slStr = "0"
                End If
                tmExport_TranInfo.sGross = Trim$(slStr)
                slRecord = slRecord & Trim$(tmExport_TranInfo.sGross) & "," 'Gross Amt
                slStr = mRemoveNoCentsDbl(dmComm)
                If Trim$(slStr) = "" Then
                    slStr = "0"
                End If
                tmExport_TranInfo.sComm = Trim$(slStr)
                slRecord = slRecord & Trim$(tmExport_TranInfo.sComm) & "," 'Comm Amt
                slStr = mRemoveNoCentsDbl(dmNet)
                If Trim$(slStr) = "" Then
                    slStr = "0"
                End If
                tmExport_TranInfo.sNet = Trim$(slStr)
                slRecord = slRecord & Trim$(tmExport_TranInfo.sNet) & "," 'Net Amt
                slRecord = slRecord & """" & """" & "," 'Cash/Trade
                slRecord = slRecord & """" & """" 'Political
            End If
            On Error GoTo mWriteErr
            
            'TTP 10626 - AURN Sage Intacct Export: Suppress total from final row of export
            If optSummaryFormat2.Value = False Then
                Print #hmCashInv, slRecord
            End If
            
        Case EXP_AUDACYINV
            'Dont export the totals
    End Select

    On Error GoTo 0
    mObtainTransAndWrite = blErrorFound
    
    Erase tlRvf
    Exit Function
    
mWriteErr:
    blErrorFound = True
    Select Case imExportOption
            Case EXP_CASH, Exp_INVREG
                gLogMsg "Error Writing Record, Export File: " & smExportName, "ExportCashOrInv.txt", False
            Case EXP_AUDACYINV
                gLogMsg "Error Writing Record, Export File: " & smExportName, "EXPINV.txt", False
    End Select

    Resume Next
End Function
'
'           mRemoveNoCents - determine if even $ amt and remove .00
'           <input>  amount to test if pennies exist
'           return - string of converted Amt, no $, commas or periods unless cents exist
Public Function mRemoveNoCents(llAmt As Long)
    Dim slStr As String
    Dim ilRemainder As Integer
    Dim slStripCents As String

    slStr = ""
    ilRemainder = llAmt Mod 100           'Detrmine if cents exist
    If ilRemainder = 0 Then         'strip off the pennies if whole number
        slStripCents = Trim$(gLongToStrDec(llAmt, 2))
        slStr = Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
    Else
        slStr = Trim$(gLongToStrDec(llAmt, 2))
    End If
    mRemoveNoCents = slStr
End Function
'
'           mRemoveNoCentsDbl - determine if even $ amt and remove .00
'           <input>  amount to test if pennies exist
'           return - string of converted Amt, no $, commas or periods unless cents exist
Public Function mRemoveNoCentsDbl(dlAmt As Double)
    Dim slStr As String
    slStr = ""
    If dlAmt = CLng(dlAmt) Then        'strip off the pennies if whole number
        slStr = Format(dlAmt, "0")
    Else
        slStr = Format(dlAmt, "0.00")
    End If
    mRemoveNoCentsDbl = CDbl(slStr)
End Function

Public Sub mMissingID(llMissingID As Long, slInvalidMsg As String)
    Dim slWhichFile As String
    If tmRvf.sTranType = "HI" Or tmRvf.iPurgeDate(0) <> 0 Then      'its History
        slWhichFile = "History"
    Else
        slWhichFile = "Receivables"
    End If
    
    gLogMsg slInvalidMsg & "(" & Trim$(str(llMissingID)) & ") in " & slWhichFile & " for Contract #: " & Trim$(str(tmRvf.lCntrNo)) & ".", "ExportCashOrInv.txt", False
    Exit Sub
End Sub

Function CreateExcel() As Boolean
    Dim ilRet As Integer
    CreateExcel = False
    ilRet = gExcelOutputGeneration("O", omBook, omSheet, 1)
    If ilRet = False Then
        ''MsgBox "Unable to Generate Export, Excel in use." & vbCrLf & "Please Close Excel and try again", vbCritical + vbOkOnly, "Generate Export Error"
        gAutomationAlertAndLogHandler "Unable to Generate Export, Excel in use." & vbCrLf & "Please Close Excel and try again", vbOkOnly + vbCritical, "Generate Export Error"
        GoTo mErrorSkipHere
    End If
    CreateExcel = True
    Exit Function
    
mErrorSkipHere:
    
End Function

Private Sub mUpdateTranSummary(tlExport_TranInfo As EXPORT_TRANINFO)
    Dim ilLoop As Integer
    Dim blFound As Boolean
    Dim slTemp As String
    Dim ilMsgbox As Integer
    
    'Check tmInvDistSummary for existing stuff
    For ilLoop = 0 To UBound(tmExport_TranSummary)
        blFound = False
        If Trim(tmExport_TranSummary(ilLoop).sInvNo) = Trim(tlExport_TranInfo.sInvNo) Then
            blFound = True
            Exit For
        End If
    Next ilLoop
    If blFound = True Then
        'Update the existing Invoice in the Array
        'tmExport_TranSummary(ilLoop).sGross = Val(tmExport_TranSummary(ilLoop).sGross) + Val(Trim(tlExport_TranInfo.sGross))
        tmExport_TranSummary(ilLoop).sNet = Val(tmExport_TranSummary(ilLoop).sNet) + Val(Trim(tlExport_TranInfo.sNet))
    Else
        'Add Invoice to Array
        tmExport_TranSummary(UBound(tmExport_TranSummary)).iAdvCode = tlExport_TranInfo.iAdvCode
        tmExport_TranSummary(UBound(tmExport_TranSummary)).iAgyCode = tlExport_TranInfo.iAgyCode
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sAccountID = Trim(tlExport_TranInfo.sAccountID)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sAction = Trim(tlExport_TranInfo.sAction)
        
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAdvertiserRefId = mGetAdfxRefID(tlExport_TranInfo.iAdvCode)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAdvName = Trim(tlExport_TranInfo.sAdvName)
        
        'JW - 7/29/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
        'will modify the WO invoice export so that when it outputs the data for a direct advertiser, if there's no agency defined, it will use the ADFXDirectRefID as the AgencyGUID for that record.
        slTemp = mGetAgfxRefID(tlExport_TranInfo.iAgyCode)
        If slTemp = "" Then slTemp = smAdfxDirectRefID
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAgencyRefId = slTemp

        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAgyName = Trim(tlExport_TranInfo.sAgyName)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sAirVehicle = Trim(tlExport_TranInfo.sBillVehicle)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sBusCat = Trim(tlExport_TranInfo.sBusCat)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sCashTrade = Trim(tlExport_TranInfo.sCashTrade)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sCheck = Trim(tlExport_TranInfo.sCheck)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sComm = Trim(tlExport_TranInfo.sComm)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sContract = Trim(tlExport_TranInfo.sContract)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sGross = Trim(tlExport_TranInfo.sGross)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sInvNo = Trim(tlExport_TranInfo.sInvNo)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sNet = Trim(tlExport_TranInfo.sNet)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sNTRType = Trim(tlExport_TranInfo.sNTRType)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sOffice = Trim(tlExport_TranInfo.sOffice)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sPolitical = Trim(tlExport_TranInfo.sPolitical)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sPostDate = Trim(tlExport_TranInfo.sPostDate)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sProduct = Trim(tlExport_TranInfo.sProduct)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sSalesSource = Trim(tlExport_TranInfo.sSalesSource)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sSlspStnCode = Trim(tlExport_TranInfo.sSlspStnCode)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sSlspName = Trim(tlExport_TranInfo.sSlspName)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sTranDate = Trim(tlExport_TranInfo.sTranDate)
        'tmExport_TranSummary(UBound(tmExport_TranSummary)).sTranType = Trim(tlExport_TranInfo.sTranType)
        ReDim Preserve tmExport_TranSummary(UBound(tmExport_TranSummary) + 1) As EXPORT_TRANSUMMARY
    End If
End Sub

'TTP 10519 - Invoice Register Export - add summary version setting to current export
Private Sub mUpdateTranSummary2(tlExport_TranInfo As EXPORT_TRANINFO)
    Dim ilLoop As Integer
    Dim blFound As Boolean
    Dim slTemp As String
    Dim ilMsgbox As Integer
    
    'Check tmInvDistSummary for existing stuff
    For ilLoop = 0 To UBound(tmExport_TranSummary)
        blFound = False
        'summary line total for each Order/Invoice
        'Per Jason Teams: I'm just noticing there's also the tran type column, which can be IN (invoice) or AN (invoice adjustment). I think we need to keep those separate, even if they have the same invoice number, for this summary version
        If Trim(tmExport_TranSummary(ilLoop).sInvNo) = Trim(tlExport_TranInfo.sInvNo) And _
            Trim(tmExport_TranSummary(ilLoop).sContract) = Trim(tlExport_TranInfo.sContract) And _
            Trim(tmExport_TranSummary(ilLoop).sTranType) = Trim(tlExport_TranInfo.sTranType) Then
            blFound = True
            Exit For
        End If
    Next ilLoop
    If blFound = True Then
        'Update the existing Invoice in the Array
        tmExport_TranSummary(ilLoop).sGross = Val(tmExport_TranSummary(ilLoop).sGross) + Val(Trim(tlExport_TranInfo.sGross))
        tmExport_TranSummary(ilLoop).sComm = Val(tmExport_TranSummary(ilLoop).sComm) + Val(Trim(tlExport_TranInfo.sComm))
        tmExport_TranSummary(ilLoop).sNet = Val(tmExport_TranSummary(ilLoop).sNet) + Val(Trim(tlExport_TranInfo.sNet))
        If Trim(tlExport_TranInfo.sNTRType) <> "" Then tmExport_TranSummary(ilLoop).sNTRType = Trim(Replace(tlExport_TranInfo.sNTRType, ",", ";"))
    Else
        'Add Order / Invoice to Array
        tmExport_TranSummary(UBound(tmExport_TranSummary)).iAdvCode = tlExport_TranInfo.iAdvCode
        tmExport_TranSummary(UBound(tmExport_TranSummary)).iAgyCode = tlExport_TranInfo.iAgyCode
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAccountID = Trim(tlExport_TranInfo.sAccountID)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAction = Trim(tlExport_TranInfo.sAction)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAdvertiserRefId = mGetAdfxRefID(tlExport_TranInfo.iAdvCode)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAdvName = Trim(Replace(tlExport_TranInfo.sAdvName, ",", ";"))
        slTemp = mGetAgfxRefID(tlExport_TranInfo.iAgyCode)
        If slTemp = "" Then slTemp = smAdfxDirectRefID
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAgencyRefId = slTemp
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAgyName = Trim(Replace(tlExport_TranInfo.sAgyName, ",", ";"))
        tmExport_TranSummary(UBound(tmExport_TranSummary)).iAirVehicleCode = Trim(tlExport_TranInfo.iAirVehicleCode)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sAirVehicle = Trim(tlExport_TranInfo.sAirVehicle)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sBusCat = Trim(tlExport_TranInfo.sBusCat)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sCashTrade = Trim(tlExport_TranInfo.sCashTrade)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sCheck = Trim(tlExport_TranInfo.sCheck)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sComm = Trim(tlExport_TranInfo.sComm)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sContract = Trim(tlExport_TranInfo.sContract)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sGross = Trim(tlExport_TranInfo.sGross)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sInvNo = Trim(tlExport_TranInfo.sInvNo)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sNet = Trim(tlExport_TranInfo.sNet)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sNTRType = Trim(Replace(tlExport_TranInfo.sNTRType, ",", ";"))
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sOffice = Trim(Replace(tlExport_TranInfo.sOffice, ",", ";"))
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sPolitical = Trim(tlExport_TranInfo.sPolitical)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sPostDate = Trim(tlExport_TranInfo.sPostDate)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sProduct = Trim(Replace(tlExport_TranInfo.sProduct, ",", ";"))
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sSalesSource = Trim(Replace(tlExport_TranInfo.sSalesSource, ",", ";"))
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sSlspStnCode = Trim(tlExport_TranInfo.sSlspStnCode)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).iSlspCode = Trim(tlExport_TranInfo.iSlspCode)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sSlspName = Trim(Replace(tlExport_TranInfo.sSlspName, ",", ";"))
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sTranDate = Trim(tlExport_TranInfo.sTranDate)
        tmExport_TranSummary(UBound(tmExport_TranSummary)).sTranType = Trim(tlExport_TranInfo.sTranType)
        ReDim Preserve tmExport_TranSummary(UBound(tmExport_TranSummary) + 1) As EXPORT_TRANSUMMARY
    End If
End Sub

Function mExportTranSummary() As Boolean
    Dim ilLoop As Integer
    Dim blFound As Boolean
    Dim slDelimiter As String
    Dim slTextQual As String
    Dim slRecord As String
    
    'TTP 10260 - JW - 7/28/21 - Support CSV
    If optFormat(1).Value = True Then
        slDelimiter = Chr(30)
        slTextQual = ""
    Else
        slDelimiter = ","
        slTextQual = """"
    End If
    
    lmRecordsExported = 0
    mExportTranSummary = False
    
    'Check tmInvDistSummary for existing stuff
    For ilLoop = 0 To UBound(tmExport_TranSummary) - 1
        blFound = False
        If imTerminate Then Exit Function
        slRecord = ""
        'TTP 10260 - JW - 7/28/21 - Support CSV
        slRecord = slRecord & slTextQual & Trim$(gStripChr0(tgSpfx.sInvExpProperty)) & slTextQual & slDelimiter                            'Property (sInvExpProperty)
        slRecord = slRecord & slTextQual & Trim$(gStripChr0(tmExport_TranSummary(ilLoop).sAdvName)) & slTextQual & slDelimiter             'Advertiser
        slRecord = slRecord & slTextQual & Trim$(gStripChr0(tmExport_TranSummary(ilLoop).sAdvertiserRefId)) & slTextQual & slDelimiter     'AdvertiserGUID
        slRecord = slRecord & slTextQual & Trim$(gStripChr0(tmExport_TranSummary(ilLoop).sAgyName)) & slTextQual & slDelimiter             'Agency
        slRecord = slRecord & slTextQual & Trim$(gStripChr0(tmExport_TranSummary(ilLoop).sAgencyRefId)) & slTextQual & slDelimiter         'AgencyGUID
        slRecord = slRecord & slTextQual & Trim$(gStripChr0(tmExport_TranSummary(ilLoop).sSlspName)) & slTextQual & slDelimiter            'AE Full Name
        slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sSlspStnCode) & slTextQual & slDelimiter                     'AECode
        slRecord = slRecord & slTextQual & Trim$(gStripChr0(tmExport_TranSummary(ilLoop).sOffice)) & slTextQual & slDelimiter              'Sales Office
        slRecord = slRecord & slTextQual & Trim$(gStripChr0(tmExport_TranSummary(ilLoop).sOffice)) & slTextQual & slDelimiter              'SalesOfficeCode (Will use the same value as the sales office field)
        If Trim$(tmExport_TranSummary(ilLoop).sAgyName) <> "" Then                                          'RevenueCode1 (
            slRecord = slRecord & slTextQual & "AGY" & slTextQual & slDelimiter                                     'If there is an agency defined for the contract/invoice, then 'AGY' will be shown in the RevenueCode1 field.
        Else
            slRecord = slRecord & slTextQual & "DIR" & slTextQual & slDelimiter                                     'If there's no agency (it's a direct advertiser), it will show 'DIR'
        End If                                                                                              ')
        slRecord = slRecord & slTextQual & "GEN" & slTextQual & slDelimiter                                         'RevenueCode2 (Will show 'GEN' for all records)
        slRecord = slRecord & slTextQual & "GEN" & slTextQual & slDelimiter                                         'RevenueCode3 (Will show 'GEN' for all records)
        slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sTranDate) & slTextQual & slDelimiter 'Date
        
        'TTP 10274 - 2 digit invoice month (example: August is 08, December is 12), followed by a 2 digit year (example: 2021 is 21), followed by the Prefix value from Site Options, followed by the Counterpoint order number.
        'slRecord = slRecord & Trim$(gStripChr0(tgSpfx.sInvExpPrefix)) & Trim$(tmExport_TranSummary(illoop).sContract) & slDelimiter    'Prefix + Order#
        slRecord = slRecord & slTextQual   '& Trim$(gStripChr0(tgSpfx.sInvExpPrefix)) & Trim$(tmExport_TranSummary(illoop).sContract) & slTextQual & slDelimiter    'Prefix + Order#
        slRecord = slRecord & right("0" & Month(DateValue(tmExport_TranSummary(ilLoop).sTranDate)), 2)
        slRecord = slRecord & right("0" & Year(DateValue(tmExport_TranSummary(ilLoop).sTranDate)), 2)
        slRecord = slRecord & Trim$(gStripChr0(tgSpfx.sInvExpPrefix))
        slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sContract)
        slRecord = slRecord & slTextQual
        slRecord = slRecord & slDelimiter
        
        slRecord = slRecord & Trim$(gStripChr0(tgSpfx.sInvExpPrefix)) & Trim$(tmExport_TranSummary(ilLoop).sInvNo) & slDelimiter       'Prefix + Invoice#
        slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sNet) & slDelimiter                                                   'Balance
        slRecord = slRecord & slTextQual & Trim$(gStripChr0(tgSpfx.sInvExpBillGroup)) & slTextQual     'BillingGroup (sInvExpBillGroup)
        
        'TTP 10260 - JW - 7/28/21 - Support CSV
        If optFormat(1).Value = True Then
            mWriteExcel slRecord, slDelimiter
        Else
            Print #hmCashInv, slRecord
        End If
        lmRecordsExported = lmRecordsExported + 1
        igDOE = igDOE + 1
        If igDOE >= 100 Then
            lacInfo(0).Caption = "Exporting " & lmRecordsExported & " records..."
            igDOE = 0
            DoEvents
        End If
    Next ilLoop
    gAutomationAlertAndLogHandler "Exported " & lmRecordsExported & " records."
    lacInfo(0).Caption = "Exported " & lmRecordsExported & " records."
    DoEvents
    mExportTranSummary = True
End Function

'TTP 10519 - Invoice Register Export - add summary version setting to current export
Function mExportTranSummary2() As Boolean
    Dim ilLoop As Integer
    Dim blFound As Boolean
    Dim slDelimiter As String
    Dim slTextQual As String
    Dim slRecord As String
    Dim sSubCompany As String
    Dim sSubtotals As String
    Dim sMarket  As String
    Dim sResearch As String
    Dim sFormat As String
    
    'If optFormat(1).Value = True Then
    '    slDelimiter = Chr(30)
    '    slTextQual = ""
    'Else
        slDelimiter = ","
        slTextQual = """"
    'End If
    
    lmRecordsExported = 0
    mExportTranSummary2 = False
    
    'Check tmInvDistSummary for existing stuff
    For ilLoop = 0 To UBound(tmExport_TranSummary) - 1
        blFound = False
        If imTerminate Then Exit Function
        slRecord = ""
                
        'TTP 10612 - Invoice Register Summary export: new 7 column option
        If optSummaryFormat1.Value = True Then
            '22-Column Format
            slRecord = slTextQual & Trim$(tmExport_TranSummary(ilLoop).sAccountID) & slTextQual & slDelimiter                   'Account ID
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sAgyName) & slTextQual & slDelimiter          'Agency
            slRecord = slRecord & slTextQual & mGetAgfxRefID(tmExport_TranSummary(ilLoop).iAgyCode) & slTextQual & slDelimiter  'Agency Reference ID
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sAdvName) & slTextQual & slDelimiter          'Advertiser
            slRecord = slRecord & slTextQual & mGetAdfxRefID(tmExport_TranSummary(ilLoop).iAdvCode) & slTextQual & slDelimiter  'Advertiser Reference ID
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sProduct) & slTextQual & slDelimiter          'Product
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sSlspName) & slTextQual & slDelimiter         'Salesperson
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).iSlspCode) & slDelimiter                                   'Salesperson Code
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sOffice) & slTextQual & slDelimiter           'Sales Office
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sSalesSource) & slTextQual & slDelimiter      'Sales Source
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sSlspStnCode) & slTextQual & slDelimiter      'Station Salesperson Code
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sBusCat) & slTextQual & slDelimiter           'Business Category
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sNTRType) & slTextQual & slDelimiter          'NTR Type
            'slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sBillVehicle) & slTextQual & slDelimiter      'Billing Vehicle
            'slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sAirVehicle) & slTextQual & slDelimiter       'Airing Vehicle
            'mGetAllVefGroupNames tmExport_TranSummary(ilLoop).iAirVehicleCode, sSubCompany, sSubtotals, sMarket, sResearch, sFormat
            'slRecord = slRecord & slTextQual & Trim$(sSubCompany) & slTextQual & slDelimiter                                    'Sub-company
            'slRecord = slRecord & slTextQual & Trim$(sSubtotals) & slTextQual & slDelimiter                                     'Subtotals
            'slRecord = slRecord & slTextQual & Trim$(sMarket) & slTextQual & slDelimiter                                        'Market
            'slRecord = slRecord & slTextQual & Trim$(sResearch) & slTextQual & slDelimiter                                      'Research
            'slRecord = slRecord & slTextQual & Trim$(sFormat) & slTextQual & slDelimiter                                        'Format
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sContract) & slDelimiter                                   'Contract #
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sTranDate) & slTextQual & slDelimiter         'Tran Date
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sInvNo) & slDelimiter                                      'Invoice #
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sTranType) & slTextQual & slDelimiter         'Tran Type
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sGross) & slDelimiter                                      'Gross Amt
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sComm) & slDelimiter                                       'Comm Amt
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sNet) & slDelimiter                                        'Net Amt
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sCashTrade) & slTextQual & slDelimiter        'Cash/Trade
            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sPolitical) & slTextQual                      'Political
        End If
        
        'TTP 10612 - Invoice Register Summary export: new 7 column option
        If optSummaryFormat2.Value = True Then
'            '7-Column Format
'            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sAgyName) & slTextQual & slDelimiter          'Agency
'            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sAdvName) & slTextQual & slDelimiter          'Advertiser
'            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sProduct) & slTextQual & slDelimiter          'Product
'            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sContract) & slDelimiter                                   'Contract #
'            slRecord = slRecord & slTextQual & Trim$(tmExport_TranSummary(ilLoop).sTranDate) & slTextQual & slDelimiter         'Tran Date
'            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sInvNo) & slDelimiter                                      'Invoice #
'            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sNet) & slDelimiter                                        'Net Amt
            
            '----------------------------------------------------
            'TTP 10626 - AURN Sage Intacct Export: "   No quotes around strings
            'INVOICE_NO: Counterpoint invoice number
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sInvNo) & slDelimiter
            'PO_NO: Counterpoint contract number
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sContract) & slDelimiter
            'CUSTOMER_ID:
            If Trim$(tmExport_TranSummary(ilLoop).sAgyName) <> "" Then
                slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sAgencyRefId) & slDelimiter          'For non-direct advertiser, use Agency Ref ID from Agfx table.
            Else
                slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sAdvertiserRefId) & slDelimiter      'For a direct advertiser, use advertiser Ref ID from Adfx
            End If
            'CREATED_DATE: use invoice date (tran date from invoice register export)
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sTranDate) & slDelimiter
            'TOTAL_DUE: net dollars due for invoice (same as "Net Amt" from invoice register summary export. Note: do not include commas in dollar figures or dollar sign
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sNet) & slDelimiter
            'TERM_NAME: it will show "Net 30". We will need to add a new field to SPFX and text entry field in Site Options
            slRecord = slRecord & Trim$(gStripChr0(tgSpfx.sSageTerm)) & slDelimiter
            'DESCRIPTION: concatenate advertiser name, hyphen, product name. Example: "Capital One - National Bank". Max length 1000 (this will never reach 1000 characters, sAdvName=30 + sProduct=40)
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sAdvName) & " - " & Trim$(tmExport_TranSummary(ilLoop).sProduct) & slDelimiter
            'LINE_NO: show number 1 for every row. The reason why is Sage Intacct supports multiline invoices and requires an invoice line number for every line item, but AURN will have a single row per invoice, and each line must be numbered for Sage, therefore each will be line 1.
            slRecord = slRecord & "1" & slDelimiter
            'ACCT_NO: will need to add a new Account Number field to SPFX and Site Options. Example: 1180.
            slRecord = slRecord & Trim$(gStripChr0(tgSpfx.sSageAccount)) & slDelimiter
            'LOCATION_ID: will need to add a new Location ID field to SPFX and Site Options. Example: 102
            slRecord = slRecord & Trim$(gStripChr0(tgSpfx.sSageLocation)) & slDelimiter
            'DEPT_ID: will need to add a new Department ID field to SPFX and Site Options. Example: D00
            slRecord = slRecord & Trim$(gStripChr0(tgSpfx.sSageDept)) & slDelimiter
            'AMOUNT: this will be the same value as "TOTAL_DUE". It's odd this is on it twice, but the Sage consultant says it's required.
            slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sNet) & slDelimiter
            'ARINVOICEITEM_ARACCOUNT: same value as ACCT_NO. It's odd this is on it twice, but the Sage consultant says it's required.
            slRecord = slRecord & Trim$(gStripChr0(tgSpfx.sSageAccount)) & slDelimiter
            'ARINVOICEITEM_CUSTOMERID: same as CUSTOMER_ID. It's odd this is on it twice, but the Sage consultant says it's required.
            If Trim$(tmExport_TranSummary(ilLoop).sAgyName) <> "" Then
                slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sAgencyRefId)      'For non-direct advertiser, use Agency Ref ID from Agfx table.
            Else
                slRecord = slRecord & Trim$(tmExport_TranSummary(ilLoop).sAdvertiserRefId)  'For a direct advertiser, use advertiser Ref ID from Adfx
            End If
        End If
        'If optFormat(1).Value = True Then
        '    mWriteExcel slRecord, slDelimiter
        'Else
            Print #hmCashInv, slRecord
        'End If
        lmRecordsExported = lmRecordsExported + 1
        igDOE = igDOE + 1
        If igDOE >= 100 Then
            lacInfo(0).Caption = "Exporting " & lmRecordsExported & " records..."
            igDOE = 0
            DoEvents
        End If
    Next ilLoop
    gAutomationAlertAndLogHandler "Exported " & lmRecordsExported & " records."
    lacInfo(0).Caption = "Exported " & lmRecordsExported & " records."
    DoEvents
    mExportTranSummary2 = True
End Function

Sub mGetExportFilename()
    Dim slRepeat As String
    Dim slClientName As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim ilTemp As Integer
    Dim slFullStartDate As String
    Dim slFullEndDate As String
    Dim ilRet As Integer
    
    lmStart = gDateValue(CSI_CalStart.Text)
    lmEnd = gDateValue(CSI_CalEnd.Text)
    smFullStartDate = Format(lmStart, "ddddd")
    smFullEndDate = Format(lmEnd, "ddddd")
    gObtainYearMonthDayStr CSI_CalStart.Text, True, slYear, slMonth, slDay
    smStart = Trim$(slMonth) & Trim$(slDay) & Mid(slYear, 3, 2)
    gObtainYearMonthDayStr CSI_CalEnd.Text, True, slYear, slMonth, slDay
    smEnd = Trim$(slMonth) & Trim$(slDay) & Mid(slYear, 3, 2)
        
    slClientName = Trim$(tgSpf.sGClient)
    If tgSpf.iMnfClientAbbr > 0 Then
        tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            slClientName = Trim$(tmMnf.sName)
        End If
    End If
    slRepeat = ""
    Do
        ilRet = 0
        'On Error GoTo cmcExportDupNameErr:
        Select Case imExportOption
            Case EXP_CASH
                smExportName = Trim$(sgExportPath) & smExportOptionName & " " & smStart & "-" & smEnd
                smExportName = Trim$(smExportName) & Trim$(slRepeat) & " " & gFileNameFilter(Trim$(slClientName)) & ".csv"
            Case Exp_INVREG
                If ckcSummary.Value = vbChecked Then
                    smExportName = Trim$(sgExportPath) & smExportOptionName & " Summary " & smStart & "-" & smEnd
                    smExportName = Trim$(smExportName) & Trim$(slRepeat) & " " & gFileNameFilter(Trim$(slClientName)) & ".csv"
                Else
                    smExportName = Trim$(sgExportPath) & smExportOptionName & " " & smStart & "-" & smEnd
                    smExportName = Trim$(smExportName) & Trim$(slRepeat) & " " & gFileNameFilter(Trim$(slClientName)) & ".csv"
                End If
            Case EXP_AUDACYINV
                smExportName = Trim$(sgExportPath) & "WOINV" & " " & smStart & "-" & smEnd
                'TTP 10260 - JW - 7/28/21 - Support CSV
                If optFormat(1).Value = True Then
                    smExportName = Trim$(smExportName) & Trim$(slRepeat) & " " & gFileNameFilter(Trim$(slClientName)) & ".xls"
                Else
                    smExportName = Trim$(smExportName) & Trim$(slRepeat) & " " & gFileNameFilter(Trim$(slClientName)) & ".csv"
                End If
        End Select
        ilRet = gFileExist(smExportName)
        If ilRet = 0 Then       'if went to mOpenErr , there was a filename that existed with same name. Increment the letter
            If slRepeat = "" Then
                slRepeat = "A"
            Else
                slRepeat = Chr(Asc(slRepeat) + 1)
            End If
        End If
    Loop While ilRet = 0
    edcTo.Text = smExportName
End Sub

Function mGetAdfxRefID(ilAdfID As Integer) As String
    Dim slSql As String
    Dim myRsQuery As ADODB.Recordset
    mGetAdfxRefID = ""
    If ilAdfID = 0 Then Exit Function
    If ilAdfID = imLastAdfCode Then
        mGetAdfxRefID = smLastAdfName
        Exit Function
    End If
    
    smAdfxDirectRefID = ""
    slSql = "select adfxRefId as Code , AdfxDirectRefID as DirectCode from ADFX_Advertisers where adfxCode = " & ilAdfID
    Set myRsQuery = gSQLSelectCall(slSql)
    If Not myRsQuery.EOF Then
        mGetAdfxRefID = myRsQuery!Code
        imLastAdfCode = ilAdfID
        smLastAdfName = myRsQuery!Code
        'JW - 7/29/21 - TTP 10261: WO Invoice Export - add direct advertiser Ref ID
        smAdfxDirectRefID = myRsQuery!DirectCode
    End If
End Function

Function mGetAgfxRefID(ilAgfID As Integer) As String
    Dim slSql As String
    Dim myRsQuery As ADODB.Recordset
    mGetAgfxRefID = ""
    If ilAgfID = 0 Then Exit Function
    If ilAgfID = imLastAgyCode Then
        mGetAgfxRefID = smLastAgyName
        Exit Function
    End If
    slSql = "select agfxRefId as Code from AGFX_Agencies where agfxCode = " & ilAgfID
    Set myRsQuery = gSQLSelectCall(slSql)
    If Not myRsQuery.EOF Then
        mGetAgfxRefID = myRsQuery!Code
        imLastAgyCode = ilAgfID
        smLastAgyName = myRsQuery!Code
        
    End If
End Function

Function mWriteExcel(slRecord As String, Optional slDelimiter As String = ",")
    Dim ilRet As Integer
    'Dim slDelimiter As String
    Dim ilColumn  As Integer
    Dim slExportFilename As String
    
    'slDelimiter = ","
    ilColumn = 1
    If imExcelRow = 0 Then imExcelRow = 1
    ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, imExcelRow, ilColumn, slDelimiter)
    imExcelRow = imExcelRow + 1
End Function

Function mDecorateExcel()
    Dim ilRet As Integer
    'Autofit
    ilRet = gExcelOutputGeneration("AF", omBook, omSheet, 1, , , -1)
End Function

Function mSaveExcel(slExcelFilename As String, Optional blView As Boolean = False)
    Dim ilRet As Integer
    
    'Save As XLS
    ilRet = gExcelOutputGeneration("S8", omBook, omSheet, 1, slExcelFilename)
    
    If blView = True Then
        'View Excel
        ilRet = gExcelOutputGeneration("V")
    End If
    
    ilRet = gExcelOutputGeneration("Q")
    Set ogExcel = Nothing
    Set omBook = Nothing
End Function

'TTP 10487 - Invoice Register export: add several new fields
Function mGetAllVefGroupNames(iVehCode As Integer, sSubCompany As String, sSubtotals As String, sMarket As String, sResearch As String, sFormat As String)
    Dim ilTemp As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    sSubCompany = ""
    sSubtotals = ""
    sMarket = ""
    sResearch = ""
    sFormat = ""
    '6/30/22 - Fix per Jason Email 6/28/22 3:13 PM
    'vefMnfVehGp2 = sub-totals (MnfUnitType 2)
    ilTemp = -1
    gGetVehGrpSets iVehCode, 2, 0, ilTemp, ilRet
    If ilTemp > 0 Then
        For ilLoop = LBound(tmVGMNF) To UBound(tmVGMNF) - 1
            If tmVGMNF(ilLoop).iCode = ilTemp Then
                sSubtotals = Trim(Mid(tmVGMNF(ilLoop).sName, 1, 20))
                Exit For
            End If
        Next ilLoop
    End If
    'vefMnfVehGp3Mkt = market (MnfUnitType 3)
    ilTemp = -1
    gGetVehGrpSets iVehCode, 3, 0, ilTemp, ilRet
    If ilTemp > 0 Then
        For ilLoop = LBound(tmVGMNF) To UBound(tmVGMNF) - 1
            If tmVGMNF(ilLoop).iCode = ilTemp Then
                sMarket = Trim(Mid(tmVGMNF(ilLoop).sName, 1, 20))
                Exit For
            End If
        Next ilLoop
    End If
    'vefMnfVehGp4Fmt = format (MnfUnitType 4)
    ilTemp = -1
    gGetVehGrpSets iVehCode, 4, 0, ilTemp, ilRet
    If ilTemp > 0 Then
        For ilLoop = LBound(tmVGMNF) To UBound(tmVGMNF) - 1
            If tmVGMNF(ilLoop).iCode = ilTemp Then
                sFormat = Trim(Mid(tmVGMNF(ilLoop).sName, 1, 20))
                Exit For
            End If
        Next ilLoop
    End If
    'vefMnfVehGp5Rsch = research (MnfUnitType 5)
    ilTemp = -1
    gGetVehGrpSets iVehCode, 5, 0, ilTemp, ilRet
    If ilTemp > 0 Then
        For ilLoop = LBound(tmVGMNF) To UBound(tmVGMNF) - 1
            If tmVGMNF(ilLoop).iCode = ilTemp Then
                sResearch = Trim(Mid(tmVGMNF(ilLoop).sName, 1, 20))
                Exit For
            End If
        Next ilLoop
    End If
    'vefMnfVehGp6Sub = sub-company (MnfUnitType 6)
    ilTemp = -1
    gGetVehGrpSets iVehCode, 6, 0, ilTemp, ilRet
    If ilTemp > 0 Then
        For ilLoop = LBound(tmVGMNF) To UBound(tmVGMNF) - 1
            If tmVGMNF(ilLoop).iCode = ilTemp Then
                sSubCompany = Trim(Mid(tmVGMNF(ilLoop).sName, 1, 20))
                Exit For
            End If
        Next ilLoop
    End If
End Function

